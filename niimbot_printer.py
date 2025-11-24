"""
NIIMBOT Printer Interface
Simplified interface for NIIMBOT B1 label printer using bleak for Bluetooth
Based on NiimPrintX library by labbots
"""

import asyncio
import struct
import math
from bleak import BleakClient, BleakScanner
from PIL import Image, ImageOps
import enum


class PrinterException(Exception):
    """Exception raised for printer-related errors"""
    pass


class RequestCodeEnum(enum.IntEnum):
    """Command codes for NIIMBOT printer"""
    GET_INFO = 64
    HEARTBEAT = 220
    SET_LABEL_TYPE = 35
    SET_LABEL_DENSITY = 33
    START_PRINT = 1
    END_PRINT = 243
    START_PAGE_PRINT = 3
    END_PAGE_PRINT = 227
    SET_DIMENSION = 19
    SET_QUANTITY = 21
    GET_PRINT_STATUS = 163
    ALLOW_PRINT_CLEAR = 32


class NiimbotPacket:
    """Packet structure for NIIMBOT printer communication"""
    def __init__(self, type_, data):
        self.type = type_
        self.data = data

    @classmethod
    def from_bytes(cls, pkt):
        assert pkt[:2] == b"\x55\x55"
        assert pkt[-2:] == b"\xaa\xaa"
        type_ = pkt[2]
        len_ = pkt[3]
        data = pkt[4 : 4 + len_]

        checksum = type_ ^ len_
        for i in data:
            checksum ^= i
        assert checksum == pkt[-3]

        return cls(type_, data)

    def to_bytes(self):
        checksum = self.type ^ len(self.data)
        for i in self.data:
            checksum ^= i
        return bytes(
            (0x55, 0x55, self.type, len(self.data), *self.data, checksum, 0xAA, 0xAA)
        )


class NiimbotPrinter:
    """Interface for NIIMBOT label printer"""
    
    # Supported printer models
    MODELS = {
        "b1": {"name": "B1", "max_width": 384},
        "b18": {"name": "B18", "max_width": 384},
        "b21": {"name": "B21", "max_width": 384},
        "d11": {"name": "D11", "max_width": 240},
        "d110": {"name": "D110", "max_width": 240},
    }
    
    # Label sizes (width x height in mm)
    LABEL_SIZES = {
        "b1": {
            "30mm x 15mm": (30, 15),
            "40mm x 12mm": (40, 12),
            "50mm x 14mm": (50, 14),
            "75mm x 12mm": (75, 12),
        },
        "d11": {
            "30mm x 14mm": (30, 14),
            "40mm x 12mm": (40, 12),
            "50mm x 14mm": (50, 14),
            "75mm x 12mm": (75, 12),
        },
    }
    
    def __init__(self, model="b1"):
        self.model = model.lower()
        if self.model not in self.MODELS:
            raise ValueError(f"Unsupported printer model: {model}")
        
        self.client = None
        self.char_uuid = None
        self.notification_event = asyncio.Event()
        self.notification_data = None
        self.connected = False
        self._buffer_cleared = False
        
    async def scan_for_printers(self, timeout=10):
        """Scan for available NIIMBOT printers"""
        print(f"Scanning for NIIMBOT printers (timeout: {timeout}s)...")
        devices = await BleakScanner.discover(timeout=timeout)
        
        niimbot_devices = []
        for device in devices:
            if device.name and ("niimbot" in device.name.lower() or "b1" in device.name.lower()):
                # Get RSSI if available (not all platforms support it)
                rssi = getattr(device, 'rssi', None)
                niimbot_devices.append({
                    "name": device.name,
                    "address": device.address,
                    "rssi": rssi
                })
        
        return niimbot_devices
    
    async def connect(self, address):
        """Connect to printer via Bluetooth"""
        try:
            print(f"Connecting to {address}...")
            self.client = BleakClient(address)
            await self.client.connect()
            
            if self.client.is_connected:
                await self._find_characteristics()
                await self._prime_printer()
                self.connected = True
                print(f"Successfully connected to printer")
                return True
            else:
                raise PrinterException("Failed to connect to printer")
                
        except Exception as e:
            print(f"Connection error: {e}")
            raise PrinterException(f"Cannot connect to printer: {str(e)}")
    
    async def disconnect(self):
        """Disconnect from printer"""
        if self.client and self.client.is_connected:
            await self.client.disconnect()
            self.connected = False
            print("Disconnected from printer")
    
    async def _find_characteristics(self):
        """Find the correct Bluetooth characteristic for communication"""
        for service in self.client.services:
            for char in service.characteristics:
                props = char.properties
                if 'read' in props and 'write-without-response' in props and 'notify' in props:
                    self.char_uuid = char.uuid
                    return
        
        if not self.char_uuid:
            raise PrinterException("Cannot find Bluetooth characteristics")
    
    def _notification_handler(self, sender, data):
        """Handle notifications from printer"""
        self.notification_data = data
        self.notification_event.set()
    
    async def _send_command(self, request_code, data, timeout=10):
        """Send command to printer and wait for response"""
        try:
            packet = NiimbotPacket(request_code, data)
            await self.client.start_notify(self.char_uuid, self._notification_handler)
            await self.client.write_gatt_char(self.char_uuid, packet.to_bytes(), response=False)
            
            await asyncio.wait_for(self.notification_event.wait(), timeout)
            response = NiimbotPacket.from_bytes(self.notification_data)
            await self.client.stop_notify(self.char_uuid)
            self.notification_event.clear()
            
            return response
        except asyncio.TimeoutError:
            print(f"Timeout occurred for request")
            raise PrinterException("Printer communication timeout")
        except Exception as e:
            print(f"Command error: {e}")
            raise PrinterException(f"Command failed: {str(e)}")
    
    async def _write_raw(self, data):
        """Write raw data to printer without waiting for response"""
        await self.client.write_gatt_char(self.char_uuid, data.to_bytes(), response=False)
    
    def _encode_image(self, image: Image):
        """Encode image for printing"""
        # Convert to monochrome (no inversion - black pixels = 0, white = 1 in output)
        img = image.convert("L").convert("1")
        
        for y in range(img.height):
            line_data = [img.getpixel((x, y)) for x in range(img.width)]
            # Build bit string: black pixel (0) -> '1', white pixel (255) -> '0'
            line_data = "".join("1" if pix == 0 else "0" for pix in line_data)
            line_data = int(line_data, 2).to_bytes(math.ceil(img.width / 8), "big")
            
            counts = (0, 0, 0)
            header = struct.pack(">H3BB", y, *counts, 1)
            pkt = NiimbotPacket(0x85, header + line_data)
            yield pkt
    
    async def print_image(self, image: Image, density: int = 3, quantity: int = 1):
        """Print an image on the label printer"""
        if not self.connected:
            raise PrinterException("Printer not connected")
        
        print(f"[PRINT] Buffer cleared flag: {self._buffer_cleared}")
        
        # Prep image before sending to ensure printer-compatible orientation and width
        processed_image = self._prepare_image(image)
        print(
            f"[PRINT] Starting print job (density={density}, quantity={quantity}, "
            f"original={image.width}x{image.height}, processed={processed_image.width}x{processed_image.height})..."
        )
        
        # Set parameters
        await self._send_command(RequestCodeEnum.SET_LABEL_DENSITY, bytes((density,)))
        await self._send_command(RequestCodeEnum.SET_LABEL_TYPE, bytes((1,)))
        await self._send_command(RequestCodeEnum.START_PRINT, b"\x01")
        await self._send_command(RequestCodeEnum.START_PAGE_PRINT, b"\x01")
        await self._send_command(
            RequestCodeEnum.SET_DIMENSION,
            struct.pack(">HH", processed_image.height, processed_image.width),
        )
        await self._send_command(RequestCodeEnum.SET_QUANTITY, struct.pack(">H", quantity))
        
        # Send image data
        print(f"Sending image data ({processed_image.width}x{processed_image.height} pixels)...")
        for pkt in self._encode_image(processed_image):
            await self._write_raw(pkt)
            await asyncio.sleep(0.01)
        
        # End page and print job
        while not await self.end_page_print():
            await asyncio.sleep(0.05)
        
        # Poll printer status before closing the job to prevent premature END_PRINT cancellation
        while True:
            status = await self.get_print_status()
            if status["page"] >= quantity:
                break
            await asyncio.sleep(0.1)

        await self._send_command(RequestCodeEnum.END_PRINT, b"\x01")
        print("Print job sent to printer")
    async def end_page_print(self):
        """Signal end of page data"""
        packet = await self._send_command(RequestCodeEnum.END_PAGE_PRINT, b"\x01")
        return bool(packet.data[0])

    async def allow_print_clear(self):
        """Clear printer state so cached jobs don't block the next run"""
        packet = await self._send_command(RequestCodeEnum.ALLOW_PRINT_CLEAR, b"\x01")
        return bool(packet.data[0])

    async def _prime_printer(self):
        if self._buffer_cleared:
            print("[PRIME] Buffer already cleared, skipping.")
            return
        print("[PRIME] Attempting to clear printer buffer...")
        try:
            result = await self.allow_print_clear()
            print(f"[PRIME] ALLOW_PRINT_CLEAR returned: {result}")
            await asyncio.sleep(0.2)
            print("[PRIME] Buffer clear complete.")
        except Exception as exc:
            print(f"[PRIME] Warning: unable to clear pending jobs: {exc}")
        finally:
            self._buffer_cleared = True

    async def get_print_status(self):
        """Read how many pages the printer has physically produced"""
        packet = await self._send_command(RequestCodeEnum.GET_PRINT_STATUS, b"\x01")
        data = packet.data

        if len(data) >= 4:
            page, progress1, progress2 = struct.unpack(">HBB", data[:4])
        elif len(data) >= 2:
            page = struct.unpack(">H", data[:2])[0]
            progress1 = data[2] if len(data) > 2 else 0
            progress2 = data[3] if len(data) > 3 else 0
        else:
            page = data[0] if data else 0
            progress1 = 0
            progress2 = 0

        return {"page": page, "progress1": progress1, "progress2": progress2}
    
    def _prepare_image(self, image: Image) -> Image:
        """Resize the image so it fits the printer without altering designer orientation."""
        if image is None:
            raise PrinterException("No label image supplied")

        img = image.copy()
        model_info = self.MODELS.get(self.model, {})
        max_width = model_info.get("max_width", 384)

        # Downscale if label exceeds printable width
        if img.width > max_width:
            scale = max_width / float(img.width)
            new_height = max(1, int(img.height * scale))
            print(
                f"Resizing label from {img.width}x{img.height} to {max_width}x{new_height} to fit printer"
            )
            img = img.resize((max_width, new_height), Image.LANCZOS)

        return img

    async def heartbeat(self):
        """Send heartbeat to keep connection alive"""
        try:
            packet = await self._send_command(RequestCodeEnum.HEARTBEAT, b"\x01")
            return {"connected": True, "closingstate": packet.data[0], "powerlevel": packet.data[1]}
        except Exception:
            return {"connected": False}
