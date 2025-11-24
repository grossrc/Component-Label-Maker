"""
Component Label Maker v2.0
Main application for generating and printing labels with QR codes for electronics components
Now with NIIMBOT B1 printer support!
"""

import tkinter as tk
from tkinter import ttk
import threading
import os
import asyncio
import time
from PIL import Image, ImageTk
import webbrowser

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None

from niimbot_printer import NiimbotPrinter
from label_designer import LabelDesigner


def load_env_file():
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if not os.path.exists(env_path):
        return

    with open(env_path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if key and value and key not in os.environ:
                os.environ[key] = value


load_env_file()


class DigiKeyClient:
    """Minimal client for retrieving DigiKey part metadata."""

    TOKEN_URL = "https://api.digikey.com/v1/oauth2/token"
    PRODUCT_DETAILS_URL = "https://api.digikey.com/products/v4/search/{partnumber}/productdetails"

    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.access_token = None
        self.token_expiry = 0
        self.session = requests.Session() if requests else None

    def is_configured(self):
        return bool(
            self.session
            and self.client_id
            and self.client_secret
        )

    def _ensure_access_token(self):
        if not self.is_configured():
            return None
        if self.access_token and time.time() < (self.token_expiry - 30):
            return self.access_token

        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
        }
        response = self.session.post(self.TOKEN_URL, data=data, timeout=20)
        response.raise_for_status()
        payload = response.json()
        self.access_token = payload.get("access_token")
        self.token_expiry = time.time() + payload.get("expires_in", 1800)
        return self.access_token

    def fetch_part(self, part_number):
        token = self._ensure_access_token()
        if not token:
            print(f"[API] No access token available for {part_number}")
            raise Exception("DigiKey API authentication failed - no access token available")

        headers = {
            "Authorization": f"Bearer {token}",
            "X-DIGIKEY-Client-Id": self.client_id,
            "Accept": "application/json",
        }
        
        url = self.PRODUCT_DETAILS_URL.format(partnumber=part_number)
        print(f"[API] Requesting part data for: {part_number}")
        print(f"[API] URL: {url}")
        
        response = self.session.get(url, headers=headers, timeout=20)
        response.raise_for_status()
        data = response.json()
        
        print(f"[API] Response keys: {list(data.keys())}")
        
        product = data.get("Product")
        if not product:
            print("[API] No Product found in response")
            raise Exception(f"DigiKey API returned no product for part number: {part_number}")
        
        print(f"[API] Product keys: {list(product.keys())}")
        
        # The Description field is a dictionary containing ProductDescription and DetailedDescription
        description_obj = product.get("Description")
        if not description_obj or not isinstance(description_obj, dict):
            print("[API] ERROR: Description field not found or is not a dictionary!")
            raise Exception(f"DigiKey API response missing 'Description' field for {part_number}")
        
        print(f"[API] Description object keys: {list(description_obj.keys())}")
        
        detailed_desc = description_obj.get("DetailedDescription", "").strip()
        product_desc = description_obj.get("ProductDescription", "").strip()
        
        if not detailed_desc:
            print("[API] WARNING: DetailedDescription is empty, using ProductDescription")
            detailed_desc = product_desc
        
        if not detailed_desc:
            print("[API] ERROR: Both description fields are empty!")
            raise Exception(f"DigiKey API returned empty description for {part_number}")
        
        print(f"[API] SUCCESS: DetailedDescription = '{detailed_desc}'")
        print(f"[API] SUCCESS: ProductDescription = '{product_desc}'")
        return product


class LabelMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Component Label Maker v2.0")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Custom styles
        style.configure('Blue.TButton', font=('Segoe UI', 10, 'bold'), background='#2196F3', foreground='white')
        style.map('Blue.TButton', background=[('active', '#1976D2'), ('disabled', '#cccccc')])
        self.bg_color = style.lookup('TFrame', 'background') or style.lookup('TLabel', 'background') or self.root.cget('bg')
        
        # API credentials from environment variables
        self.client_id = os.getenv("DIGIKEY_CLIENT_ID")
        self.client_secret = os.getenv("DIGIKEY_CLIENT_SECRET")
        self.digikey_client = DigiKeyClient(self.client_id, self.client_secret)
        
        # Printer setup
        self.printer = None
        self.printer_model = "B1"  # Default to B1
        self.available_printers = []
        self.async_loop = None
        self.async_thread = None
        
        # Label preview
        self.current_label_image = None
        self.preview_photo = None
        
        # UI State
        self.is_busy = False
        
        # Start async loop for printer communication
        self.setup_async_loop()
        
        self.setup_ui()
        
    def setup_async_loop(self):
        """Setup asyncio event loop for printer communication"""
        self.async_loop = asyncio.new_event_loop()
        self.async_thread = threading.Thread(target=self._run_async_loop, daemon=True)
        self.async_thread.start()
        
    def _run_async_loop(self):
        """Run asyncio event loop in separate thread"""
        asyncio.set_event_loop(self.async_loop)
        self.async_loop.run_forever()
        
    def run_async(self, coro):
        """Run coroutine in async loop"""
        return asyncio.run_coroutine_threadsafe(coro, self.async_loop)
        
    def setup_ui(self):
        """Create the user interface"""
        # 1. Status Bar at the bottom (Always visible)
        status_frame = ttk.Frame(self.root, relief="sunken", padding="2")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(status_frame, text="Ready", font=('Segoe UI', 9))
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        # 2. Main Content Area
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # --- Header ---
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(header_frame, text="Component Label Maker", font=('Segoe UI', 18, 'bold')).pack()
        
        # Description (Compact)
        desc_text = tk.Text(header_frame, height=6, width=100, font=('Segoe UI', 10), 
                           bg=self.bg_color, relief='flat', wrap=tk.WORD)
        desc_text.pack(pady=2)
        
        full_text = (
            "This program was written to supplement the DigiKey Organizer Project, and the documentation for this program specifically can be found here. "
            "It takes the part number and quantity as inputs and embeds the information into a QR code resembling the DataMatrix printed on DigiKey parts. "
            "The part number, quantity, and description are also printed to allow a quick user-readable identifier.\n\n\n\n"
        )
        
        desc_text.insert(tk.END, full_text)
        
        # Add links
        desc_text.tag_config("link", foreground="blue", underline=1)
        desc_text.tag_bind("link", "<Enter>", lambda e: desc_text.config(cursor="hand2"))
        desc_text.tag_bind("link", "<Leave>", lambda e: desc_text.config(cursor=""))
        
        # Find and tag "DigiKey Organizer Project"
        start_idx = full_text.find("DigiKey Organizer Project")
        if start_idx != -1:
            end_idx = start_idx + len("DigiKey Organizer Project")
            desc_text.tag_add("link", f"1.0 + {start_idx} chars", f"1.0 + {end_idx} chars")
            desc_text.tag_bind("link", "<Button-1>", lambda e: webbrowser.open("https://github.com/grossrc/DigiKey_Organizer"))
            
        # Find and tag "here"
        start_idx = full_text.find("here")
        if start_idx != -1:
            end_idx = start_idx + len("here")
            desc_text.tag_add("link_docs", f"1.0 + {start_idx} chars", f"1.0 + {end_idx} chars")
            desc_text.tag_config("link_docs", foreground="blue", underline=1)
            desc_text.tag_bind("link_docs", "<Enter>", lambda e: desc_text.config(cursor="hand2"))
            desc_text.tag_bind("link_docs", "<Leave>", lambda e: desc_text.config(cursor=""))
            desc_text.tag_bind("link_docs", "<Button-1>", lambda e: webbrowser.open("https://github.com/grossrc/Component-Label-Maker"))
            
        desc_text.config(state='disabled')
        
        # --- Controls Container ---
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=5)
        
        # Part Info
        info_frame = ttk.LabelFrame(controls_frame, text="Part Information", padding="5")
        info_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(info_frame, text="DigiKey Part Number:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        self.part_number_entry = ttk.Entry(info_frame, width=25, font=('Segoe UI', 10))
        self.part_number_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(info_frame, text="Quantity:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        self.quantity_entry = ttk.Entry(info_frame, width=5, font=('Segoe UI', 10))
        self.quantity_entry.pack(side=tk.LEFT, padx=5)
        self.quantity_entry.insert(0, "1")
        
        ttk.Button(info_frame, text="✕ Clear", width=8, command=self.clear_part_number).pack(side=tk.LEFT, padx=5)
        
        # Printer Settings
        printer_frame = ttk.LabelFrame(controls_frame, text="Printer Settings", padding="5")
        printer_frame.pack(fill=tk.X, pady=2)
        
        # Row 1: Model & Size
        pf_row1 = ttk.Frame(printer_frame)
        pf_row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(pf_row1, text="Printer Model:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        self.model_var = tk.StringVar(value="B1")
        ttk.Combobox(pf_row1, textvariable=self.model_var, values=["B1", "B18", "B21", "D11", "D110"], 
                     state="readonly", width=10).pack(side=tk.LEFT, padx=5)
                     
        ttk.Label(pf_row1, text="Label Size:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=15)
        self.label_size_var = tk.StringVar(value="50mm x 30mm")
        ttk.Combobox(pf_row1, textvariable=self.label_size_var, values=list(LabelDesigner.LABEL_SIZES.keys()),
                     state="readonly", width=15).pack(side=tk.LEFT, padx=5)
        
        # Row 2: Scan & Connect
        pf_row2 = ttk.Frame(printer_frame)
        pf_row2.pack(fill=tk.X, pady=2)
        
        self.scan_button = ttk.Button(pf_row2, text="Scan for Printers", command=self.scan_printers)
        self.scan_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(pf_row2, text="Select Printer:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        self.printer_var = tk.StringVar(value="No printers found")
        self.printer_combo = ttk.Combobox(pf_row2, textvariable=self.printer_var, state="readonly", width=25)
        self.printer_combo.pack(side=tk.LEFT, padx=5)
        
        self.connect_button = ttk.Button(pf_row2, text="Connect", command=self.toggle_connection)
        self.connect_button.pack(side=tk.LEFT, padx=5)
        
        self.connection_status = ttk.Label(pf_row2, text="⚫ Disconnected", font=('Segoe UI', 9), foreground='red')
        self.connection_status.pack(side=tk.LEFT, padx=10)
        
        # --- Action Buttons ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        btn_inner = ttk.Frame(btn_frame)
        btn_inner.pack(anchor=tk.CENTER)
        
        self.generate_button = ttk.Button(btn_inner, text="Generate Label Preview", command=self.generate_label)
        self.generate_button.pack(side=tk.LEFT, padx=10)
        
        self.print_button = ttk.Button(btn_inner, text="✓ Print Label", command=self.print_current_label, 
                                      state='disabled', style='Blue.TButton')
        self.print_button.pack(side=tk.LEFT, padx=10)
        
        # --- Preview Section ---
        preview_frame = ttk.LabelFrame(main_frame, text="Label Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=False, pady=5)
        
        self.preview_info_label = ttk.Label(preview_frame, text="No preview generated yet", font=('Segoe UI', 9), foreground='gray')
        self.preview_info_label.pack()
        
        self.preview_container = ttk.Frame(preview_frame, height=220)
        self.preview_container.pack_propagate(False)
        self.preview_container.pack(fill=tk.X, expand=False, pady=5)
        
        self.preview_label = ttk.Label(self.preview_container, text="")
        self.preview_label.pack(expand=True)
        
        self.preview_note = ttk.Label(preview_frame, text="Tip: Generate a label preview to see how it will look before printing.", 
                                     font=('Segoe UI', 8), foreground='gray')
        self.preview_note.pack(side=tk.BOTTOM)

        # Output Label
        self.output_label = ttk.Label(main_frame, text="", font=('Segoe UI', 9, 'bold'), foreground='blue')
        self.output_label.pack(side=tk.BOTTOM, pady=5)
        
    def update_progress(self, value, status_text):
        """Update progress bar and status label"""
        self.progress_bar['value'] = value
        self.status_label.config(text=status_text)
        self.root.update_idletasks()

    def notify_user(self, message, level="info"):
        """Display inline status feedback instead of pop-up dialogs."""
        colors = {
            "info": "#1f77b4",
            "success": "#2ca02c",
            "warning": "#d68910",
            "error": "#c0392b",
        }

        def _update():
            self.output_label.config(text=message, foreground=colors.get(level, "#1f77b4"))

        self.root.after(0, _update)
        
    def scan_printers(self):
        """Scan for available NIIMBOT printers"""
        self.set_loading_state(True, "Scanning for NIIMBOT printers...")
        
        def scan_thread():
            try:
                printer = NiimbotPrinter(self.model_var.get())
                future = self.run_async(printer.scan_for_printers(timeout=10))
                devices = future.result(timeout=15)
                
                self.available_printers = devices
                
                if devices:
                    device_names = [f"{d['name']} ({d['address']})" for d in devices]
                    self.root.after(0, lambda: self.printer_combo.config(values=device_names))
                    self.root.after(0, lambda: self.printer_var.set(device_names[0]))
                    self.root.after(0, lambda: self.notify_user(f"Found {len(devices)} NIIMBOT printer(s).", "success"))
                else:
                    self.root.after(
                        0,
                        lambda: self.notify_user(
                            "No NIIMBOT printers found. Ensure the printer is powered on and Bluetooth is enabled.",
                            "warning",
                        ),
                    )
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda msg=error_msg: self.notify_user(f"Error scanning for printers: {msg}", "error"))
            finally:
                self.root.after(0, lambda: self.set_loading_state(False, "Scan complete"))
        
        thread = threading.Thread(target=scan_thread, daemon=True)
        thread.start()
        
    def toggle_connection(self):
        """Connect or disconnect from printer"""
        if self.printer and self.printer.connected:
            # Disconnect
            self.disconnect_printer()
        else:
            # Connect
            self.connect_printer()
            
    def connect_printer(self):
        """Connect to selected printer"""
        if not self.available_printers:
            self.notify_user("Please scan for printers before connecting.", "warning")
            return
            
        selected_idx = self.printer_combo.current()
        if selected_idx < 0:
            self.notify_user("Select a printer from the dropdown before connecting.", "warning")
            return
            
        self.set_loading_state(True, "Connecting to printer...")
        
        def connect_thread():
            try:
                device = self.available_printers[selected_idx]
                self.printer = NiimbotPrinter(self.model_var.get())
                
                future = self.run_async(self.printer.connect(device['address']))
                success = future.result(timeout=30)
                
                if success:
                    self.root.after(0, lambda: self.connection_status.config(text="🟢 Connected", foreground='green'))
                    self.root.after(0, lambda: self.notify_user(f"Connected to {device['name']}", "success"))
            except Exception as e:
                error_msg = str(e)
                self.printer = None
                self.root.after(0, lambda msg=error_msg: self.notify_user(f"Failed to connect: {msg}", "error"))
            finally:
                self.root.after(0, lambda: self.set_loading_state(False, "Connection complete" if self.printer and self.printer.connected else "Connection failed"))
        
        thread = threading.Thread(target=connect_thread, daemon=True)
        thread.start()
        
    def disconnect_printer(self):
        """Disconnect from printer"""
        if not self.printer:
            return
            
        self.set_loading_state(True, "Disconnecting...")
        
        def disconnect_thread():
            try:
                future = self.run_async(self.printer.disconnect())
                future.result(timeout=10)
                
                self.printer = None
                self.root.after(0, lambda: self.connection_status.config(text="⚫ Disconnected", foreground='red'))
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda msg=error_msg: self.notify_user(f"Error disconnecting: {msg}", "error"))
            finally:
                self.root.after(0, lambda: self.set_loading_state(False, "Disconnected"))
        
        thread = threading.Thread(target=disconnect_thread, daemon=True)
        thread.start()
        
    def generate_label(self):
        """Generate label preview"""
        # Validate inputs
        part_number = self.part_number_entry.get().strip()
        quantity = self.quantity_entry.get().strip()
        
        if not part_number:
            self.notify_user("Please enter a DigiKey part number.", "error")
            return
            
        if not quantity or not quantity.isdigit():
            self.notify_user("Quantity must be a positive integer.", "error")
            return
        
        self.set_loading_state(True, "Generating label...")
        self.notify_user("", "info")
        
        # Run in separate thread
        thread = threading.Thread(target=self.process_label, 
                                 args=(part_number, quantity))
        thread.daemon = True
        thread.start()
        
    def process_label(self, part_number, quantity):
        """Process the label generation (runs in separate thread)"""
        try:
            # Step 1: Fetch part details
            self.root.after(0, lambda: self.status_label.config(text="Fetching part details from DigiKey..."))
            part_details = self.fetch_part_details(part_number)
            detailed_description = part_details.get("detailed_description") or f"Component: {part_number}"
            display_part_number = part_details.get("display_part_number", part_number)
            
            # Step 2: Generate label design
            self.root.after(0, lambda: self.status_label.config(text="Designing label..."))
            designer = LabelDesigner(self.label_size_var.get())
            label_image = designer.create_label(display_part_number, quantity, detailed_description)
            
            self.current_label_image = label_image
            
            # Step 3: Show preview
            self.root.after(0, lambda: self.show_preview(label_image, display_part_number, quantity, detailed_description))
            
        except Exception as e:
            error_msg = str(e)
            
            # Check for API/NoneType errors and add helpful hint
            if "NoneType" in error_msg or "DigiKey API" in error_msg:
                error_msg += "\n\nTip: The DigiKey API call failed. Try using the DigiKey ascribed part number (e.g., 'P12345-ND') instead of the manufacturer part number."
            
            self.root.after(0, lambda msg=error_msg: self.notify_user(f"An error occurred: {msg}", "error"))
        finally:
            self.root.after(0, lambda: self.set_loading_state(False, "Preview generated" if self.current_label_image else "Generation failed"))
    
    def show_preview(self, label_image, part_number, quantity, description):
        """Show label preview in embedded UI"""
        # Update info label
        info_text = f"Part: {part_number} | Quantity: {quantity} | Size: {self.label_size_var.get()}"
        self.preview_info_label.config(text=info_text, foreground='black')
        
        # Calculate scale to fit within preview area (max 850px wide, 210px tall)
        max_width = 850
        max_height = 210
        
        # Calculate scale factor to fit within bounds while maintaining aspect ratio
        width_scale = max_width / label_image.width
        height_scale = max_height / label_image.height
        scale = min(width_scale, height_scale)
        
        # Calculate new dimensions
        new_width = int(label_image.width * scale)
        new_height = int(label_image.height * scale)
        
        # Resize with high-quality resampling for better preview
        display_image = label_image.resize((new_width, new_height), Image.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(display_image)
        
        # Update preview image
        self.preview_label.config(image=self.preview_photo)
        self.preview_label.image = self.preview_photo  # Keep reference
        
        # Update note with actual scale used
        scale_percent = int(scale * 100)
        self.preview_note.config(
            text=f"Preview scaled to {scale_percent}% for better layout view. Actual label will be printed at full resolution.",
            foreground='gray'
        )
        
        # Enable print button
        self.print_button.config(state='normal')
        
        # Notify success
        self.notify_user("Label preview generated successfully. Click 'Print Label' to print.", "success")
    
    def print_current_label(self):
        """Print the current label"""
        if not self.printer or not self.printer.connected:
            self.notify_user("Connect to a printer before printing.", "error")
            return
        
        if not self.current_label_image:
            self.notify_user("Generate a label preview before printing.", "error")
            return

        self.set_loading_state(True, "Preparing to print...")

        model_info = NiimbotPrinter.MODELS.get(self.model_var.get().lower())
        if model_info:
            max_width = model_info.get("max_width")
            if max_width and self.current_label_image.width > max_width:
                print(
                    f"Label width {self.current_label_image.width}px exceeds {self.model_var.get().upper()} limit; auto-scaling before print."
                )
        self.notify_user("", "info")
        
        def print_thread():
            try:
                self.root.after(0, lambda: self.status_label.config(text="Sending to printer..."))
                
                # Print with default density of 3
                future = self.run_async(
                    self.printer.print_image(
                        self.current_label_image,
                        density=3,
                        quantity=1,
                    )
                )
                
                self.root.after(0, lambda: self.status_label.config(text="Printing in progress..."))
                
                # Wait for print to actually complete with longer timeout
                result = future.result(timeout=120)
                
                self.root.after(0, lambda: self.status_label.config(text="Finalizing..."))
                
                # Give printer a moment to finish
                time.sleep(2)
                
                self.root.after(0, lambda: self.notify_user("Label printed successfully!", "success"))
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda msg=error_msg: self.notify_user(f"Failed to print: {msg}", "error"))
            finally:
                self.root.after(0, lambda: self.set_loading_state(False, "Print complete"))
        
        thread = threading.Thread(target=print_thread, daemon=True)
        thread.start()
    
    def fetch_part_details(self, part_number):
        """Fetch part details from DigiKey API with strict validation"""
        part_info = {
            "display_part_number": part_number,
            "detailed_description": None,
        }

        if not self.digikey_client or not self.digikey_client.is_configured():
            raise Exception("DigiKey API is not configured. Please check .env file for CLIENT_ID and CLIENT_SECRET")

        # Strict mode: API must succeed
        product = self.digikey_client.fetch_part(part_number)
        
        # Extract part number (prefer manufacturer part number)
        part_info["display_part_number"] = (
            product.get("ManufacturerProductNumber")
            or part_number
        )
        
        # Extract DetailedDescription from nested Description object
        description_obj = product.get("Description", {})
        if isinstance(description_obj, dict):
            detailed_desc = (description_obj.get("DetailedDescription") or "").strip()
            product_desc = (description_obj.get("ProductDescription") or "").strip()
            # Prefer detailed description, fall back to product description
            part_info["detailed_description"] = detailed_desc or product_desc
        else:
            part_info["detailed_description"] = str(description_obj).strip()
        
        # Optional fields
        manufacturer = product.get("Manufacturer") or {}
        if isinstance(manufacturer, dict):
            part_info["manufacturer"] = manufacturer.get("Value") or manufacturer.get("Name")
        else:
            part_info["manufacturer"] = str(manufacturer)
            
        part_info["product_url"] = product.get("ProductUrl")
        
        print(f"[API] Final extracted description: '{part_info['detailed_description']}'")
        print(f"[API] Final part number: '{part_info['display_part_number']}'")
        return part_info
    
    def on_closing(self):
        """Handle window closing"""
        if self.printer and self.printer.connected:
            try:
                future = self.run_async(self.printer.disconnect())
                future.result(timeout=5)
            except:
                pass
        
        if self.async_loop:
            self.async_loop.call_soon_threadsafe(self.async_loop.stop)
        
        self.root.destroy()
        
    def set_loading_state(self, is_loading, message=""):
        """Set the UI to a loading state"""
        self.is_busy = is_loading
        
        if is_loading:
            self.root.config(cursor="wait")
            self.progress_bar.config(mode='indeterminate')
            self.progress_bar.start(10)
            self.status_label.config(text=message)
            
            # Disable controls
            self.scan_button.config(state='disabled')
            self.connect_button.config(state='disabled')
            self.generate_button.config(state='disabled')
            self.print_button.config(state='disabled')
            self.part_number_entry.config(state='disabled')
            self.quantity_entry.config(state='disabled')
            
        else:
            self.root.config(cursor="")
            self.progress_bar.stop()
            self.progress_bar.config(mode='determinate', value=0)
            self.status_label.config(text=message if message else "Ready")
            
            # Enable controls
            self.scan_button.config(state='normal')
            self.connect_button.config(state='normal')
            self.generate_button.config(state='normal')
            self.part_number_entry.config(state='normal')
            self.quantity_entry.config(state='normal')
            
            # Only enable print if we have a preview and printer is connected
            if self.current_label_image and self.printer and self.printer.connected:
                self.print_button.config(state='normal')
            else:
                self.print_button.config(state='disabled')
                
            # Update connect button text based on state
            if self.printer and self.printer.connected:
                self.connect_button.config(text="Disconnect")
            else:
                self.connect_button.config(text="Connect")

    def clear_part_number(self):
        """Clear the part number entry field and reset quantity"""
        if not self.is_busy:
            self.part_number_entry.delete(0, tk.END)
            self.quantity_entry.delete(0, tk.END)
            self.quantity_entry.insert(0, "1")
            self.part_number_entry.focus()


def main():
    """Main entry point"""
    root = tk.Tk()
    app = LabelMakerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
