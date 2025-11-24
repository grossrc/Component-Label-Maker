"""
Label Designer
Creates label images with QR code and text for printing
"""

from PIL import Image, ImageDraw, ImageFont
import qrcode


class LabelDesigner:
    """Design and generate label images for printing"""

    DPI = 203
    RENDER_SCALE = 2
    LABEL_SIZES = {
        "30mm x 15mm": (236, 118),
        "40mm x 12mm": (315, 94),
        "50mm x 14mm": (394, 110),
        "75mm x 12mm": (591, 94),
        "50mm x 30mm": (394, 236),
    }

    def __init__(self, label_size="50mm x 14mm"):
        if label_size not in self.LABEL_SIZES:
            raise ValueError(f"Unsupported label size: {label_size}")

        self.label_size = label_size
        self.width, self.height = self.LABEL_SIZES[label_size]
        self.scale = self.RENDER_SCALE

    def mm_to_pixels(self, mm):
        inches = mm / 25.4
        return int(inches * self.DPI)

    def create_qr_code(self, part_number, quantity):
        qr_content = f"[)><RS>06<GS>1P{part_number}<GS>Q{quantity}<RS><EOT>"
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=2,
        )
        qr.add_data(qr_content)
        qr.make(fit=True)
        return qr.make_image(fill_color="black", back_color="white")

    def get_font(self, size, bold=False, italic=False):
        try:
            if bold and italic:
                font_path = "C:/Windows/Fonts/arialbi.ttf"
            elif bold:
                font_path = "C:/Windows/Fonts/arialbd.ttf"
            elif italic:
                font_path = "C:/Windows/Fonts/ariali.ttf"
            else:
                font_path = "C:/Windows/Fonts/arial.ttf"
            return ImageFont.truetype(font_path, size)
        except Exception:
            return ImageFont.load_default()

    def create_label(self, part_number, quantity, detailed_description):
        scale = self.scale
        canvas = Image.new("RGB", (self.width * scale, self.height * scale), "white")
        draw = ImageDraw.Draw(canvas)

        qr_img = self.create_qr_code(part_number, quantity)
        qr_size = self.height * scale - self._scale(8)
        qr_img = qr_img.resize((qr_size, qr_size), Image.LANCZOS)
        qr_x = self._scale(4)
        qr_y = (self.height * scale - qr_size) // 2
        canvas.paste(qr_img, (qr_x, qr_y))

        text_x = qr_x + qr_size + self._scale(10)
        text_width = self.width * scale - text_x - self._scale(10)
        available_text_height = self.height * scale - self._scale(8)  # Total available height with margins
        
        print(f"[LABEL] Canvas: {self.width * scale}x{self.height * scale}, Text area: {text_width}x{available_text_height}px")

        # Calculate 2mm in pixels at current DPI (203 DPI) and scale
        # 2mm = 0.0787 inches, at 203 DPI = ~16 pixels, scaled by 2 = 32 pixels
        max_text_height_px = int(self.mm_to_pixels(3) * scale)
        
        # Initial font sizes (these are the defaults we want to use if they fit)
        initial_part_size = min(self._scale(44), max_text_height_px)
        initial_qty_size = min(self._scale(20), max_text_height_px)
        initial_desc_size = self._scale(18)
        
        # Minimum font sizes before truncating
        min_part_size = self._scale(14)
        min_qty_size = self._scale(12)
        min_desc_size = self._scale(10)
        
        part_text = str(part_number)
        qty_text = f"Qty: {quantity}"
        
        # First, fit the part number to the available width
        # This is independent of height considerations
        part_font = self._fit_font(
            draw,
            part_text,
            text_width,
            start_size=initial_part_size,
            min_size=min_part_size,
            bold=True,
        )
        actual_part_size = getattr(part_font, 'size', initial_part_size)
        print(f"[LABEL] Part number '{part_text}' fitted to size: {actual_part_size}")
        
        # Try to fit with initial sizes, reducing if necessary
        line_spacing = self._scale(6)
        attempt = 0
        max_attempts = 20
        
        while attempt < max_attempts:
            attempt += 1
            
            # Calculate current font sizes for this attempt
            # Part number size is already determined by width fitting
            if attempt == 1:
                part_font_size = actual_part_size
                qty_font_size = initial_qty_size
                desc_font_size = initial_desc_size
            else:
                # Part number stays at its width-fitted size
                # Only reduce qty and desc if needed for height
                reduction = self._scale(2) * (attempt - 1)
                part_font_size = actual_part_size  # Don't reduce part number further
                qty_font_size = max(min_qty_size, initial_qty_size - reduction)
                desc_font_size = max(min_desc_size, initial_desc_size - reduction)
            
            # Create fonts (part_font already created from width fitting)
            qty_font = self.get_font(qty_font_size, bold=True)
            desc_font = self.get_font(desc_font_size, italic=True)
            
            # Calculate heights
            part_bbox = draw.textbbox((0, 0), part_text, font=part_font)
            part_height = part_bbox[3] - part_bbox[1]
            
            qty_bbox = draw.textbbox((0, 0), qty_text, font=qty_font)
            qty_height = qty_bbox[3] - qty_bbox[1]
            
            desc_bbox = draw.textbbox((0, 0), "Test", font=desc_font)
            desc_line_height = desc_bbox[3] - desc_bbox[1]
            
            # Calculate space for description
            used_height = part_height + line_spacing + qty_height + line_spacing
            desc_available_height = available_text_height - used_height
            
            # Calculate max lines
            if desc_available_height > 0:
                max_desc_lines = max(1, int(desc_available_height / (desc_line_height + line_spacing)))
            else:
                max_desc_lines = 0
            
            # Wrap description
            desc_lines = []
            if detailed_description and max_desc_lines > 0:
                desc_lines = self._wrap_text(draw, detailed_description.strip(), desc_font, text_width, max_lines=max_desc_lines)
            
            # Calculate total height of all content
            total_content_height = part_height + line_spacing + qty_height
            if desc_lines:
                total_content_height += line_spacing + (len(desc_lines) * desc_line_height) + ((len(desc_lines) - 1) * line_spacing)
            
            # Check if it fits
            if total_content_height <= available_text_height:
                # It fits! Use these sizes
                print(f"[LABEL] Fit achieved on attempt {attempt}")
                print(f"[LABEL] Font sizes - Part: {part_font_size}, Qty: {qty_font_size}, Desc: {desc_font_size}")
                print(f"[LABEL] Total content height: {total_content_height}px / {available_text_height}px available")
                print(f"[LABEL] Description lines: {len(desc_lines)}")
                break
            
            # Check if we've reached minimum sizes
            if qty_font_size == min_qty_size and desc_font_size == min_desc_size:
                # At minimum sizes, just truncate
                print(f"[LABEL] Reached minimum font sizes, truncating content")
                print(f"[LABEL] Font sizes - Part: {part_font_size}, Qty: {qty_font_size}, Desc: {desc_font_size}")
                break
        
        # Build final layout
        layout_lines = [(part_text, part_font), (qty_text, qty_font)]
        layout_lines.extend((line, desc_font) for line in desc_lines)
        
        # Calculate metrics for rendering
        metrics = []
        total_height = line_spacing * (len(layout_lines) - 1) if layout_lines else 0
        for text, font in layout_lines:
            bbox = draw.textbbox((0, 0), text, font=font)
            line_height = bbox[3] - bbox[1]
            metrics.append((text, font, line_height))
            total_height += line_height

        # Center vertically
        current_y = max(self._scale(4), (self.height * scale - total_height) // 2)
        
        # Render text
        for idx, (text, font, line_height) in enumerate(metrics):
            draw.text((text_x, current_y), text, fill="black", font=font)
            current_y += line_height
            if idx < len(metrics) - 1:
                current_y += line_spacing

        return canvas.resize((self.width, self.height), Image.LANCZOS)

    def save_label_preview(self, label_image, output_path):
        label_image.save(output_path, "PNG")
        return output_path

    def _fit_font(self, draw, text, max_width, start_size=24, min_size=10, **font_kwargs):
        size = start_size
        decrement = max(2, self.scale * 2)
        while size >= min_size:
            font = self.get_font(size, **font_kwargs)
            try:
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
            except Exception:
                text_width = len(text) * size * 0.5
            if text_width <= max_width:
                return font
            size -= decrement
        return self.get_font(min_size, **font_kwargs)

    def _wrap_text(self, draw, text, font, max_width, max_lines=3):
        words = text.split()
        lines = []
        current = ""
        idx = 0
        truncated = False

        while idx < len(words):
            word = words[idx]
            candidate = f"{current} {word}".strip() if current else word
            try:
                bbox = draw.textbbox((0, 0), candidate, font=font)
                width = bbox[2] - bbox[0]
            except Exception:
                fallback_size = getattr(font, "size", 10)
                width = len(candidate) * fallback_size * 0.5

            if width <= max_width:
                current = candidate
                idx += 1
                continue

            if current:
                lines.append(current)
            else:
                lines.append(word)
                idx += 1
            current = ""

            if len(lines) == max_lines:
                truncated = True
                break

        if len(lines) < max_lines and current:
            lines.append(current)
        elif len(lines) == max_lines and current and not truncated:
            truncated = True
            lines[-1] = current

        if not truncated and idx < len(words):
            truncated = True

        if truncated and lines:
            lines[-1] = lines[-1].rstrip(".") + "..."

        return lines

    def _scale(self, value):
        return int(value * self.scale)
