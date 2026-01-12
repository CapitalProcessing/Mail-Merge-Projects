import fitz  # PyMuPDF
from PIL import Image

# Open the background PDF
pdf_path = r"C:\Python\Mail Merge\2025 1099-NEC-5111 Blank. pdf"
doc = fitz.open(pdf_path)
page = doc[0]

print(f"PDF rotation flag: {page.rotation}°")

# Convert to high-resolution PNG (300 DPI)
zoom = 300 / 72  # 300 DPI
mat = fitz.Matrix(zoom, zoom)
pix = page.get_pixmap(matrix=mat)

# Save to temporary bytes
import io
img_data = pix.tobytes("png")

# Open with PIL
img = Image.open(io.BytesIO(img_data))

print(f"Original image size: {img.size}")
print(f"Rotating to correct orientation...")

# Rotate to make it upright (counterclockwise rotations)
# If sideways with text reading from bottom-to-top, rotate 90° clockwise (270° CCW)
img_rotated = img.rotate(-90, expand=True)  # -90 = 90° clockwise

# OR try this if text reads top-to-bottom: 
# img_rotated = img.rotate(90, expand=True)  # 90° counterclockwise

print(f"Rotated image size: {img_rotated.size}")

# Save as PNG
output_png = r"C:\Python\Mail Merge\2025 1099-NEC-5111 Blank.png"
img_rotated.save(output_png, dpi=(300, 300))

doc.close()

print(f"\n✓ Created: {output_png}")
print(f"  Size: {img_rotated.width} x {img_rotated.height} pixels")
print(f"\nOpen the PNG to verify it's upright before proceeding!")