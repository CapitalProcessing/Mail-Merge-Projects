from PyPDF2 import PdfReader, PdfWriter

# Read the ORIGINAL rotated PDF
reader = PdfReader(r"C:\Python\Mail Merge\2025 1099-NEC-5111 Blank.pdf")
page = reader.pages[0]

print(f"Original rotation: {page.get('/Rotate', 0)}°")

# Create writer
writer = PdfWriter()

# Remove rotation flag
if '/Rotate' in page:
    del page['/Rotate']
    print("Removed rotation flag")

# Rotate -90° (same as 270°) to counter the physical rotation
page.rotate(-90)
print("Rotated -90°")

writer.add_page(page)

# Save
output_path = r"C:\Python\Mail Merge\2025 1099-NEC-5111 Blank-CORRECTED.pdf"
with open(output_path, 'wb') as f:
    writer.write(f)

print(f"\n✓ Created:  {output_path}")

# Verify
verify = PdfReader(output_path)
verify_page = verify.pages[0]
print(f"\nNew rotation: {verify_page.get('/Rotate', 0)}°")