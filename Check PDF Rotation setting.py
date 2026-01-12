from PyPDF2 import PdfReader

pdf = PdfReader(r"C:\Python\Mail Merge\2025 1099-NEC-5111 Blank-CORRECTED.pdf")
page = pdf.pages[0]
rotation = page.get('/Rotate', 0)
print(f"PDF Rotation: {rotation} degrees")