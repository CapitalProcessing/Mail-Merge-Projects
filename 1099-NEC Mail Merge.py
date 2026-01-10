"""
1099-NEC Form Printer - Optimized for Pre-printed Forms
Prints data from CSV onto pre-printed 1099-NEC forms (3 per page)
Updated:  2026-01-10 - All alignment adjustments + right-aligned currency
"""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import json
import csv
import os

# Prompt for file paths
print("=" * 60)
print("1099-NEC Form Printer")
print("=" * 60)

# Get JSON coordinates file
json_file = input("\nEnter path to JSON coordinates file: ").strip().strip('"')
if not os.path.exists(json_file):
    print(f"ERROR: File not found: {json_file}")
    exit(1)

# Get CSV data file
csv_file = input("Enter path to CSV data file: ").strip().strip('"')
if not os.path.exists(csv_file):
    print(f"ERROR: File not found:  {csv_file}")
    exit(1)

# Get output PDF path
output_pdf = input("Enter output PDF filename (or full path): ").strip().strip('"')
if not output_pdf.lower().endswith('.pdf'):
    output_pdf += '.pdf'

print("\nLoading data...")

# Load field coordinates from JSON
with open(json_file, 'r') as f:
    field_coords_raw = json.load(f)

# Extract fields dictionary
field_coords = field_coords_raw.get('fields', field_coords_raw)

# Read CSV data
with open(csv_file, 'r') as f:
    reader = csv.DictReader(f)
    recipients_raw = list(reader)

# Map IRS CSV columns to simplified format
def map_csv_row(row):
    """Map IRS CSV format to internal format"""
    
    # Payer Name (business or individual)
    if row.get('Payer Business or Entity Name Line 1', '').strip():
        payer_name = row['Payer Business or Entity Name Line 1'].strip()
        if row.get('Payer Business or Entity Name Line 2', '').strip():
            payer_name += ' ' + row['Payer Business or Entity Name Line 2'].strip()
    else:
        # Individual name
        name_parts = []
        if row.get('Payer First Name', '').strip():
            name_parts.append(row['Payer First Name'].strip())
        if row.get('Payer Middle Name', '').strip():
            name_parts.append(row['Payer Middle Name'].strip())
        if row.get('Payer Last Name (Surname)', '').strip():
            name_parts.append(row['Payer Last Name (Surname)'].strip())
        if row.get('Payer Suffix', '').strip():
            name_parts.append(row['Payer Suffix'].strip())
        payer_name = ' '.join(name_parts)
    
    # Recipient Name (business or individual)
    if row.get('Recipient Business or Entity Name Line 1', '').strip():
        recipient_name = row['Recipient Business or Entity Name Line 1'].strip()
        if row.get('Recipient Business or Entity Name Line 2', '').strip():
            recipient_name += ' ' + row['Recipient Business or Entity Name Line 2'].strip()
    else:
        # Individual name
        name_parts = []
        if row.get('Recipient First Name', '').strip():
            name_parts.append(row['Recipient First Name'].strip())
        if row.get('Recipient Middle Name', '').strip():
            name_parts.append(row['Recipient Middle Name'].strip())
        if row.get('Recipient Last Name (Surname)', '').strip():
            name_parts.append(row['Recipient Last Name (Surname)'].strip())
        if row.get('Recipient Suffix', '').strip():
            name_parts.append(row['Recipient Suffix'].strip())
        recipient_name = ' '.join(name_parts)
    
    # Payer Address
    payer_address = row.get('Payer Address Line 1', '').strip()
    if row.get('Payer Address Line 2', '').strip():
        payer_address += ' ' + row['Payer Address Line 2'].strip()
    
    payer_city_state_zip = f"{row.get('Payer City/Town', '').strip()}, {row.get('Payer State/Province/Territory', '').strip()} {row.get('Payer ZIP/Postal Code', '').strip()}"
    
    # Recipient Address
    recipient_address = row.get('Recipient Address Line 1', '').strip()
    if row.get('Recipient Address Line 2', '').strip():
        recipient_address += ' ' + row['Recipient Address Line 2'].strip()
    
    recipient_city_state_zip = f"{row.get('Recipient City/Town', '').strip()}, {row.get('Recipient State/Province/Territory', '').strip()} {row.get('Recipient ZIP/Postal Code', '').strip()}"
    
    return {
        'Payer Name':  payer_name,
        'Payer Address': payer_address,
        'Payer City State Zip': payer_city_state_zip,
        'Payer TIN': row.get('Payer Taxpayer ID Number', ''),
        'Recipient Name': recipient_name,
        'Recipient Address': recipient_address,
        'Recipient City State Zip':  recipient_city_state_zip,
        'Recipient TIN':  row.get('Recipient Taxpayer ID Number', ''),
        'Box 1': row.get('Box 1 - Nonemployee Compensation', ''),
        'Box 2': row.get('Box 2 - Payer made direct sales totaling $5,000 or more of consumer products to a recipient for resale', ''),
        'Box 3': row.get('Box 3 - Excess golden parachute payments', ''),
        'Box 4': row.get('Box 4 - Federal income tax withheld', ''),
        'Box 5': row.get('State 1 - State tax withheld', ''),
        'Box 5a': row.get('State 2 - State tax withheld', ''),
        'Box 6': row.get('State 1 - State/Payer state number', ''),
        'Box 6a': row.get('State 2 - State/Payer state number', ''),
        'Box 7': row.get('State 1 - State income', ''),
        'Box 7a': row.get('State 2 - State income', ''),
        'Account Number': row.get('Form Account Number', ''),
        'Tax Year': row.get('Tax Year', ''),
    }

# Map all recipients
recipients = [map_csv_row(row) for row in recipients_raw]

# Create PDF
c = canvas.Canvas(output_pdf, pagesize=letter)
page_width, page_height = letter

# Store Section 1 X coordinates for Box 1, 2, 3, 4 (MASTER X coordinates)
section1_box_x = {}
for box_num in ['1', '2', '3', '4']:
    # Note: JSON has "BOX  1 -  1" with DOUBLE spaces!
    box_field = f"BOX  {box_num} -  1"
    if box_field in field_coords:
        section1_box_x[box_num] = field_coords[box_field]['x']

def draw_text_field(c, text, x, y, font_size=9):
    """Draw text at specified position (LEFT-ALIGNED)"""
    if text: 
        c.setFont("Helvetica", font_size)
        c.drawString(x, y, str(text))

def draw_text_field_right_aligned(c, text, x, width, y, font_size=9):
    """Draw text RIGHT-ALIGNED at right edge of field"""
    if text: 
        c.setFont("Helvetica", font_size)
        # Anchor at RIGHT edge (x + width - small padding for margin)
        c.drawRightString(x + width - 2, y, str(text))

def draw_multiline_text(c, text, x, y, font_size=9, line_height=10):
    """Draw multi-line text (LEFT-ALIGNED)"""
    if text:
        c.setFont("Helvetica", font_size)
        lines = text.split('\n')
        for i, line in enumerate(lines):
            c.drawString(x, y - (i * line_height), line)

def draw_form_section(c, row, suffix):
    """Draw one 1099-NEC form section with proper alignment"""
    
    # PAYER'S NAME AND ADDRESS
    payer_field = f"PAYER {suffix}"
    if payer_field in field_coords:
        coord = field_coords[payer_field]
        if suffix == "1":
            adjusted_y = coord['y']  # No change
        elif suffix == "2":
            adjusted_y = coord['y']  # No change
        else:   # Section 3
            adjusted_y = coord['y'] - 16.5  # DOWN 12 points
        
        payer_address = f"{row['Payer Name']}\n{row['Payer Address']}\n{row['Payer City State Zip']}"
        draw_multiline_text(c, payer_address, coord['x'], adjusted_y, font_size=9, line_height=10)
    
    # PAYER'S TIN (Note: JSON has "PAYERS TIN" with S!)
    payer_tin_field = f"PAYERS TIN {suffix}"
    if payer_tin_field in field_coords:
        coord = field_coords[payer_tin_field]
        if suffix == "1": 
            adjusted_y = coord['y'] + 12  # UP 3 points (from +9 to +12)
        elif suffix == "2":
            adjusted_y = coord['y'] + 2  # UP 2 points
        else:  # Section 3
            adjusted_y = coord['y'] - 1.5  # UP 3 points (from -4.5 to -1.5)
        
        draw_text_field(c, row['Payer TIN'], coord['x'], adjusted_y, font_size=9)
    
    # RECIPIENT'S NAME AND ADDRESS (JSON has combined field!)
    recipient_field = f"RECIPIENT NAME ADDRESS BLOCK {suffix}"
    if recipient_field in field_coords:
        coord = field_coords[recipient_field]
        if suffix == "1":
            adjusted_y = coord['y']  # No change
        elif suffix == "2":
            adjusted_y = coord['y']  # No change
        else:  # Section 3
            adjusted_y = coord['y'] - 16.5  # DOWN 12 points
        
        recipient_address = f"{row['Recipient Name']}\n{row['Recipient Address']}\n{row['Recipient City State Zip']}"
        draw_multiline_text(c, recipient_address, coord['x'], adjusted_y, font_size=9, line_height=10)
    
    # RECIPIENT'S TIN (Note: JSON has "RECIPIENT'S TIN" with apostrophe-S!)
    recipient_tin_field = f"RECIPIENT'S TIN {suffix}"
    if recipient_tin_field in field_coords:
        coord = field_coords[recipient_tin_field]
        if suffix == "1": 
            adjusted_y = coord['y'] + 12  # UP 3 points (from +9 to +12)
        elif suffix == "2":
            adjusted_y = coord['y'] + 2  # UP 2 points
        else:  # Section 3
            adjusted_y = coord['y'] - 1.5  # UP 3 points (from -4.5 to -1.5)
        
        draw_text_field(c, row['Recipient TIN'], coord['x'], adjusted_y, font_size=9)
    
    # BOX 1 - Nonemployee compensation (RIGHT-ALIGNED, using Section 1 X)
    # Note: JSON has "BOX  1 -  1" with DOUBLE spaces!
    box1_field = f"BOX  1 -  {suffix}"
    if box1_field in field_coords:
        coord = field_coords[box1_field]
        
        # Use Section 1 X for all sections
        if suffix == "1":
            adjusted_x = coord['x']  # Use Section 1 X (master)
            adjusted_y = coord['y'] + 9  # No change
        else: 
            adjusted_x = section1_box_x.get('1', coord['x'])  # Use Section 1 X
            if suffix == "2":
                adjusted_y = coord['y'] + 2  # UP 2 points
            else:  # Section 3
                adjusted_y = coord['y'] - 1.5  # UP 16.5 points (from -18 to -1.5)
        
        draw_text_field_right_aligned(c, row['Box 1'], adjusted_x, coord['width'], adjusted_y, font_size=10)
    
    # BOX 2 - Payer made direct sales checkbox (LEFT-ALIGNED, using Section 1 X)
    box2_field = f"BOX 2 - {suffix}"
    if box2_field in field_coords: 
        coord = field_coords[box2_field]
        
        # Use Section 1 X for all sections
        if suffix == "1":
            adjusted_x = coord['x']  # Use Section 1 X (master)
            adjusted_y = coord['y'] + 12  # UP 3 points (from +9 to +12)
        else:
            adjusted_x = section1_box_x.get('2', coord['x'])  # Use Section 1 X
            if suffix == "2":
                adjusted_y = coord['y'] + 8.5  # UP 22 points (from -13.5 to +8.5)
            else:  # Section 3
                adjusted_y = coord['y'] - 13.5  # No change
        
        if row['Box 2'].strip().upper() in ['Y', 'YES', 'X', 'N']: 
            draw_text_field(c, 'N' if row['Box 2'].strip().upper() == 'N' else 'X', adjusted_x, adjusted_y, font_size=10)
    
    # BOX 3 - Excess golden parachute (RIGHT-ALIGNED, using Section 1 X)
    box3_field = f"BOX 3 - {suffix}"
    if box3_field in field_coords:
        coord = field_coords[box3_field]
        
        # Use Section 1 X for all sections
        if suffix == "1":
            adjusted_x = coord['x']  # Use Section 1 X (master)
            adjusted_y = coord['y'] + 12  # UP 3 points (from +9 to +12)
        else:
            adjusted_x = section1_box_x.get('3', coord['x'])  # Use Section 1 X
            if suffix == "2":
                adjusted_y = coord['y'] + 3.5  # DOWN 1 point (from +4.5 to +3.5)
            else:  # Section 3
                adjusted_y = coord['y'] + 7.5  # UP 3 points (from +4.5 to +7.5)
        
        draw_text_field_right_aligned(c, row['Box 3'], adjusted_x, coord['width'], adjusted_y, font_size=10)
    
    # BOX 4 - Federal income tax withheld (RIGHT-ALIGNED, using Section 1 X)
    box4_field = f"BOX 4 - {suffix}"
    if box4_field in field_coords:
        coord = field_coords[box4_field]
        
        # Use Section 1 X for all sections
        if suffix == "1":
            adjusted_x = coord['x']  # Use Section 1 X (master)
            adjusted_y = coord['y'] + 12  # UP 3 points (from +9 to +12)
        else:
            adjusted_x = section1_box_x.get('4', coord['x'])  # Use Section 1 X
            if suffix == "2":
                adjusted_y = coord['y']  # No change (0)
            else:  # Section 3
                adjusted_y = coord['y'] - 4  # DOWN 4 points (from 0 to -4)
        
        draw_text_field_right_aligned(c, row['Box 4'], adjusted_x, coord['width'], adjusted_y, font_size=10)
    
    # BOX 5 - State 1 tax withheld (RIGHT-ALIGNED, X + 36pts)
    box5_field = f"BOX 5 - {suffix}"
    if box5_field in field_coords: 
        coord = field_coords[box5_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1":
            adjusted_y = coord['y']  # No change from original
        elif suffix == "2": 
            adjusted_y = coord['y'] - 2.25  # Keep current adjustment
        else:  # Section 3
            adjusted_y = coord['y'] - 0.25  # UP 2 points (from -2.25 to -0.25)
        
        draw_text_field_right_aligned(c, row['Box 5'], adjusted_x, coord['width'], adjusted_y, font_size=9)
    
    # BOX 5a - State 2 tax withheld (RIGHT-ALIGNED, X + 36pts, uses Box 6a Y)
    box5a_field = f"BOX 5a - {suffix}"
    box6a_field = f"BOX 6a - {suffix}"
    if box5a_field in field_coords and box6a_field in field_coords:
        coord = field_coords[box5a_field]
        coord_6a = field_coords[box6a_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1": 
            adjusted_y = coord_6a['y']  # Use Box 6a Y
        elif suffix == "2":
            adjusted_y = coord_6a['y'] - 2.25  # Use Box 6a Y with adjustment
        else:  # Section 3
            adjusted_y = coord_6a['y'] - 0.25  # Use Box 6a Y, UP 2 points
        
        draw_text_field_right_aligned(c, row.get('Box 5a', ''), adjusted_x, coord['width'], adjusted_y, font_size=9)
    
    # BOX 6 - State 1 number (LEFT-ALIGNED, X + 36pts)
    box6_field = f"BOX 6 - {suffix}"
    if box6_field in field_coords:
        coord = field_coords[box6_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1":
            adjusted_y = coord['y']  # No change from original
        elif suffix == "2":
            adjusted_y = coord['y'] - 2.25  # Keep current adjustment
        else:  # Section 3
            adjusted_y = coord['y'] - 0.25  # UP 2 points (from -2.25 to -0.25)
        
        draw_text_field(c, row['Box 6'], adjusted_x, adjusted_y, font_size=9)
    
    # BOX 6a - State 2 number (LEFT-ALIGNED, X + 36pts)
    box6a_field = f"BOX 6a - {suffix}"
    if box6a_field in field_coords: 
        coord = field_coords[box6a_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1":
            adjusted_y = coord['y']  # No change from original
        elif suffix == "2":
            adjusted_y = coord['y'] - 2.25  # Keep current adjustment
        else:  # Section 3
            adjusted_y = coord['y'] - 0.25  # UP 2 points (from -2.25 to -0.25)
        
        draw_text_field(c, row.get('Box 6a', ''), adjusted_x, adjusted_y, font_size=9)
    
    # BOX 7 - State 1 income (RIGHT-ALIGNED, X + 36pts)
    box7_field = f"BOX 7 - {suffix}"
    if box7_field in field_coords: 
        coord = field_coords[box7_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1":
            adjusted_y = coord['y']  # No change from original
        elif suffix == "2":
            adjusted_y = coord['y'] - 2.25  # Keep current adjustment
        else:  # Section 3
            adjusted_y = coord['y'] - 0.25  # UP 2 points (from -2.25 to -0.25)
        
        draw_text_field_right_aligned(c, row['Box 7'], adjusted_x, coord['width'], adjusted_y, font_size=9)
    
    # BOX 7a - State 2 income (RIGHT-ALIGNED, X + 36pts)
    box7a_field = f"BOX 7a - {suffix}"
    if box7a_field in field_coords: 
        coord = field_coords[box7a_field]
        adjusted_x = coord['x'] + 36  # Move right 36 points
        
        if suffix == "1":
            adjusted_y = coord['y']  # No change from original
        elif suffix == "2":
            adjusted_y = coord['y'] - 2.25  # Keep current adjustment
        else:  # Section 3
            adjusted_y = coord['y'] - 0.25  # UP 2 points (from -2.25 to -0.25)
        
        draw_text_field_right_aligned(c, row.get('Box 7a', ''), adjusted_x, coord['width'], adjusted_y, font_size=9)
    
    # ACCOUNT NUMBER
    acct_field = f"ACCT NUMBER {suffix}"
    if acct_field in field_coords:
        coord = field_coords[acct_field]
        if suffix == "1":
            adjusted_y = coord['y'] + 7.5  # UP 3 points (from +4.5 to +7.5)
        elif suffix == "2": 
            adjusted_y = coord['y'] + 4.5  # Keep current
        else:  # Section 3
            adjusted_y = coord['y'] + 2.5  # DOWN 2 points (from +4.5 to +2.5)
        
        draw_text_field(c, row['Account Number'], coord['x'], adjusted_y, font_size=9)
    
    # YEAR
    year_field = f"YEAR {suffix}"
    if year_field in field_coords: 
        coord = field_coords[year_field]
        adjusted_x = coord['x'] + 18  # Keep X adjustment
        
        if suffix == "1":
            adjusted_y = coord['y'] + 13.5  # No change
        elif suffix == "2":
            adjusted_y = coord['y'] + 6  # DOWN 3 points (from +9 to +6)
        else:  # Section 3
            adjusted_y = coord['y'] - 6  # UP 12 points (from -18 to -6)
        
        draw_text_field(c, row.get('Tax Year', ''), adjusted_x, adjusted_y, font_size=10)

# Print forms - 3 per page
print(f"Generating forms for {len(recipients)} recipients...")
for i in range(0, len(recipients), 3):
    page_recipients = recipients[i:i+3]
    
    # Draw each section
    for idx, recipient in enumerate(page_recipients):
        section_suffix = str(idx + 1)
        draw_form_section(c, recipient, section_suffix)
    
    # Start new page if more recipients remain
    if i + 3 < len(recipients):
        c.showPage()

# Save PDF
c.save()
print("\n" + "=" * 60)
print(f"✅ PDF created: {output_pdf}")
print(f"✅ Processed {len(recipients)} recipients")
print(f"✅ All adjustments applied:")
print(f"   - Currency boxes RIGHT-ALIGNED (Box 1,3,4,5,5a,7,7a)")
print(f"   - Box 1,2,3,4 use consistent X across all sections")
print(f"   - State columns moved right 36 points")
print(f"   - All Y adjustments per specifications")
print("=" * 60)