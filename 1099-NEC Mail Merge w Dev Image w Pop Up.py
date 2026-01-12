"""
1099-NEC Form Filler - Direct Coordinate Approach
Uses JSON master template with exact field positions
Toggle USE_BACKGROUND_IMAGE for development vs production
Compatible with IRS Publication 1220 CSV format
"""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import json
import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================================================================
# CONFIGURATION
# ============================================================================

# TOGGLE THIS:      True = show background (development), False = data only (production)
USE_BACKGROUND_IMAGE = False

# Font settings
FONT_NAME = "Courier"
FONT_SIZE = 9

# Section spacing
SECTION_1_Y_OFFSET = 18      # Move down (-) or up (+)
SECTION_2_Y_OFFSET = -257    # Move down (-) or up (+)
SECTION_3_Y_OFFSET = -532    # Move down (-) or up (+)

# Background image offset and scaling (for independent adjustment during development)
BACKGROUND_IMAGE_X_OFFSET = -3       # Move background left (-) or right (+)
BACKGROUND_IMAGE_Y_OFFSET = -15.5    # Move background down (-) or up (+)
BACKGROUND_IMAGE_WIDTH_STRETCH = 22.5     # Add/subtract width 
BACKGROUND_IMAGE_HEIGHT_STRETCH = 31    # Add/subtract height

# Right-aligned fields (numbers/amounts)
RIGHT_ALIGNED_FIELDS = {
    'BOX 1', 'BOX 3', 'BOX 4',
    'BOX 5', 'BOX 5a', 'BOX 6', 'BOX 6a', 'BOX 7', 'BOX 7a'
}

# ============================================================================
# FILE SELECTION FUNCTIONS
# ============================================================================

def show_large_message(title, message):
    """Show a large, readable message box"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Configure message box font (larger and bold)
    messagebox.showinfo(title, message)
    root.destroy()

def select_csv_file():
    """Prompt user to select CSV input file"""
    show_large_message(
        "Step 1: Select CSV Data File",
        "📄 SELECT YOUR CSV INPUT FILE\n\n"
        "This file contains the 1099-NEC data for all recipients.\n\n"
        "Format: IRS Publication 1220 CSV format\n"
        "Example: 1099_CAI_pdf_IRS_Test-Data_2025.csv\n\n"
        "Click OK to browse for your CSV file..."
    )
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select CSV Data File",
        filetypes=[
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~")
    )
    
    root.destroy()
    
    if not file_path: 
        messagebox.showerror("Error", "No CSV file selected.  Exiting.")
        exit()
    
    return file_path

def select_json_file():
    """Prompt user to select JSON field positions file"""
    show_large_message(
        "Step 2: Select JSON Field Positions",
        "📐 SELECT YOUR JSON FIELD POSITIONS FILE\n\n"
        "This file contains the X/Y coordinates for all form fields.\n\n"
        "Example: JSON Sourced - 2025 1099-NEC w Section 1 Fillable_fields.json\n\n"
        "Click OK to browse for your JSON file..."
    )
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select JSON Field Positions File",
        filetypes=[
            ("JSON files", "*.json"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~")
    )
    
    root.destroy()
    
    if not file_path:
        messagebox.showerror("Error", "No JSON file selected. Exiting.")
        exit()
    
    return file_path

def select_background_image():
    """Prompt user to select background image (only in dev mode)"""
    show_large_message(
        "Step 3: Select Background Image",
        "🖼️ SELECT YOUR BACKGROUND IMAGE FILE\n\n"
        "This is the scanned image of a blank 1099-NEC form.\n"
        "Used for development/alignment purposes.\n\n"
        "Example: 2025 1099-NEC-5111 Blank Top Top.png\n\n"
        "Click OK to browse for your image file..."
    )
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Background Image File (Dev Mode)",
        filetypes=[
            ("PNG files", "*.png"),
            ("JPEG files", "*.jpg;*.jpeg"),
            ("All image files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~")
    )
    
    root.destroy()
    
    if not file_path:
        messagebox.showwarning("Warning", "No background image selected.Continuing without image.")
    
    return file_path

def select_output_location(mode):
    """Prompt user to select output PDF location"""
    if mode == "DEV":
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"1099_NEC_Forms_DEV_{timestamp}.pdf"
        message_text = (
            "💾 SELECT OUTPUT LOCATION (Development Mode)\n\n"
            "Choose where to save your TEST PDF file.\n"
            "This file will include the background image for alignment.\n\n"
            f"Suggested filename: {default_filename}\n\n"
            "Click OK to choose location..."
        )
    else:
        default_filename = "1099_NEC_Forms_Filled.pdf"
        message_text = (
            "💾 SELECT OUTPUT LOCATION (Production Mode)\n\n"
            "Choose where to save your FINAL PDF file.\n"
            "This file will contain ONLY the data (no background).\n"
            "Print this on pre-printed IRS 1099-NEC forms.\n\n"
            f"Suggested filename:  {default_filename}\n\n"
            "Click OK to choose location..."
        )
    
    show_large_message(
        "Step 4: Select Output Location",
        message_text
    )
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.asksaveasfilename(
        title="Save PDF Output As",
        defaultextension=".pdf",
        initialfile=default_filename,
        filetypes=[
            ("PDF files", "*.pdf"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~")
    )
    
    root.destroy()
    
    if not file_path:
        messagebox.showerror("Error", "No output location selected. Exiting.")
        exit()
    
    return file_path

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def format_currency(amount):
    """Format currency with commas and 2 decimal places"""
    if not amount or amount == '0' or amount == '0.00':
        return ''
    try:
        return f"{float(amount):,.2f}"
    except:
        return amount

def format_tin(tin):
    """Format TIN as XX-XXXXXXX"""
    tin = str(tin).replace('-', '').replace(' ', '')
    if len(tin) == 9:
        return f"{tin[:2]}-{tin[2:]}"
    return tin

def get_recipient_name(recipient):
    """Get recipient name from IRS CSV format (handles both business and individual)"""
    # Check for business name first
    business_name = recipient.get('Recipient Business or Entity Name Line 1', '').strip()
    if business_name:
        return business_name
    
    # Otherwise construct from first/last name
    first_name = recipient.get('Recipient First Name', '').strip()
    middle_name = recipient.get('Recipient Middle Name', '').strip()
    last_name = recipient.get('Recipient Last Name (Surname)', '').strip()
    suffix = recipient.get('Recipient Suffix', '').strip()
    
    name_parts = [first_name, middle_name, last_name, suffix]
    return ' '.join([part for part in name_parts if part])

def get_payer_name(recipient):
    """Get payer name from IRS CSV format (handles both business and individual)"""
    # Check for business name first
    business_name = recipient.get('Payer Business or Entity Name Line 1', '').strip()
    if business_name:  
        return business_name
    
    # Otherwise construct from first/last name
    first_name = recipient.get('Payer First Name', '').strip()
    middle_name = recipient.get('Payer Middle Name', '').strip()
    last_name = recipient.get('Payer Last Name (Surname)', '').strip()
    suffix = recipient.get('Payer Suffix', '').strip()
    
    name_parts = [first_name, middle_name, last_name, suffix]
    return ' '.join([part for part in name_parts if part])

def draw_text(c, field_name, text, y_offset=0):
    """Draw text at field position with proper alignment"""
    if field_name not in MASTER_FIELDS:
        print(f"⚠️  Field '{field_name}' not found in master template")
        return
    
    if not text or str(text).strip() == '':
        return
    
    pos = MASTER_FIELDS[field_name]
    x = pos['x']
    y = pos['y'] + y_offset
    
    # Right-align numeric fields
    if field_name in RIGHT_ALIGNED_FIELDS:    
        text_width = c.stringWidth(str(text), FONT_NAME, FONT_SIZE)
        x = x - text_width
    
    c.drawString(x, y, str(text))

def draw_multiline_address(c, field_name, lines, y_offset=0):
    """Draw multi-line address with proper line spacing"""
    if field_name not in MASTER_FIELDS:  
        return
    
    pos = MASTER_FIELDS[field_name]
    x = pos['x']
    y = pos['y'] + y_offset
    
    line_height = FONT_SIZE * 1.2  # 120% line spacing
    
    for i, line in enumerate(lines):
        if line and line.strip():
            c.drawString(x, y - (i * line_height), line.strip())

def format_address_lines(name, address, city, state, zip_code):
    """Format address into lines"""
    lines = []
    if name:
        lines.append(name)
    if address:
        lines.append(address)
    
    city_state_zip = []
    if city:  
        city_state_zip.append(city)
    if state:  
        city_state_zip.append(state)
    if zip_code:  
        city_state_zip.append(zip_code)
    
    if city_state_zip:
        lines.append(' '.join(city_state_zip))
    
    return lines

# ============================================================================
# MAIN FORM FILLING FUNCTION
# ============================================================================

def fill_1099_nec_form(csv_file, output_pdf, background_image_path=None):
    """Fill 1099-NEC forms from CSV data"""
    
    # Read CSV data
    with open(csv_file, 'r') as f:
        reader = csv.DictReader(f)
        recipients = list(reader)
    
    if not recipients:
        print("No data found in CSV!")
        return
    
    # Show mode
    mode = "DEVELOPMENT (with background)" if USE_BACKGROUND_IMAGE else "PRODUCTION (data only)"
    print(f"\n{'='*80}")
    print(f"Processing {len(recipients)} recipients")
    print(f"Mode:      {mode}")
    print(f"{'='*80}\n")
    
    # Create PDF
    c = canvas.Canvas(output_pdf, pagesize=letter)

    # Process each recipient (3 per page)
    for page_num, i in enumerate(range(0, len(recipients), 3)):
        page_recipients = recipients[i:i+3]
    
        print(f"Page {page_num + 1}:      Processing {len(page_recipients)} sections")
    
        # Set font for THIS page (must be done after showPage())
        c.setFont(FONT_NAME, FONT_SIZE)
    
        # Draw background image FIRST if enabled
        if USE_BACKGROUND_IMAGE and background_image_path and os.path.exists(background_image_path):
            c.drawImage(background_image_path, 
                       BACKGROUND_IMAGE_X_OFFSET, 
                       BACKGROUND_IMAGE_Y_OFFSET, 
                       width=letter[0] + BACKGROUND_IMAGE_WIDTH_STRETCH, 
                       height=letter[1] + BACKGROUND_IMAGE_HEIGHT_STRETCH, 
                       preserveAspectRatio=False, 
                       mask='auto')
        
        # Fill each section
        for section_num, recipient in enumerate(page_recipients):
            y_offset = [SECTION_1_Y_OFFSET, SECTION_2_Y_OFFSET, SECTION_3_Y_OFFSET][section_num]
            
            recipient_name = get_recipient_name(recipient)
            print(f"  Section {section_num + 1}:  {recipient_name}")
            
            # PAYER INFORMATION
            payer_name = get_payer_name(recipient)
            payer_lines = format_address_lines(
                payer_name,
                recipient.get('Payer Address Line 1', ''),
                recipient.get('Payer City/Town', ''),
                recipient.get('Payer State/Province/Territory', ''),
                recipient.get('Payer ZIP/Postal Code', '')
            )
            draw_multiline_address(c, 'PAYER', payer_lines, y_offset)
            
            # PAYER'S TIN
            payer_tin = format_tin(recipient.get('Payer Taxpayer ID Number', ''))
            draw_text(c, "PAYER'S TIN", payer_tin, y_offset)
            
            # RECIPIENT INFORMATION
            recipient_lines = format_address_lines(
                recipient_name,
                recipient.get('Recipient Address Line 1', ''),
                recipient.get('Recipient City/Town', ''),
                recipient.get('Recipient State/Province/Territory', ''),
                recipient.get('Recipient ZIP/Postal Code', '')
            )
            draw_multiline_address(c, 'RECIPIENT', recipient_lines, y_offset)
            
            # RECIPIENT'S TIN
            recipient_tin = format_tin(recipient.get('Recipient Taxpayer ID Number', ''))
            draw_text(c, "RECIPIENT'S TIN", recipient_tin, y_offset)
            
            # ACCOUNT NUMBER
            account_num = recipient.get('Form Account Number', '')
            draw_text(c, 'ACCOUNT NUMBER', account_num, y_offset)
            
            # YEAR
            year = recipient.get('Tax Year', str(datetime.now().year))
            draw_text(c, 'YEAR', year, y_offset)
            
            # BOX 1 - Nonemployee compensation
            box1 = format_currency(recipient.get('Box 1 - Nonemployee Compensation', ''))
            draw_text(c, 'BOX 1', box1, y_offset)
            
            # BOX 2 - Direct sales checkbox
            box2_value = recipient.get('Box 2 - Payer made direct sales totaling $5,000 or more of consumer products to a recipient for resale', '')
            if box2_value and str(box2_value).strip().upper() in ['YES', 'Y', 'X', 'TRUE', '1']: 
                draw_text(c, 'BOX 2', 'X', y_offset)
            
            # BOX 3 - Other income
            box3 = format_currency(recipient.get('Box 3 - Excess golden parachute payments', ''))
            draw_text(c, 'BOX 3', box3, y_offset)
            
            # BOX 4 - Federal income tax withheld
            box4 = format_currency(recipient.get('Box 4 - Federal income tax withheld', ''))
            draw_text(c, 'BOX 4', box4, y_offset)
            
            # BOX 5/5a - State tax withheld
            box5 = format_currency(recipient.get('State 1 - State tax withheld', ''))
            draw_text(c, 'BOX 5', box5, y_offset)
            
            box5a = format_currency(recipient.get('State 2 - State tax withheld', ''))
            draw_text(c, 'BOX 5a', box5a, y_offset)
            
            # BOX 6/6a - State/Payer's state no.     
            state1 = recipient.get('State 1', '')
            payer_state_no1 = recipient.get('State 1 - State/Payer state number', '')
            if state1 or payer_state_no1:
                box6_value = f"{state1}/{payer_state_no1}" if state1 and payer_state_no1 else (state1 or payer_state_no1)
                draw_text(c, 'BOX 6', box6_value, y_offset)
            
            state2 = recipient.get('State 2', '')
            payer_state_no2 = recipient.get('State 2 - State/Payer state number', '')
            if state2 or payer_state_no2:
                box6a_value = f"{state2}/{payer_state_no2}" if state2 and payer_state_no2 else (state2 or payer_state_no2)
                draw_text(c, 'BOX 6a', box6a_value, y_offset)
            
            # BOX 7/7a - State income
            box7 = format_currency(recipient.get('State 1 - State income', ''))
            draw_text(c, 'BOX 7', box7, y_offset)
            
            box7a = format_currency(recipient.get('State 2 - State income', ''))
            draw_text(c, 'BOX 7a', box7a, y_offset)
        
        # Finish page
        c.showPage()
    
    # Save PDF
    c.save()
    
    print(f"\n✓ Created:     {output_pdf}")
    print(f"✓ Total pages:      {(len(recipients) + 2) // 3}")
    print(f"{'='*80}\n")
    
    # Show completion message
    messagebox.showinfo(
        "Success!",
        f"✅ PDF CREATED SUCCESSFULLY!\n\n"
        f"File:  {os.path.basename(output_pdf)}\n"
        f"Location: {os.path.dirname(output_pdf)}\n\n"
        f"Recipients processed: {len(recipients)}\n"
        f"Total pages: {(len(recipients) + 2) // 3}\n\n"
        f"Mode: {mode}"
    )

# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == "__main__":
    # Welcome message
    show_large_message(
        "1099-NEC Form Filler",
        "📋 WELCOME TO 1099-NEC FORM FILLER\n\n"
        f"Mode: {'DEVELOPMENT (with background)' if USE_BACKGROUND_IMAGE else 'PRODUCTION (data only)'}\n\n"
        "You will be prompted to select:\n"
        "  1. CSV data file (recipient information)\n"
        "  2. JSON field positions file (coordinates)\n"
        f"  {'3. Background image file (for alignment)' if USE_BACKGROUND_IMAGE else ''}\n"
        f"  {'4' if USE_BACKGROUND_IMAGE else '3'}.Output PDF location\n\n"
        "Click OK to begin..."
    )
    
    # Step 1: Select CSV file
    CSV_FILE = select_csv_file()
    print(f"✓ CSV file selected: {CSV_FILE}")
    
    # Step 2: Select JSON file
    JSON_FILE = select_json_file()
    print(f"✓ JSON file selected: {JSON_FILE}")
    
    # Load JSON field positions
    with open(JSON_FILE, 'r') as f:
        FIELD_POSITIONS = json.load(f)
    MASTER_FIELDS = FIELD_POSITIONS['fields']
    
    # Calculate 5a, 6a, 7a positions
    HALF_LINE_OFFSET = 13.5
    if 'BOX 5' in MASTER_FIELDS:  
        MASTER_FIELDS['BOX 5a'] = {
            'x': MASTER_FIELDS['BOX 5']['x'],
            'y': MASTER_FIELDS['BOX 5']['y'] - HALF_LINE_OFFSET
        }
    if 'BOX 6' in MASTER_FIELDS:   
        MASTER_FIELDS['BOX 6a'] = {
            'x': MASTER_FIELDS['BOX 6']['x'],
            'y':  MASTER_FIELDS['BOX 6']['y'] - HALF_LINE_OFFSET
        }
    if 'BOX 7' in MASTER_FIELDS:   
        MASTER_FIELDS['BOX 7a'] = {
            'x': MASTER_FIELDS['BOX 7']['x'],
            'y':  MASTER_FIELDS['BOX 7']['y'] - HALF_LINE_OFFSET
        }
    
    # Step 3: Select background image (only in dev mode)
    BACKGROUND_IMAGE_PATH = None
    if USE_BACKGROUND_IMAGE: 
        BACKGROUND_IMAGE_PATH = select_background_image()
        if BACKGROUND_IMAGE_PATH:
            print(f"✓ Background image selected: {BACKGROUND_IMAGE_PATH}")
    
    # Step 4: Select output location
    mode_name = "DEV" if USE_BACKGROUND_IMAGE else "PROD"
    OUTPUT_PDF = select_output_location(mode_name)
    print(f"✓ Output location selected: {OUTPUT_PDF}")
    
    # Process the forms
    fill_1099_nec_form(CSV_FILE, OUTPUT_PDF, BACKGROUND_IMAGE_PATH)