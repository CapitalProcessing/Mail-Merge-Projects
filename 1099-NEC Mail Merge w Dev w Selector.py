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
import tempfile

# ============================================================================
# CONFIGURATION
# ============================================================================

# TOGGLE THIS:  True = show background (development), False = data only (production)
USE_BACKGROUND_IMAGE = False

# Font settings
FONT_NAME = "Courier"
FONT_SIZE = 9

# Section spacing
SECTION_1_Y_OFFSET = 22      # Move down (-) or up (+)
SECTION_2_Y_OFFSET = -253    # Move down (-) or up (+)
SECTION_3_Y_OFFSET = -528    # Move down (-) or up (+)

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
        "Step 3: Select JSON Field Positions",
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
        "Step 4: Select Background Image",
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
        messagebox.showwarning("Warning", "No background image selected. Continuing without image.")
    
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
    
    step_number = "5" if USE_BACKGROUND_IMAGE else "4"
    show_large_message(
        f"Step {step_number}: Select Output Location",
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
        messagebox.showerror("Error", "No output location selected.Exiting.")
        exit()
    
    return file_path

# ============================================================================
# ROW SELECTION FUNCTIONS
# ============================================================================

def select_all_or_some():
    """Ask user if they want to process all rows or specific rows"""
    root = tk.Tk()
    root.withdraw()
    
    result = messagebox.askyesno(
        "Row Selection",
        "📋 PROCESS ALL RECIPIENTS OR SELECT SPECIFIC ROWS?\n\n"
        "• Click YES to process ALL rows from the CSV file\n"
        "• Click NO to select SPECIFIC rows only\n\n"
        "Example use case for 'NO':\n"
        "  - You need to reprint forms for specific recipients\n"
        "  - You want to test with just a few entries\n"
        "  - You discovered missing data for certain rows\n\n"
        "Process ALL recipients?"
    )
    
    root.destroy()
    return result

def get_row_numbers(total_rows):
    """Prompt user to enter specific row numbers"""
    
    # Show instructions first
    instructions_msg = (
        f"📋 YOUR CSV FILE HAS {total_rows} DATA ROWS\n\n"
        "In the next dialog, enter the row numbers you want to process:\n\n"
        "EXAMPLES:\n"
        "  • Single rows:          1,3,5,7\n"
        "  • Range of rows:     1-5\n"
        "  • Mixed format:      1,3,5-10,15,20-25\n"
        "  • Just one row:      7\n\n"
        "NOTES:\n"
        "  • Row 1 = first data row (after the header)\n"
        "  • Separate with commas\n"
        "  • Use hyphens for ranges\n"
        "  • Spaces are ignored\n\n"
        "Click OK to continue..."
    )
    
    messagebox.showinfo("Enter Row Numbers", instructions_msg)
    
    # Create visible root window for input
    root = tk.Tk()
    root.title("Enter Row Numbers")
    
    # Center and size the window - MADE LARGER
    window_width = 600
    window_height = 300
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f'{window_width}x{window_height}+{x}+{y}')
    
    # Make it topmost
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    
    # Create frame with more padding
    main_frame = tk.Frame(root, padx=30, pady=25)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Label
    label = tk.Label(
        main_frame,
        text=f"Enter row numbers to process\n(CSV has {total_rows} rows):",
        font=("Arial", 12, "bold"),
        justify=tk.LEFT
    )
    label.pack(pady=(0, 15))
    
    # Entry
    entry_var = tk.StringVar()
    entry = tk.Entry(main_frame, textvariable=entry_var, font=("Arial", 13), width=50)
    entry.pack(pady=15, fill=tk.X, ipady=5)  # Added ipady for taller entry
    entry.focus_set()
    
    # Example label
    example_label = tk.Label(
        main_frame,
        text="Examples:  1,3,5  or  1-5  or  1,3,5-10,15",
        font=("Arial", 10),
        fg="gray"
    )
    example_label.pack(pady=(0, 20))
    
    # Result storage
    result = {"value": None, "cancelled": False}
    
    def on_ok(event=None):
        result["value"] = entry_var.get()
        root.destroy()
    
    def on_cancel(event=None):
        result["cancelled"] = True
        root.destroy()
    
    # Buttons frame with more space
    button_frame = tk.Frame(main_frame)
    button_frame.pack(pady=20)
    
    # Create OK button using Frame (works better on Windows)
    ok_frame = tk.Frame(
        button_frame,
        bg="#4CAF50",
        relief=tk.RAISED,
        borderwidth=3
    )
    ok_frame.pack(side=tk.LEFT, padx=25)
    
    ok_label = tk.Label(
        ok_frame,
        text="OK",
        font=("Arial", 16, "bold"),
        bg="#4CAF50",
        fg="white",
        width=12,
        height=3,
        cursor="hand2"
    )
    ok_label.pack(padx=5, pady=5)
    ok_label.bind("<Button-1>", on_ok)
    ok_frame.bind("<Button-1>", on_ok)
    
    # Create Cancel button using Frame
    cancel_frame = tk.Frame(
        button_frame,
        bg="#E0E0E0",
        relief=tk.RAISED,
        borderwidth=3
    )
    cancel_frame.pack(side=tk.LEFT, padx=25)
    
    cancel_label = tk.Label(
        cancel_frame,
        text="Cancel",
        font=("Arial", 16),
        bg="#E0E0E0",
        fg="black",
        width=12,
        height=3,
        cursor="hand2"
    )
    cancel_label.pack(padx=5, pady=5)
    cancel_label.bind("<Button-1>", on_cancel)
    cancel_frame.bind("<Button-1>", on_cancel)
    
    # Bind Enter key
    entry.bind('<Return>', on_ok)
    entry.bind('<KP_Enter>', on_ok)  # Numeric keypad Enter
    
    # Keep window on top initially
    root.after(200, lambda: root.attributes('-topmost', False))
    
    # Wait for window
    root.mainloop()
    
    if result["cancelled"]:
        return None
    
    return result["value"]

def parse_row_numbers(row_string, total_rows):
    """
    Parse row number string into a list of row indices
    Supports:  "1,3,5", "1-5", "1,3,5-10,15"
    Returns list of 0-based indices
    """
    if not row_string or not row_string.strip():
        return None
    
    row_indices = set()
    
    # Remove spaces
    row_string = row_string.replace(' ', '')
    
    # Split by comma
    parts = row_string.split(',')
    
    try:
        for part in parts: 
            if '-' in part:
                # Range
                start, end = part.split('-')
                start = int(start)
                end = int(end)
                
                if start < 1 or end > total_rows:
                    raise ValueError(f"Row numbers must be between 1 and {total_rows}")
                if start > end:
                    raise ValueError(f"Invalid range: {part} (start must be <= end)")
                
                # Add range (convert to 0-based)
                for i in range(start - 1, end):
                    row_indices.add(i)
            else:
                # Single number
                num = int(part)
                if num < 1 or num > total_rows:
                    raise ValueError(f"Row number {num} is out of range (1-{total_rows})")
                row_indices.add(num - 1)  # Convert to 0-based
    
    except ValueError as e:
        messagebox.showerror(
            "Invalid Input",
            f"❌ ERROR PARSING ROW NUMBERS\n\n{str(e)}\n\nPlease try again."
        )
        return None
    
    # Convert to sorted list
    return sorted(list(row_indices))

def filter_recipients(recipients, row_indices):
    """Filter recipients list to only include specified row indices"""
    if row_indices is None:
        return recipients
    
    filtered = [recipients[i] for i in row_indices if i < len(recipients)]
    
    # Show confirmation
    row_numbers_display = ', '.join([str(i+1) for i in row_indices[: 10]])
    if len(row_indices) > 10:
        row_numbers_display += f", ... ({len(row_indices)} total)"
    
    messagebox.showinfo(
        "Row Selection Confirmed",
        f"✅ SELECTED ROWS CONFIRMED\n\n"
        f"Processing {len(filtered)} recipient(s)\n"
        f"Row numbers: {row_numbers_display}\n\n"
        f"Click OK to continue..."
    )
    
    return filtered

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
        return f"{tin[: 2]}-{tin[2:]}"
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
        print(f"⚠️ Field '{field_name}' not found in master template")
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
    print(f"Mode:  {mode}")
    print(f"{'='*80}\n")
    
    # Create PDF
    c = canvas.Canvas(output_pdf, pagesize=letter)

    # Process each recipient (3 per page)
    for page_num, i in enumerate(range(0, len(recipients), 3)):
        page_recipients = recipients[i:i+3]
    
        print(f"Page {page_num + 1}:  Processing {len(page_recipients)} sections")
    
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
            print(f"  Section {section_num + 1}: {recipient_name}")
            
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
    
    print(f"\n✓ Created:  {output_pdf}")
    print(f"✓ Total pages: {(len(recipients) + 2) // 3}")
    print(f"{'='*80}\n")
    
    # Show completion message
    messagebox.showinfo(
        "Success!",
        f"✅ PDF CREATED SUCCESSFULLY!\n\n"
        f"File: {os.path.basename(output_pdf)}\n"
        f"Location: {os.path.dirname(output_pdf)}\n\n"
        f"Recipients processed: {len(recipients)}\n"
        f"Total pages:  {(len(recipients) + 2) // 3}\n\n"
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
        "  2. Which rows to process (all or specific)\n"
        "  3. JSON field positions file (coordinates)\n"
        f"  {'4. Background image file (for alignment)' if USE_BACKGROUND_IMAGE else ''}\n"
        f"  {'5' if USE_BACKGROUND_IMAGE else '4'}. Output PDF location\n\n"
        "Click OK to begin..."
    )
    
    # Step 1: Select CSV file
    CSV_FILE = select_csv_file()
    print(f"✓ CSV file selected: {CSV_FILE}")
    
    # Load CSV to get row count
    with open(CSV_FILE, 'r') as f:
        reader = csv.DictReader(f)
        all_recipients = list(reader)
    
    total_rows = len(all_recipients)
    print(f"✓ CSV loaded: {total_rows} total rows")
    
    # Step 2: Ask if user wants all or some rows
    process_all = select_all_or_some()
    
    selected_indices = None
    if not process_all:
        # User wants specific rows
        while True:
            row_input = get_row_numbers(total_rows)
            
            if row_input is None: 
                # User cancelled
                messagebox.showwarning("Cancelled", "Operation cancelled by user.")
                exit()
            
            selected_indices = parse_row_numbers(row_input, total_rows)
            
            if selected_indices is not None:
                break  # Valid input received
            # If invalid, loop will repeat with error message shown
    
    # Filter recipients if specific rows selected
    if selected_indices: 
        recipients_to_process = filter_recipients(all_recipients, selected_indices)
        print(f"✓ Selected {len(recipients_to_process)} specific rows")
    else:
        recipients_to_process = all_recipients
        print(f"✓ Processing all {len(recipients_to_process)} rows")
    
    # Step 3: Select JSON file
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
            'x':  MASTER_FIELDS['BOX 6']['x'],
            'y': MASTER_FIELDS['BOX 6']['y'] - HALF_LINE_OFFSET
        }
    if 'BOX 7' in MASTER_FIELDS:
        MASTER_FIELDS['BOX 7a'] = {
            'x':  MASTER_FIELDS['BOX 7']['x'],
            'y': MASTER_FIELDS['BOX 7']['y'] - HALF_LINE_OFFSET
        }
    
    # Step 4: Select background image (only in dev mode)
    BACKGROUND_IMAGE_PATH = None
    if USE_BACKGROUND_IMAGE: 
        BACKGROUND_IMAGE_PATH = select_background_image()
        if BACKGROUND_IMAGE_PATH: 
            print(f"✓ Background image selected: {BACKGROUND_IMAGE_PATH}")
    
    # Step 5: Select output location
    mode_name = "DEV" if USE_BACKGROUND_IMAGE else "PROD"
    OUTPUT_PDF = select_output_location(mode_name)
    print(f"✓ Output location selected: {OUTPUT_PDF}")
    
    # Create temporary CSV with selected rows only
    temp_csv = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', newline='')
    
    with open(CSV_FILE, 'r') as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        
        writer = csv.DictWriter(temp_csv, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(recipients_to_process)
    
    temp_csv.close()
    
    # Process with temp file
    fill_1099_nec_form(temp_csv.name, OUTPUT_PDF, BACKGROUND_IMAGE_PATH)
    
    # Clean up temp file
    os.unlink(temp_csv.name)