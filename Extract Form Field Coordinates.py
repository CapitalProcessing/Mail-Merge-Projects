"""Extract Form Field Coordinates from PDF - With Kids Support"""
import PyPDF2
import json

def extract_field_info(pdf_path):
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        
        # Get page size
        page = reader.pages[0]
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        print(f"Page size: {page_width} × {page_height} points")
        print(f"Page size: {page_width/72:.2f}\" × {page_height/72:.2f}\"\n")
        
        # Get form fields
        fields = reader. get_fields()
        
        if not fields:
            print("No form fields found!")
            return
        
        print(f"Found {len(fields)} form fields:\n")
        print("="*100)
        
        field_data = {}
        
        for field_name, field_info in fields.items():
            print(f"\nField: {field_name}")
            
            # Check if field has /Kids (child widgets)
            if '/Kids' in field_info: 
                print(f"  Has {len(field_info['/Kids'])} child widgets")
                # Get first child widget
                try: 
                    kid = field_info['/Kids'][0].get_object()
                    if '/Rect' in kid: 
                        rect = kid['/Rect']
                        x = float(rect[0])
                        y = float(rect[1])
                        
                        # ONLY x and y - no width/height
                        field_data[field_name] = {
                            'x': round(x, 2),
                            'y': round(y, 2)
                        }
                        
                        print(f"  ✓ Position from child widget:  X={x:.2f}, Y={y:.2f}")
                        print(f"  (X in inches: {x/72:.3f}\", Y in inches:  {y/72:.3f}\")")
                    else: 
                        print(f"  ⚠️  Child widget has no /Rect")
                except Exception as e: 
                    print(f"  ⚠️  Error accessing child widget: {e}")
            
            # Also check direct /Rect
            elif '/Rect' in field_info:
                rect = field_info['/Rect']
                x = float(rect[0])
                y = float(rect[1])
                
                # ONLY x and y - no width/height
                field_data[field_name] = {
                    'x': round(x, 2),
                    'y':  round(y, 2)
                }
                
                print(f"  ✓ Position:  X={x:.2f}, Y={y:.2f}")
                print(f"  (X in inches: {x/72:.3f}\", Y in inches: {y/72:.3f}\")")
            
            # Fallback:  search page annotations
            else: 
                print(f"  Searching page annotations...")
                found = False
                for page_num, page in enumerate(reader.pages):
                    if '/Annots' in page:
                        for annot in page['/Annots']: 
                            annot_obj = annot.get_object()
                            if '/T' in annot_obj and annot_obj['/T'] == field_name: 
                                if '/Rect' in annot_obj: 
                                    rect = annot_obj['/Rect']
                                    x = float(rect[0])
                                    y = float(rect[1])
                                    
                                    # ONLY x and y - no width/height
                                    field_data[field_name] = {
                                        'x': round(x, 2),
                                        'y': round(y, 2)
                                    }
                                    
                                    print(f"  ✓ Found in page annotations:")
                                    print(f"  Position: X={x:.2f}, Y={y:.2f}")
                                    print(f"  (X in inches: {x/72:.3f}\", Y in inches: {y/72:.3f}\")")
                                    found = True
                                    break
                    if found:
                        break
                
                if not found: 
                    print(f"  ❌ No coordinates found")
            
            print("-"*100)
        
        # Save to JSON
        output_file = pdf_path.replace('.pdf', '_fields.json')
        with open(output_file, 'w') as out:
            json.dump({
                'page_size': {'width':  page_width, 'height':  page_height},
                'fields': field_data
            }, out, indent=2)
        
        print(f"\n✓ Field data saved to: {output_file}")
        print(f"✓ Extracted coordinates for {len(field_data)} out of {len(fields)} fields")

if __name__ == "__main__":
    pdf_path = input("Enter path to fillable PDF: ").strip('"')
    extract_field_info(pdf_path)