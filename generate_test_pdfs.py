import os
import random
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def create_invoice(filename, vendor, date_str, amount_label, amount, layout_type):
    filepath = os.path.join("data", "input_pdfs", filename)
    c = canvas.Canvas(filepath, pagesize=letter)
    width, height = letter

    # Default starting positions
    y_pos = height - 50

    if layout_type == 7: # Missing Vendor Test
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, y_pos, "INVOICE")
    else:
        c.setFont("Helvetica-Bold", 20)
        if layout_type == 6: # Vendor on the right
            c.drawString(width - 250, y_pos, f"Vendor: {vendor}")
        else:
            c.drawString(50, y_pos, vendor)
    
    y_pos -= 50
    c.setFont("Helvetica", 12)
    
    # Date Variations
    c.drawString(50, y_pos, f"Invoice Date: {date_str}")
    y_pos -= 50

    # Body / Noise
    c.drawString(50, y_pos, "Description: Web Development Services")
    y_pos -= 20
    c.drawString(50, y_pos, "Hours: 40 @ $50/hr")
    y_pos -= 20
    
    if layout_type == 5: # Noise Maker
        c.drawString(50, y_pos, "Note: Previous balance of $500.00 paid on 2026-01-01.")
        y_pos -= 20
        c.drawString(50, y_pos, "Reference Number: 99482-112")
        y_pos -= 20

    # Total Amount Variations
    c.setFont("Helvetica-Bold", 14)
    if layout_type == 4: # Top-Heavy (Total at the top)
        c.drawString(width - 200, height - 50, f"{amount_label} ${amount}")
    else:
        # Standard bottom-ish placement
        c.drawString(50, y_pos - 40, f"{amount_label} ${amount}")

    c.save()
    print(f"Generated: {filename}")

if __name__ == "__main__":
    # Ensure directory exists
    os.makedirs(os.path.join("data", "input_pdfs"), exist_ok=True)

    print("Generating 7 Test Invoices...\n")
    
    # 1. Standard
    create_invoice("test_01_standard.pdf", "Acme Corp", "2026-05-30", "Total:", "1250.00", 1)
    
    # 2. Different Date Format & Keyword
    create_invoice("test_02_dateshift.pdf", "TechFlow Inc", "30/05/2026", "Amount Due:", "4500.50", 2)
    
    # 3. Wordy Date
    create_invoice("test_03_wordy.pdf", "Global Solutions", "May 30, 2026", "Grand Total:", "899.99", 3)
    
    # 4. Total at the Top
    create_invoice("test_04_topheavy.pdf", "Skyline Design", "2026/05/30", "Total Payable:", "2100.00", 4)
    
    # 5. Noise / Decoys
    create_invoice("test_05_noise.pdf", "DataSystems LLC", "05-30-2026", "Total:", "3300.00", 5)
    
    # 6. Vendor on Right Side
    create_invoice("test_06_rightvendor.pdf", "Creative Media", "2026-05-30", "Balance:", "150.00", 6)
    
    # 7. Missing Vendor entirely
    create_invoice("test_07_novendor.pdf", "", "2026-05-30", "Total:", "600.00", 7)

    print("\nDone! Check your data/input_pdfs folder.")