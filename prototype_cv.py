import cv2
import numpy as np
import os
from docx2pdf import convert
import fitz  # PyMuPDF

# ==============================================================================
# STAGE 2: THE BUSINESS LOGIC (ADJUSTED INK PRICES)
# ==============================================================================


def calculate_price_options(coverage_percentage):
    """
    Calculates prices using the 4-tier model.
    Paper Price is fixed at 1.00.
    Ink Price is adjusted to match the target totals.
    """
    PAPER_PRICE = 1.00

    tier_name = "N/A"
    bw_ink_price = 0.00
    color_ink_price = 0.00

    # --- 4-Tier Pricing Logic (Adjusted for Paper Price of 1.00) ---

    # Tier 1: 0% - 25% (Target Totals: P2 / P3)
    if 0 <= coverage_percentage <= 25.0:
        tier_name = "Tier 1: Light (0-25%)"
        bw_ink_price = 1.00
        color_ink_price = 2.00

    # Tier 2: 26% - 50% (Target Totals: P3 / P5)
    elif 25.0 < coverage_percentage <= 50.0:
        tier_name = "Tier 2: Medium (26-50%)"
        bw_ink_price = 2.00
        color_ink_price = 4.00

    # Tier 3: 51% - 75% (Target Totals: P4 / P7)
    elif 50.0 < coverage_percentage <= 75.0:
        tier_name = "Tier 3: Heavy (51-75%)"
        bw_ink_price = 3.00
        color_ink_price = 6.00

    # Tier 4: 76% - 100% (Target Totals: P5 / P10)
    elif coverage_percentage > 75.0:
        tier_name = "Tier 4: Dense/Full (76-100%)"
        bw_ink_price = 4.00
        color_ink_price = 9.00

    # Calculate Finals
    final_bw_price = int(PAPER_PRICE + bw_ink_price)
    final_color_price = int(PAPER_PRICE + color_ink_price)

    return {
        "paper_price": int(PAPER_PRICE),
        "tier_name": tier_name,
        "bw_ink_price": int(bw_ink_price),
        "color_ink_price": int(color_ink_price),
        "final_bw_price": final_bw_price,
        "final_color_price": final_color_price
    }

# ==============================================================================
# STAGE 1: THE TECHNICAL PIPELINE (ANALYSIS IN RAM)
# ==============================================================================


def perform_technical_analysis(open_cv_image):
    try:
        # Check if color exists by subtracting channels
        b, g, r = cv2.split(open_cv_image)
        original_doc_type = "Color" if np.count_nonzero(g - r) > 0 else "B&W"

        # Convert to grayscale and threshold to isolate ink/content
        gray_image = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
        _, thresholded_image = cv2.threshold(
            gray_image, 240, 255, cv2.THRESH_BINARY_INV)

        content_pixels = cv2.countNonZero(thresholded_image)

        # Multiply Height (index 0) by Width (index 1)
        total_pixels = thresholded_image.shape[0] * thresholded_image.shape[1]

        coverage_percentage = (content_pixels / total_pixels) * 100

        return {
            "success": True,
            "original_doc_type": original_doc_type,
            "coverage_percentage": coverage_percentage
        }

    except Exception as e:
        print(f"An unexpected error occurred during analysis: {e}")
        return {"success": False}


# ==============================================================================
# MAIN PIPELINE: DOCX -> PDF -> IMAGES -> ANALYSIS
# ==============================================================================

def process_document(docx_path):
    print("\n===================================================")
    print(f"STARTING PIPELINE FOR: {docx_path}")
    print("===================================================")

    if not os.path.exists(docx_path):
        print(f"ERROR: File '{docx_path}' not found.")
        return

    # STEP 1: Convert DOCX to PDF
    pdf_path = docx_path.replace(".docx", ".pdf")
    print("\n Converting DOCX to PDF...")
    try:
        convert(docx_path, pdf_path)
    except Exception as e:
        print(
            f"Failed to convert DOCX to PDF. Ensure MS Word is closed. Error: {e}")
        return

    # STEP 2: Open the PDF and prepare for Grand Totals
    print(f"\n Opening PDF and extracting pages...")
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)

    grand_total_bw = 0
    grand_total_color = 0

    print(f"\n Analyzing {total_pages} page(s)...\n")

    # STEP 3: Loop through each page, render to image, and analyze
    for page_num in range(total_pages):
        page = pdf_document.load_page(page_num)

        # Render PDF page to an image in memory (200 DPI is the sweet spot)
        pix = page.get_pixmap(dpi=200)

        # Convert the raw pixels to a format OpenCV understands
        img_array = np.frombuffer(
            pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        # Handle Color Spaces (RGB vs RGBA)
        if pix.n == 4:
            open_cv_image = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
        else:
            open_cv_image = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

        # Pass the image directly to the analyzer
        tech_results = perform_technical_analysis(open_cv_image)

        # Ensure the analysis was successful
        if not tech_results or not tech_results.get("success"):
            print(f"Failed to analyze Page {page_num + 1}")
            continue

        # Extract Data
        coverage = tech_results["coverage_percentage"]
        doc_type = tech_results["original_doc_type"]

        # Calculate Pricing for this specific page
        pricing = calculate_price_options(coverage)

        # Extract Prices
        paper_price = pricing["paper_price"]
        tier_name = pricing["tier_name"]

        bw_ink = pricing["bw_ink_price"]
        final_bw = pricing["final_bw_price"]

        col_ink = pricing["color_ink_price"]
        final_col = pricing["final_color_price"]

        # Add to Grand Totals
        grand_total_bw += final_bw
        grand_total_color += final_col

        # --- Print Results for the Current Page ---
        print(f"--- PAGE {page_num + 1} OF {total_pages} ---")
        print(f"Detected Type: {doc_type}")
        print(f"Ink Coverage : {coverage:.2f}%")
        print(f"Assigned Tier: {tier_name}")
        print(
            f"  -> B&W   : Paper {paper_price} + Ink {bw_ink} = PHP {final_bw}.00")
        print(
            f"  -> Color : Paper {paper_price} + Ink {col_ink} = PHP {final_col}.00\n")

    # Clean up the PDF object
    pdf_document.close()

    print(f"NOTE: PDF file kept at: {pdf_path}")

    # --- Print Grand Totals ---
    print("===================================================")
    print("                GRAND TOTAL ESTIMATE               ")
    print("===================================================")
    print(f"Total Pages Processed: {total_pages}")
    print(f"TOTAL IF PRINTED IN B&W   : PHP {grand_total_bw}.00")
    print(f"TOTAL IF PRINTED IN COLOR : PHP {grand_total_color}.00")
    print("===================================================\n")


# ==============================================================================
# --- --- --- --- --- --- SIMPLE TEST --- --- --- --- --- ---
# ==============================================================================
if __name__ == '__main__':
    # Place a sample Word document in the same folder as this script
    test_docx_file = 'test_document.docx'

    if not os.path.exists(test_docx_file):
        print(
            f"Warning: '{test_docx_file}' not found. Please create a dummy Word file named '{test_docx_file}' in this folder to run the test.")
    else:
        process_document(test_docx_file)
