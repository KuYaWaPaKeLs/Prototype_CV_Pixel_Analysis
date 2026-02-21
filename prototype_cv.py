import cv2
import numpy as np
import os
import time
from docx2pdf import convert
import fitz  # PyMuPDF
import win32com.client  # <--- NEW: Used to talk directly to Windows

# ==============================================================================
# STAGE 2: THE BUSINESS LOGIC (PAPER + INK MODEL)
# ==============================================================================


def calculate_price_options(coverage_percentage):
    PAPER_PRICE = 1.00
    tier_name = "N/A"
    bw_ink_price = 0.00
    color_ink_price = 0.00

    if 0 <= coverage_percentage <= 25.0:
        tier_name = "Tier 1: Light (0-25%)"
        bw_ink_price = 1.00
        color_ink_price = 2.00
    elif 25.0 < coverage_percentage <= 50.0:
        tier_name = "Tier 2: Medium (26-50%)"
        bw_ink_price = 2.00
        color_ink_price = 4.00
    elif 50.0 < coverage_percentage <= 75.0:
        tier_name = "Tier 3: Heavy (51-75%)"
        bw_ink_price = 3.00
        color_ink_price = 6.00
    elif coverage_percentage > 75.0:
        tier_name = "Tier 4: Dense/Full (76-100%)"
        bw_ink_price = 4.00
        color_ink_price = 9.00

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
# STAGE 1: THE TECHNICAL PIPELINE
# ==============================================================================


def perform_technical_analysis(open_cv_image):
    try:
        b, g, r = cv2.split(open_cv_image)
        original_doc_type = "Color" if np.count_nonzero(g - r) > 0 else "B&W"

        gray_image = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
        _, thresholded_image = cv2.threshold(
            gray_image, 240, 255, cv2.THRESH_BINARY_INV)

        content_pixels = cv2.countNonZero(thresholded_image)
        total_pixels = thresholded_image.shape[0] * thresholded_image.shape[1]
        coverage_percentage = (content_pixels / total_pixels) * 100

        return {
            "success": True,
            "original_doc_type": original_doc_type,
            "coverage_percentage": coverage_percentage
        }

    except Exception as e:
        return {"success": False}

# ==============================================================================
# MAIN PROCESSING FUNCTION
# ==============================================================================


def process_document(docx_path):
    print("\n===================================================")
    print(f"STARTING ANALYSIS FOR: {docx_path}")
    print("===================================================")

    if not os.path.exists(docx_path):
        print(f"ERROR: File '{docx_path}' not found.")
        return 0, 0

    # 1. Convert to PDF
    pdf_path = docx_path.replace(".docx", ".pdf")
    print(" Converting DOCX to PDF...")

    try:
        convert(docx_path, pdf_path)
    except Exception as e:
        print(f"Conversion failed: {e}")
        return 0, 0

    # 2. Open PDF
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)
    grand_total_bw = 0
    grand_total_color = 0

    print(f" Analyzing {total_pages} page(s)...\n")

    # 3. Analyze Pages
    for page_num in range(total_pages):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap(dpi=200)
        img_array = np.frombuffer(
            pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        if pix.n == 4:
            open_cv_image = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
        else:
            open_cv_image = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

        tech_results = perform_technical_analysis(open_cv_image)

        if not tech_results or not tech_results.get("success"):
            continue

        coverage = tech_results["coverage_percentage"]
        doc_type = tech_results["original_doc_type"]
        pricing = calculate_price_options(coverage)

        # Extract Pricing
        paper_price = pricing["paper_price"]
        tier_name = pricing["tier_name"]
        bw_ink = pricing["bw_ink_price"]
        final_bw = pricing["final_bw_price"]
        col_ink = pricing["color_ink_price"]
        final_col = pricing["final_color_price"]

        grand_total_bw += final_bw
        grand_total_color += final_col

        print(f"--- PAGE {page_num + 1} ---")
        print(f"Type: {doc_type} | Ink: {coverage:.2f}% | {tier_name}")
        print(
            f"  -> B&W   : Paper {paper_price} + Ink {bw_ink} = PHP {final_bw}.00")
        print(
            f"  -> Color : Paper {paper_price} + Ink {col_ink} = PHP {final_col}.00\n")

    pdf_document.close()
    print(f"NOTE: PDF file kept at: {pdf_path}")

    return grand_total_bw, grand_total_color


# ==============================================================================
# THE EXECUTION BLOCK (SMART WARM UP + TIMER)
# ==============================================================================
if __name__ == '__main__':
    test_docx_file = 'test_document.docx'

    if os.path.exists(test_docx_file):

        # --- PHASE 1: WARM UP MS WORD (THE SMART WAY) ---
        print("\n[SYSTEM] Waking up MS Word engine in the background...")
        try:
            # This directly launches Word into memory without opening any files!
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Quit()  # Close it immediately
            print("[SYSTEM] Word is awake and cached in RAM.")
        except Exception as e:
            print(f"[SYSTEM] Warm-up skipped or failed: {e}")

        time.sleep(1)  # Give Windows 1 second to release the process

        # --- PHASE 2: REAL TIMED RUN ---
        print("\n[TIMER STARTED] Processing the document...")

        # Start the high-precision timer
        start_time = time.perf_counter()

        # Run the full pipeline
        total_bw, total_color = process_document(test_docx_file)

        # Stop the high-precision timer
        end_time = time.perf_counter()

        # Calculate duration
        total_duration = end_time - start_time

        # --- PHASE 3: PRINT FINAL RESULTS ---
        print("\n===================================================")
        print("                GRAND TOTAL ESTIMATE               ")
        print("===================================================")
        print(f"TOTAL IF PRINTED IN B&W   : PHP {total_bw}.00")
        print(f"TOTAL IF PRINTED IN COLOR : PHP {total_color}.00")
        print("===================================================")
        print("                 PERFORMANCE REPORT                ")
        print("===================================================")
        print(f"Time Taken: {total_duration:.4f} seconds")
        print("===================================================\n")

    else:
        print(f"Warning: '{test_docx_file}' not found.")
