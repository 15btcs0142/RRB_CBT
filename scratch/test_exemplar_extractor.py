import os
import sys
import json

# Setup workspace path
script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import fitz  # PyMuPDF
import exemplar_extractor
import image_handler

def create_sample_exemplar_pdf(output_pdf_path):
    """
    Programmatically create a realistic NCERT Exemplar sample chapter PDF
    with MCQs, Short Answer questions, and embedded diagram illustration.
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
    doc = fitz.open()

    # --- Page 1: Chapter Header + MCQs with a diagram ---
    page1 = doc.new_page(width=595, height=842) # A4

    # Header
    page1.insert_text((50, 40), "NCERT EXEMPLAR PROBLEMS — CLASS X", fontsize=11, fontname="helv", color=(0.2, 0.2, 0.5))
    page1.insert_text((50, 60), "CHAPTER 1: Chemical Reactions and Equations", fontsize=14, fontname="hebo", color=(0.1, 0.1, 0.3))
    page1.draw_line((50, 70), (545, 70), color=(0.7, 0.7, 0.7), width=1)

    # Question 1 (MCQ)
    q1_text = (
        "1. Which of the following is not a physical change?\n"
        "   (a) Boiling of water to give water vapour\n"
        "   (b) Melting of ice to give water\n"
        "   (c) Dissolution of salt in water\n"
        "   (d) Combustion of Liquefied Petroleum Gas (LPG)"
    )
    page1.insert_text((50, 95), q1_text, fontsize=10, fontname="helv", color=(0, 0, 0))

    # Question 2 (MCQ with diagram)
    q2_text = (
        "2. Study the reaction apparatus shown in the figure below. In which test tube\n"
        "   will the iron nails undergo maximum rusting?\n"
        "   (a) Test tube A (containing air and water)\n"
        "   (b) Test tube B (containing boiled distilled water and oil)\n"
        "   (c) Test tube C (containing dry air and anhydrous CaCl2)\n"
        "   (d) Rusting occurs equally in all three test tubes"
    )
    page1.insert_text((50, 175), q2_text, fontsize=10, fontname="helv", color=(0, 0, 0))

    # Embed a sample test diagram image on Page 1 (e.g. from static/images/ncert if available, or generated)
    sample_img_source = os.path.join(rrb_dir, "static", "images", "ncert", "bio_reflex_arc.png")
    if os.path.exists(sample_img_source):
        img_rect = fitz.Rect(180, 260, 420, 390)
        page1.insert_image(img_rect, filename=sample_img_source)
        page1.insert_text((220, 405), "Figure 1.1: Experimental Setup", fontsize=9, fontname="helv", color=(0.3, 0.3, 0.3))

    # Question 3 (MCQ)
    q3_text = (
        "3. Which among the following is (are) double displacement reaction(s)?\n"
        "   (i) Pb + CuCl2 -> PbCl2 + Cu\n"
        "   (ii) Na2SO4 + BaCl2 -> BaSO4 + 2NaCl\n"
        "   (iii) C + O2 -> CO2\n"
        "   (iv) CH4 + 2O2 -> CO2 + 2H2O\n"
        "   (a) (i) and (iv)\n"
        "   (b) (ii) only\n"
        "   (c) (i) and (ii)\n"
        "   (d) (iii) and (iv)"
    )
    page1.insert_text((50, 430), q3_text, fontsize=10, fontname="helv", color=(0, 0, 0))

    # --- Page 2: Short and Long Answer Questions ---
    page2 = doc.new_page(width=595, height=842)
    page2.insert_text((50, 40), "NCERT EXEMPLAR PROBLEMS — CLASS X SCIENCE", fontsize=11, fontname="helv", color=(0.2, 0.2, 0.5))
    page2.insert_text((50, 60), "Short Answer Questions", fontsize=13, fontname="hebo", color=(0.1, 0.1, 0.3))
    page2.draw_line((50, 70), (545, 70), color=(0.7, 0.7, 0.7), width=1)

    # Question 4 (Short Answer)
    q4_text = (
        "4. Why do we store silver chloride in dark coloured bottles? Write the chemical equation\n"
        "   involved in this phenomenon."
    )
    page2.insert_text((50, 95), q4_text, fontsize=10, fontname="helv", color=(0, 0, 0))

    # Question 5 (Short Answer)
    q5_text = (
        "5. A magnesium ribbon is burnt in oxygen to give a white compound X accompanied by\n"
        "   emission of light. If the burning ribbon is now placed in an atmosphere of nitrogen, it\n"
        "   continues to burn and forms a compound Y.\n"
        "   (a) Write the chemical formulae of X and Y.\n"
        "   (b) Write a balanced chemical equation, when X is dissolved in water."
    )
    page2.insert_text((50, 150), q5_text, fontsize=10, fontname="helv", color=(0, 0, 0))

    doc.save(output_pdf_path)
    doc.close()
    print(f"[PASS] Created realistic NCERT Exemplar sample PDF at: {output_pdf_path}")
    return output_pdf_path


def run_all_exemplar_tests():
    print("\n=======================================================")
    print("RUNNING NCERT EXEMPLAR EXTRACTOR TEST SUITE")
    print("=======================================================\n")


    test_dir = os.path.join(script_dir, "test_output")
    os.makedirs(test_dir, exist_ok=True)
    sample_pdf_path = os.path.join(test_dir, "sample_ncert_exemplar_ch1.pdf")
    test_index_path = os.path.join(test_dir, "test_exemplar_index.json")

    # Step 1: Create Sample PDF
    create_sample_exemplar_pdf(sample_pdf_path)

    # Step 2: Run Master Driver process_exemplar_pdf
    print("\n--- TEST 1: Process Exemplar PDF Driver ---")
    chapter_entry = exemplar_extractor.process_exemplar_pdf(
        pdf_path=sample_pdf_path,
        subject="Science",
        class_num="10",
        chapter_name="Chemical Reactions and Equations",
        index_path=test_index_path
    )

    assert chapter_entry, "process_exemplar_pdf returned None"
    assert chapter_entry.get("subject") == "Science", "Subject mismatch in index"
    assert chapter_entry.get("class") == "10", "Class mismatch in index"
    assert len(chapter_entry.get("questions", [])) >= 3, f"Expected at least 3 questions, got {len(chapter_entry.get('questions', []))}"
    print(f"[PASS] Successfully processed Exemplar PDF with {len(chapter_entry.get('questions', []))} questions and {len(chapter_entry.get('images', []))} images.")

    # Step 3: Verify Index Persistence & Querying
    print("\n--- TEST 2: Query Exemplar Bank by Subject/Class/Type ---")
    chapters = exemplar_extractor.get_exemplar_chapters(subject="Science", class_num="10", index_path=test_index_path)
    assert len(chapters) >= 1, "Failed to retrieve exemplar chapters list"
    print(f"[PASS] get_exemplar_chapters returned: {chapters[0]['chapter']} (Total: {chapters[0]['total_questions']} questions)")

    mcqs = exemplar_extractor.get_exemplar_questions("Science", "10", question_type="MCQ", count=5, index_path=test_index_path)
    assert len(mcqs) >= 1, "Failed to retrieve MCQs from exemplar bank"
    print(f"[PASS] get_exemplar_questions retrieved {len(mcqs)} MCQs.")

    # Step 4: Paper Generation from Exemplar Bank
    print("\n--- TEST 3: Generate Structured Paper Data & Native DOCX ---")
    paper_data = exemplar_extractor.generate_paper_from_exemplar(
        subject="Science",
        class_num="10",
        chapter_name="Chemical Reactions and Equations",
        mcq_count=2,
        short_count=2,
        long_count=1,
        index_path=test_index_path
    )

    assert paper_data.get("sections"), "Generated paper_data has no sections"
    assert len(paper_data["sections"]) >= 1, "Paper sections empty"
    print(f"[PASS] generate_paper_from_exemplar created paper structure with {len(paper_data['sections'])} sections.")

    # Process Images & Generate DOCX
    docx_output = os.path.join(test_dir, "exemplar_generated_paper.docx")
    meta = {
        "school_name": "RRB Central School",
        "school_address": "Exemplar Testing Campus",
        "academic_session": "2024-25",
        "exam_type": "NCERT Exemplar Test",
        "class": "10",
        "subject": "Science",
        "duration": "1 Hour",
        "max_marks": "25",
        "teacher_name": "Gaurav Shukla"
    }

    image_handler.process_question_images(paper_data, output_dir=test_dir)
    res_docx = image_handler.create_paper_docx(paper_data, meta, docx_output)

    assert res_docx and os.path.exists(res_docx), "Failed to generate native docx"
    assert os.path.getsize(res_docx) > 5000, f"DOCX is too small: {os.path.getsize(res_docx)} bytes"
    print(f"[PASS] Native DOCX generated from Exemplar Bank at: {res_docx} ({os.path.getsize(res_docx)} bytes)")

    print("\n=======================================================")
    print("ALL NCERT EXEMPLAR EXTRACTOR TESTS PASSED SUCCESSFULLY!")
    print("=======================================================\n")


if __name__ == "__main__":
    run_all_exemplar_tests()
