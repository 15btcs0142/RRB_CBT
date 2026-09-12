import os
import sys
import shutil

# Add workspace directory to path
script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import image_handler

def test_rdkit_molecule_generation():
    print("\n--- TEST 1: RDKit Molecular Structure Generation ---")
    out_dir = os.path.join(script_dir, "test_output")
    os.makedirs(out_dir, exist_ok=True)
    
    # 1. By SMILES (Benzene)
    path_benzene = os.path.join(out_dir, "test_benzene.png")
    res1 = image_handler.generate_molecule_structure("c1ccccc1", path_benzene)
    assert res1 and os.path.exists(res1), "Failed to generate benzene from SMILES"
    assert os.path.getsize(res1) > 0, "Benzene PNG is empty"
    print("[PASS] SMILES (c1ccccc1) generated successfully:", res1, f"({os.path.getsize(res1)} bytes)")
    
    # 2. By Common NCERT Name (Ethanol)
    path_ethanol = os.path.join(out_dir, "test_ethanol.png")
    res2 = image_handler.generate_molecule_structure("ethanol", path_ethanol)
    assert res2 and os.path.exists(res2), "Failed to generate ethanol from compound name"
    print("[PASS] Compound Name ('ethanol') generated successfully:", res2, f"({os.path.getsize(res2)} bytes)")

    # 3. By Complex NCERT Name (Glucose)
    path_glucose = os.path.join(out_dir, "test_glucose.png")
    res3 = image_handler.generate_molecule_structure("glucose", path_glucose)
    assert res3 and os.path.exists(res3), "Failed to generate glucose from compound name"
    print("[PASS] Compound Name ('glucose') generated successfully:", res3, f"({os.path.getsize(res3)} bytes)")

    # 4. By Pharmaceutical NCERT Name (Aspirin)
    path_aspirin = os.path.join(out_dir, "test_aspirin.png")
    res4 = image_handler.generate_molecule_structure("aspirin", path_aspirin)
    assert res4 and os.path.exists(res4), "Failed to generate aspirin from compound name"
    print("[PASS] Compound Name ('aspirin') generated successfully:", res4, f"({os.path.getsize(res4)} bytes)")

    # 5. Invalid input handling
    res_inv = image_handler.generate_molecule_structure("invalid_xyz_chemical_compound_12345")
    assert res_inv is None, "Invalid input should return None"
    print("[PASS] Invalid compound safely returned None without crashing.")

def test_wikimedia_fetch():
    print("\n--- TEST 2: Wikimedia Commons API Fetching ---")
    out_dir = os.path.join(script_dir, "test_output")
    os.makedirs(out_dir, exist_ok=True)

    path_heart = os.path.join(out_dir, "test_heart.png")
    res_heart = image_handler.fetch_from_wikimedia("human heart diagram", path_heart, min_width=600)
    assert res_heart and os.path.exists(res_heart), "Failed to fetch heart diagram from Wikimedia Commons"
    assert os.path.getsize(res_heart) > 0, "Heart diagram PNG is empty"
    print("[PASS] Wikimedia search for 'human heart diagram' downloaded successfully:", res_heart, f"({os.path.getsize(res_heart)} bytes)")

    # Test invalid / non-existent search
    res_none = image_handler.fetch_from_wikimedia("zxqj98274982347923847293847928374928374")
    assert res_none is None, "Non-existent search should gracefully return None"
    print("[PASS] Non-existent search query safely returned None.")

def test_image_cache():
    print("\n--- TEST 3: Image Hash Cache Verification ---")
    out_dir = os.path.join(script_dir, "test_output")
    
    # 1. Molecule Cache
    cache_path_before = image_handler._get_cached_image("ethanol_CCO", prefix='mol', ext='png')
    assert cache_path_before and os.path.exists(cache_path_before), "Molecule should be cached from Test 1"
    print("[PASS] Molecule cache hit verified at:", cache_path_before)

    # 2. Wikimedia Cache
    cache_path_wiki = image_handler._get_cached_image("human heart diagram", prefix='wiki', ext='png')
    assert cache_path_wiki and os.path.exists(cache_path_wiki), "Wikimedia diagram should be cached from Test 2"
    print("[PASS] Wikimedia cache hit verified at:", cache_path_wiki)

def test_pipeline_and_docx():
    print("\n--- TEST 4: Full Pipeline & DOCX Generation ---")
    out_dir = os.path.join(script_dir, "test_output")
    docx_path = os.path.join(out_dir, "test_generated_paper.docx")

    dummy_paper_data = {
        "class": "10",
        "subject": "Science",
        "sections": [
            {
                "section_label": "Section A",
                "section_title": "Multiple Choice Questions",
                "marks_per_question": 1,
                "instruction": "Select the correct option.",
                "questions": [
                    {
                        "number": 1,
                        "question": "What is the equivalent resistance across the circuit shown below?",
                        "options": ["(a) 30 Ω", "(b) 15 Ω", "(c) 20 Ω", "(d) 10 Ω"],
                        "needs_image": True,
                        "image_category": "circuit",
                        "circuit_type": "series_resistors",
                        "params": {"voltage": "12V", "r1": "10Ω", "r2": "20Ω"}
                    },
                    {
                        "number": 2,
                        "question": "Identify the functional groups present in the given molecular structure of aspirin:",
                        "options": ["(a) Carboxylic acid and ester", "(b) Aldehyde and ketone", "(c) Alcohol and ether", "(d) Amine and amide"],
                        "needs_image": True,
                        "image_category": "chemistry",
                        "compound_name": "aspirin",
                        "smiles": "CC(=O)Oc1ccccc1C(=O)O"
                    },
                    {
                        "number": 3,
                        "question": "Which part of the human heart pumps oxygenated blood to the rest of the body?",
                        "options": ["(a) Left Ventricle", "(b) Right Ventricle", "(c) Left Atrium", "(d) Right Atrium"],
                        "needs_image": True,
                        "image_category": "wikimedia",
                        "image_keywords": ["human heart diagram cross section"]
                    },
                    {
                        "number": 4,
                        "question": "Define Ohm's Law and state its SI unit of resistance.",
                        "options": [],
                        "needs_image": False,
                        "image_category": "none"
                    }
                ]
            }
        ]
    }

    # Process Images
    processed = image_handler.process_question_images(dummy_paper_data, output_dir=out_dir)
    
    q1 = processed['sections'][0]['questions'][0]
    q2 = processed['sections'][0]['questions'][1]
    q3 = processed['sections'][0]['questions'][2]
    q4 = processed['sections'][0]['questions'][3]

    assert q1.get('image_path') and os.path.exists(q1['image_path']), "Q1 circuit image missing"
    assert q2.get('image_path') and os.path.exists(q2['image_path']), "Q2 chemistry molecule image missing"
    assert q3.get('image_path') and os.path.exists(q3['image_path']), "Q3 wikimedia image missing"
    assert not q4.get('image_path'), "Q4 should not have image"

    print("[PASS] Pipeline processed all questions and assigned image paths successfully.")

    # Create DOCX
    meta = {
        "school_name": "RRB Central School",
        "school_address": "Knowledge Campus, New Delhi",
        "academic_session": "2024-25",
        "exam_type": "Unit Test 1",
        "class": "10",
        "subject": "Science",
        "duration": "1.5 Hours",
        "max_marks": "40",
        "teacher_name": "Dr. Sharma"
    }
    
    res_docx = image_handler.create_paper_docx(processed, meta, docx_path)
    assert res_docx and os.path.exists(res_docx), "Failed to create docx"
    print("[PASS] Native DOCX question paper created successfully at:", res_docx, f"({os.path.getsize(res_docx)} bytes)")

if __name__ == "__main__":
    test_rdkit_molecule_generation()
    test_wikimedia_fetch()
    test_image_cache()
    test_pipeline_and_docx()
    print("\n==========================================")
    print("ALL 4 TEST SUITES PASSED SUCCESSFULLY!")
    print("==========================================")

