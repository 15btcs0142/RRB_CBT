import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import app, save_paper_to_repository, download_shared_paper

paper_data = {
    'sections': [
        {
            'section_label': 'Section A',
            'section_title': 'Short Answer Questions',
            'questions': [
                {'number': '1', 'question': 'Solve $\\frac{x}{2} = 5$'}
            ]
        }
    ]
}

abs_path, doc_filename, rel_link = save_paper_to_repository(paper_data, '10', 'A', 'Mathematics', 'Exam', 'TestTeacher')
print("Saved file abs path:", abs_path)
print("Saved rel_link:", rel_link)
print("File exists?", os.path.exists(abs_path))

# Test download route handling with relative link
with app.test_client() as client:
    with client.session_transaction() as sess:
        sess['admin_logged_in'] = True

    # Test downloading using rel_link
    dl_url = f"/api/shared_papers/download/{rel_link}"
    res = client.get(dl_url)
    print("Download response status:", res.status_code)
    print("Download content length:", len(res.data))

    # Test downloading with legacy 'paper/' prefix
    legacy_url = f"/api/shared_papers/download/paper/{rel_link}"
    res_legacy = client.get(legacy_url)
    print("Legacy Download response status:", res_legacy.status_code)
