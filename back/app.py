from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
from docx import Document
from playwright.sync_api import sync_playwright
from datetime import datetime
import os, json, uuid, html as html_module

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
DATA_FOLDER = os.path.join(BASE_DIR, 'data')
EXPORT_FOLDER = os.path.join(BASE_DIR, 'exports')
FRONTEND_PATH = os.path.join(os.path.dirname(BASE_DIR), 'front')

for folder in [UPLOAD_FOLDER, DATA_FOLDER, EXPORT_FOLDER]:
    os.makedirs(folder, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ALLOWED_EXTENSIONS = {'docx', 'doc'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_data_path(doc_id):
    return os.path.join(DATA_FOLDER, f"{doc_id}.json")

def load_doc(doc_id):
    path = get_data_path(doc_id)
    if not os.path.exists(path):
        return None
    with open(path, 'r') as f:
        return json.load(f)

def save_doc(doc_id, data):
    data['updated_at'] = datetime.now().isoformat()
    with open(get_data_path(doc_id), 'w') as f:
        json.dump(data, f, indent=2)

def extract_text_from_docx(filepath):
    doc = Document(filepath)
    content = []
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            content.append({'id': f'p-{i}', 'type': 'paragraph', 'text': p.text})
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                if cell.text.strip():
                    content.append({'id': f't-{ti}-{ri}-{ci}', 'type': 'table-cell', 'text': cell.text})
    return content

def build_styled_html(content, styles):
    styles_by_para = {}
    for s in styles:
        styles_by_para.setdefault(s['paraIndex'], []).append(s)

    paragraphs = []
    for i, para in enumerate(content):
        text = para['text']
        para_styles = styles_by_para.get(i, [])

        if not para_styles:
            paragraphs.append(f'<p>{html_module.escape(text)}</p>')
            continue

        bounds = sorted({0, len(text)} | {max(0, min(s[k], len(text))) for s in para_styles for k in ('startOffset', 'endOffset')})
        parts = ''
        for j in range(len(bounds) - 1):
            start, end = bounds[j], bounds[j + 1]
            seg = text[start:end]
            if not seg:
                continue
            active = [s for s in para_styles if s['startOffset'] <= start and s['endOffset'] >= end]
            if not active:
                parts += html_module.escape(seg)
            else:
                classes = ['styled-text'] + [s['type'] for s in active]
                inline = []
                for s in active:
                    t, c = s['type'], s['color']
                    if t == 'highlight': inline.append(f'background-color:{c}')
                    elif t == 'textcolor': inline.append(f'color:{c}')
                    elif t in ('border', 'circle'): inline.append(f'border-color:{c}')
                    elif t in ('underline', 'strikethrough'): inline.append(f'text-decoration-color:{c}')
                style_attr = f' style="{";".join(inline)}"' if inline else ''
                parts += f'<span class="{" ".join(classes)}"{style_attr}>{html_module.escape(seg)}</span>'
        paragraphs.append(f'<p>{parts}</p>')

    return f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;line-height:1.5;color:#1a1a1a;background:#fafafa;-webkit-font-smoothing:antialiased}}
.document-container{{max-width:800px;margin:24px auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.05)}}
.document-content{{padding:32px 48px;font-size:15px;line-height:1.8;color:#1a1a1a}}
.document-content p{{margin-bottom:1em;position:relative}}
.document-content p:last-child{{margin-bottom:0}}
.styled-text{{position:relative;display:inline}}
.styled-text.bold{{font-weight:700}}
.styled-text.italic{{font-style:italic}}
.styled-text.underline{{text-decoration:underline;text-decoration-thickness:2px;text-underline-offset:2px}}
.styled-text.strikethrough{{text-decoration:line-through;text-decoration-thickness:2px}}
.styled-text.highlight{{padding:0 2px;border-radius:2px}}
.styled-text.border{{border:2px solid;border-radius:3px;padding:0 4px;margin:0 2px}}
.styled-text.circle{{border:2px solid;border-radius:100px;padding:0 6px;margin:0 2px}}
</style>
</head>
<body>
<div class="document-container"><div class="document-content">{''.join(paragraphs)}</div></div>
</body>
</html>'''

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy", "message": "Document Typography API is running"})

@app.route('/api/upload', methods=['POST'])
def upload_document():
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Only Word documents (.docx) are allowed"}), 400

    doc_id = str(uuid.uuid4())
    original_filename = secure_filename(file.filename)
    file_path = os.path.join(UPLOAD_FOLDER, f"{doc_id}.docx")
    file.save(file_path)

    try:
        content = extract_text_from_docx(file_path)
    except Exception as e:
        os.remove(file_path)
        return jsonify({"error": f"Failed to parse document: {str(e)}"}), 400

    doc_data = {
        "doc_id": doc_id, "original_filename": original_filename,
        "created_at": datetime.now().isoformat(), "updated_at": datetime.now().isoformat(),
        "content": content, "styles": [], "enrichment_log": []
    }
    save_doc(doc_id, doc_data)

    return jsonify({"success": True, "doc_id": doc_id, "filename": original_filename, "content": content, "message": "Document uploaded successfully"})

@app.route('/api/document/<doc_id>', methods=['GET'])
def get_document(doc_id):
    doc = load_doc(doc_id)
    return jsonify(doc) if doc else (jsonify({"error": "Document not found"}), 404)

@app.route('/api/document/<doc_id>/styles', methods=['POST'])
def save_styles(doc_id):
    doc = load_doc(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    doc['styles'] = request.get_json().get('styles', [])
    save_doc(doc_id, doc)
    return jsonify({"success": True, "message": "Styles saved successfully", "updated_at": doc['updated_at']})

@app.route('/api/document/<doc_id>/log', methods=['POST'])
def log_enrichment_action(doc_id):
    doc = load_doc(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    entry = request.get_json()
    if not entry or 'action' not in entry or 'timestamp' not in entry:
        return jsonify({"error": "Invalid log entry"}), 400
    if entry['action'] not in ('add', 'delete', 'clear'):
        return jsonify({"error": "Invalid action type"}), 400
    doc.setdefault('enrichment_log', []).append(entry)
    save_doc(doc_id, doc)
    return jsonify({"success": True, "message": "Action logged", "log_count": len(doc['enrichment_log'])})

@app.route('/api/document/<doc_id>/export', methods=['POST'])
def export_document(doc_id):
    doc = load_doc(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404

    styles = request.get_json().get('styles', [])
    doc['styles'] = styles
    save_doc(doc_id, doc)

    base_name = os.path.splitext(doc['original_filename'])[0]
    html_path = os.path.join(EXPORT_FOLDER, f"{doc_id}_temp.html")
    export_filename = f"{base_name}_enriched.png"
    export_path = os.path.join(EXPORT_FOLDER, f"{doc_id}_{export_filename}")

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(build_styled_html(doc['content'], styles))

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={'width': 900, 'height': 800})
            page.goto(f'file://{html_path}')
            page.wait_for_load_state('networkidle')
            page.wait_for_timeout(500)
            page.screenshot(path=export_path, full_page=True)
            browser.close()
    finally:
        if os.path.exists(html_path):
            os.remove(html_path)

    return jsonify({"success": True, "message": "Document exported successfully", "export_filename": export_filename, "download_url": f"/api/document/{doc_id}/download"})

@app.route('/api/document/<doc_id>/download', methods=['GET'])
def download_document(doc_id):
    doc = load_doc(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    base_name = os.path.splitext(doc['original_filename'])[0]
    export_filename = f"{base_name}_enriched.png"
    export_path = os.path.join(EXPORT_FOLDER, f"{doc_id}_{export_filename}")
    if not os.path.exists(export_path):
        return jsonify({"error": "Export not found. Please save the document first."}), 404
    return send_file(export_path, as_attachment=True, download_name=export_filename)

@app.route('/api/documents', methods=['GET'])
def list_documents():
    docs = []
    for f in os.listdir(DATA_FOLDER):
        if f.endswith('.json'):
            with open(os.path.join(DATA_FOLDER, f), 'r') as file:
                d = json.load(file)
                docs.append({"doc_id": d['doc_id'], "filename": d['original_filename'], "created_at": d['created_at'], "updated_at": d['updated_at']})
    return jsonify(docs)

@app.route('/api/document/<doc_id>', methods=['DELETE'])
def delete_document(doc_id):
    for path in [os.path.join(UPLOAD_FOLDER, f"{doc_id}.docx"), get_data_path(doc_id)]:
        if os.path.exists(path):
            os.remove(path)
    return jsonify({"success": True, "message": "Document deleted successfully"})

@app.route('/')
def serve_frontend():
    return send_from_directory(FRONTEND_PATH, 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(FRONTEND_PATH, filename)

if __name__ == '__main__':
    print("=" * 50)
    print("Server running at: http://localhost:5001")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=5001)
