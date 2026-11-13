from flask import Flask, request, jsonify, send_file
import requests
import tempfile
import os
from report_engine import process_smartmoving_reports

app = Flask(__name__)

@app.route('/generate-report', methods=['POST'])
def generate_report():
    data = request.json
    period_label = data.get('period_label', 'Weekly / Monthly Period')
    file_urls = data.get('files', {})

    temp_files = {}
    try:
        # Download the 5 Excel files passed by Zapier
        for key, url in file_urls.items():
            if url:
                resp = requests.get(url)
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                tmp.write(resp.content)
                tmp.close()
                temp_files[key] = tmp.name

        result = process_smartmoving_reports(temp_files, period_label=period_label)
        
        # Return weekly text and a URL or base64 for the PDF
        pdf_filename = os.path.basename(result['pdf_path'])
        return jsonify({
            "status": "success",
            "weekly_report_text": result['weekly_text'],
            "pdf_download_url": f"{request.host_url}download/{pdf_filename}"
        })
    finally:
        for path in temp_files.values():
            if os.path.exists(path):
                os.remove(path)

@app.route('/download/<filename>', methods=['GET'])
def download_pdf(filename):
    return send_file(f"/tmp/{filename}", as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
