# app.py - Simple Bold1 Checker
from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
from docx import Document
import io
import os
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

# Azure Storage Configuration - REPLACE WITH YOUR ACTUAL VALUES
AZURE_STORAGE_CONNECTION_STRING = “DefaultEndpointsProtocol=https;AccountName=smartlearn;AccountKey=xm2gYLKStgXzNXOB7cuhDvC/Kyqt1Uzd4TMHIT4trfRbUpcBd3dMPPA1Ct80ZpM1/On7F4O3zoxs+AStV3G/Tg==;EndpointSuffix=core.windows.net”

AZURE_ACCOUNT_NAME = "smartlearn"
AZURE_ACCOUNT_KEY ="xm2gYLKStgXzNXOB7cuhDvC/Kyqt1Uzd4TMHIT4trfRbUpcBd3dMPPA1Ct80ZpM1/On7F4O3zoxs+AStV3G/Tg==”

CONTAINER_NAME = "practice-files"

# Initialize Azure Blob client
blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)

def check_bold1_formatting(doc):
    """Simple function to check if 'Bold1' text is bold"""
    results = {
        "found_bold1": False,
        "is_bold": False,
        "score": 0,
        "message": "",
        "debug_info": []
    }
    
    try:
        # Search through all paragraphs and runs
        for paragraph_num, paragraph in enumerate(doc.paragraphs):
            for run_num, run in enumerate(paragraph.runs):
                text = run.text.strip()
                results["debug_info"].append(f"Para {paragraph_num}, Run {run_num}: '{text}' (Bold: {run.bold})")
                
                if "Bold1" in text:
                    results["found_bold1"] = True
                    results["is_bold"] = bool(run.bold)
                    break
            
            if results["found_bold1"]:
                break
        
        # Determine score and message
        if results["found_bold1"] and results["is_bold"]:
            results["score"] = 100
            results["message"] = "✅ Perfect! 'Bold1' found and correctly formatted as bold!"
        elif results["found_bold1"] and not results["is_bold"]:
            results["score"] = 0
            results["message"] = "❌ 'Bold1' found but it's not bold. Please make it bold."
        else:
            results["score"] = 0
            results["message"] = "❌ 'Bold1' text not found in document. Please add it."
        
        return results
        
    except Exception as e:
        return {
            "found_bold1": False,
            "is_bold": False, 
            "score": 0,
            "message": f"❌ Error analyzing document: {str(e)}",
            "debug_info": [f"Error: {str(e)}"]
        }

def create_practice_file():
    """Create a simple practice document and upload to Azure"""
    try:
        # Create practice document
        doc = Document()
        
        # Add title
        title = doc.add_paragraph("Bold Formatting Practice")
        title.runs[0].bold = True
        title.runs[0].font.size = 177800  # 14pt in EMUs
        
        # Add instructions
        doc.add_paragraph("")
        doc.add_paragraph("Instructions: Make the text 'Bold1' below bold, then upload for checking.")
        doc.add_paragraph("")
        
        # Add the test text (NOT bold - student must make it bold)
        test_paragraph = doc.add_paragraph("Bold1")
        # Explicitly ensure it's NOT bold
        test_paragraph.runs[0].bold = False
        
        doc.add_paragraph("")
        doc.add_paragraph("Save this document and upload it to the checker website.")
        
        # Save to memory
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        # Upload to Azure
        blob_client = blob_service_client.get_blob_client(
            container=CONTAINER_NAME,
            blob="bold-practice.docx"
        )
        
        blob_client.upload_blob(doc_stream.getvalue(), overwrite=True)
        print("✅ Practice file uploaded to Azure!")
        return True
        
    except Exception as e:
        print(f"❌ Error creating practice file: {e}")
        return False

def get_practice_file_download_url():
    """Generate download URL for practice file"""
    try:
        # Generate SAS token for download
        sas_token = generate_blob_sas(
            account_name=AZURE_ACCOUNT_NAME,
            container_name=CONTAINER_NAME,
            blob_name="bold-practice.docx",
            account_key=AZURE_ACCOUNT_KEY,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.utcnow() + timedelta(hours=1)
        )
        
        download_url = f"https://{AZURE_ACCOUNT_NAME}.blob.core.windows.net/{CONTAINER_NAME}/bold-practice.docx?{sas_token}"
        return download_url
        
    except Exception as e:
        print(f"Error generating download URL: {e}")
        return None

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Bold1 Checker - Test</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .step {
            background: #f8f9fa;
            margin: 20px 0;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #007bff;
        }
        button {
            background: #007bff;
            color: white;
            padding: 12px 24px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            margin: 10px 5px;
        }
        button:hover {
            background: #0056b3;
        }
        .upload-area {
            border: 2px dashed #ccc;
            padding: 40px;
            text-align: center;
            border-radius: 8px;
            margin: 20px 0;
        }
        .results {
            margin-top: 20px;
            padding: 20px;
            border-radius: 8px;
        }
        .success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .error {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        .debug {
            background: #e2e3e5;
            border: 1px solid #d6d8db;
            color: #383d41;
            font-family: monospace;
            font-size: 12px;
            max-height: 200px;
            overflow-y: auto;
            margin-top: 10px;
            padding: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🧪 Bold1 Checker - Test</h1>
        <p>Simple test to check if "Bold1" text is bold in a Word document.</p>
        
        <div class="step">
            <h3>Step 1: Download Practice File</h3>
            <p>Download the test document that contains "Bold1" text (not yet bold).</p>
            <button onclick="downloadPracticeFile()">📁 Download Practice File</button>
        </div>
        
        <div class="step">
            <h3>Step 2: Make "Bold1" Bold</h3>
            <ul>
                <li>Open the downloaded file in Microsoft Word</li>
                <li>Find the text "Bold1"</li>
                <li>Select it and make it <strong>bold</strong> (Ctrl+B)</li>
                <li>Save the document</li>
            </ul>
        </div>
        
        <div class="step">
            <h3>Step 3: Upload for Testing</h3>
            <div class="upload-area">
                <input type="file" id="fileInput" accept=".docx" style="display: none;" onchange="uploadFile()">
                <button onclick="document.getElementById('fileInput').click()">📤 Upload Document</button>
                <p>Select your modified .docx file</p>
            </div>
            <div id="results"></div>
        </div>
    </div>
    
    <script>
        function downloadPracticeFile() {
            fetch('/download-practice')
                .then(response => response.json())
                .then(data => {
                    if (data.download_url) {
                        // Open download in new tab
                        window.open(data.download_url, '_blank');
                        alert('Practice file download started! Check your downloads folder.');
                    } else {
                        alert('Error: ' + (data.error || 'Could not generate download link'));
                    }
                })
                .catch(error => {
                    alert('Error downloading file: ' + error);
                });
        }
        
        function uploadFile() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];
            
            if (!file) {
                alert('Please select a file');
                return;
            }
            
            if (!file.name.endsWith('.docx')) {
                alert('Please select a .docx file');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', file);
            
            document.getElementById('results').innerHTML = '<p>🔍 Checking if "Bold1" is bold...</p>';
            
            fetch('/check-bold', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                displayResults(data);
            })
            .catch(error => {
                document.getElementById('results').innerHTML = 
                    '<div class="results error"><p>❌ Error: ' + error + '</p></div>';
            });
        }
        
        function displayResults(data) {
            const resultsDiv = document.getElementById('results');
            const resultClass = data.score === 100 ? 'success' : 'error';
            
            let html = `<div class="results ${resultClass}">`;
            html += `<h3>📊 Test Results</h3>`;
            html += `<h2>Score: ${data.score}/100</h2>`;
            html += `<p><strong>${data.message}</strong></p>`;
            
            if (data.found_bold1) {
                html += `<p>✅ Found "Bold1" text</p>`;
                html += `<p>${data.is_bold ? '✅' : '❌'} Bold formatting: ${data.is_bold ? 'YES' : 'NO'}</p>`;
            } else {
                html += `<p>❌ "Bold1" text not found</p>`;
            }
            
            // Debug information
            if (data.debug_info && data.debug_info.length > 0) {
                html += `<details><summary>🔍 Debug Info (click to expand)</summary>`;
                html += `<div class="debug">`;
                data.debug_info.forEach(info => {
                    html += `<div>${info}</div>`;
                });
                html += `</div></details>`;
            }
            
            html += `</div>`;
            resultsDiv.innerHTML = html;
        }
    </script>
</body>
</html>
"""

# Routes
@app.route('/')
def home():
    """Main page"""
    return HTML_TEMPLATE

@app.route('/setup')
def setup():
    """Setup route - run this first to create and upload practice file"""
    try:
        if create_practice_file():
            return "✅ Setup complete! Practice file uploaded to Azure. Go to <a href='/'>home page</a> to test."
        else:
            return "❌ Setup failed. Check your Azure credentials in the code."
    except Exception as e:
        return f"❌ Setup error: {str(e)}"

@app.route('/download-practice')
def download_practice():
    """Provide download URL for practice file"""
    try:
        download_url = get_practice_file_download_url()
        if download_url:
            return jsonify({"download_url": download_url})
        else:
            return jsonify({"error": "Could not generate download URL"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/check-bold', methods=['POST'])
def check_bold():
    """Check if Bold1 is bold in uploaded document"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Read the uploaded file
        file_stream = io.BytesIO(file.read())
        doc = Document(file_stream)
        
        # Check Bold1 formatting
        results = check_bold1_formatting(doc)
        
        return jsonify(results)
        
    except Exception as e:
        return jsonify({
            'error': 'Failed to analyze document',
            'message': str(e),
            'score': 0,
            'found_bold1': False,
            'is_bold': False
        }), 500

if __name__ == '__main__':
    print("🚀 Starting Bold1 Checker...")
    print("📝 First visit: http://localhost:5000/setup")
    print("📝 Then visit: http://localhost:5000")
    print("⚙️  Make sure to update Azure credentials in the code!")
    app.run(debug=True, host='0.0.0.0', port=5000)
