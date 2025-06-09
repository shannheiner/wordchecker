# app.py - Enhanced Bold1 + Margin Checker
from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
from docx import Document
import io
import os
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

# Azure Storage Configuration - Using Environment Variables (Secure)
AZURE_STORAGE_CONNECTION_STRING = os.getenv('AZURE_CONNECTION_STRING')
AZURE_ACCOUNT_NAME = os.getenv('AZURE_ACCOUNT_NAME') 
AZURE_ACCOUNT_KEY = os.getenv('AZURE_ACCOUNT_KEY')
CONTAINER_NAME = "practice-files"

# Check if credentials are loaded
if not all([AZURE_STORAGE_CONNECTION_STRING, AZURE_ACCOUNT_NAME, AZURE_ACCOUNT_KEY]):
    print("❌ Error: Azure credentials not found in environment variables")
    print("Run the setup commands in terminal first!")

# Initialize Azure Blob client
blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)

def check_document_formatting(doc):
    """Enhanced function to check Bold1 and top margin"""
    results = {
        "bold_check": {"found": False, "is_bold": False, "score": 0},
        "margin_check": {"correct": False, "actual_margin": 0, "score": 0},
        "total_score": 0,
        "max_score": 20,  # 10 for bold + 10 for margin
        "message": "",
        "debug_info": []
    }
    
    try:
        # Check Bold1 formatting
        for paragraph_num, paragraph in enumerate(doc.paragraphs):
            paragraph_text = paragraph.text.strip()
            
            if "Bold1" in paragraph_text or "bold1" in paragraph_text.lower():
                results["bold_check"]["found"] = True
                results["debug_info"].append(f"Found 'Bold1' in paragraph {paragraph_num}")
                
                # Check if any run containing Bold1 is bold
                for run_num, run in enumerate(paragraph.runs):
                    text = run.text.strip()
                    is_bold = bool(run.bold) if run.bold is not None else False
                    results["debug_info"].append(f"  Run {run_num}: '{text}' (Bold: {is_bold})")
                    
                    if "Bold1" in text or "bold1" in text.lower():
                        results["bold_check"]["is_bold"] = is_bold
                break
        
        # Check top margin (0.5 inches)
        if doc.sections:
            section = doc.sections[0]  # First section
            top_margin_inches = round(section.top_margin.inches, 2)
            results["margin_check"]["actual_margin"] = top_margin_inches
            
            # Check if margin is approximately 0.5 inches (allow 0.1 inch tolerance)
            target_margin = 0.5
            tolerance = 0.1
            margin_correct = abs(top_margin_inches - target_margin) <= tolerance
            results["margin_check"]["correct"] = margin_correct
            
            results["debug_info"].append(f"Top margin: {top_margin_inches} inches (target: 0.5 inches)")
        else:
            results["debug_info"].append("No document sections found")
        
        # Calculate scores
        if results["bold_check"]["found"] and results["bold_check"]["is_bold"]:
            results["bold_check"]["score"] = 10
        
        if results["margin_check"]["correct"]:
            results["margin_check"]["score"] = 10
        
        results["total_score"] = results["bold_check"]["score"] + results["margin_check"]["score"]
        
        # Create message
        bold_msg = "✅ Bold1 correct" if results["bold_check"]["is_bold"] else "❌ Bold1 not bold"
        margin_msg = f"✅ Top margin correct (0.5\")" if results["margin_check"]["correct"] else f"❌ Top margin is {results['margin_check']['actual_margin']}\" (should be 0.5\")"
        
        results["message"] = f"{bold_msg} | {margin_msg}"
        
        return results
        
    except Exception as e:
        return {
            "bold_check": {"found": False, "is_bold": False, "score": 0},
            "margin_check": {"correct": False, "actual_margin": 0, "score": 0},
            "total_score": 0,
            "max_score": 20,
            "message": f"❌ Error analyzing document: {str(e)}",
            "debug_info": [f"Error: {str(e)}"]
        }

def create_practice_file():
    """Create a simple practice document and upload to Azure"""
    try:
        # Create practice document
        doc = Document()
        
        # Add title
        title = doc.add_paragraph("Formatting Practice Document")
        title.runs[0].bold = True
        title.runs[0].font.size = 177800  # 14pt in EMUs
        
        # Add instructions
        doc.add_paragraph("")
        doc.add_paragraph("Instructions:")
        doc.add_paragraph("1. Make the text 'Bold1' below bold")
        doc.add_paragraph("2. Set the top margin to 0.5 inches")
        doc.add_paragraph("3. Save and upload for checking")
        doc.add_paragraph("")
        
        # Add the test text (NOT bold - student must make it bold)
        test_paragraph = doc.add_paragraph("Bold1")
        # Explicitly ensure it's NOT bold
        test_paragraph.runs[0].bold = False
        
        doc.add_paragraph("")
        doc.add_paragraph("How to set margins:")
        doc.add_paragraph("• Go to Layout tab → Margins → Custom Margins")
        doc.add_paragraph("• Set Top margin to 0.5 inches")
        doc.add_paragraph("• Click OK")
        
        # Save to memory
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        # Upload to Azure
        blob_client = blob_service_client.get_blob_client(
            container=CONTAINER_NAME,
            blob="formatting-practice.docx"
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
            blob_name="formatting-practice.docx",
            account_key=AZURE_ACCOUNT_KEY,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.utcnow() + timedelta(hours=1)
        )
        
        download_url = f"https://{AZURE_ACCOUNT_NAME}.blob.core.windows.net/{CONTAINER_NAME}/formatting-practice.docx?{sas_token}"
        return download_url
        
    except Exception as e:
        print(f"Error generating download URL: {e}")
        return None

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Word Formatting Checker</title>
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
        .score-display {
            font-size: 24px;
            text-align: center;
            margin: 20px 0;
            padding: 15px;
            border-radius: 8px;
        }
        .breakdown {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 15px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📝 Word Formatting Checker</h1>
        <p>Test your Word document formatting skills - Bold text and margin settings!</p>
        
        <div class="step">
            <h3>Step 1: Download Practice File</h3>
            <p>Download the practice document that contains formatting tasks.</p>
            <button onclick="downloadPracticeFile()">📁 Download Practice File</button>
        </div>
        
        <div class="step">
            <h3>Step 2: Apply Required Formatting</h3>
            <ul>
                <li>Open the downloaded file in Microsoft Word</li>
                <li>Find the text "Bold1" and make it <strong>bold</strong> (Ctrl+B)</li>
                <li><strong>Set top margin to 0.5 inches:</strong>
                    <ul>
                        <li>Layout tab → Margins → Custom Margins</li>
                        <li>Set "Top" to 0.5"</li>
                        <li>Click OK</li>
                    </ul>
                </li>
                <li>Save the document</li>
            </ul>
        </div>
        
        <div class="step">
            <h3>Step 3: Upload for Checking</h3>
            <div class="upload-area">
                <input type="file" id="fileInput" accept=".docx" style="display: none;" onchange="uploadFile()">
                <button onclick="document.getElementById('fileInput').click()">📤 Upload Document</button>
                <p>Select your completed .docx file</p>
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
            
            document.getElementById('results').innerHTML = '<p>🔍 Analyzing document formatting...</p>';
            
            fetch('/check-formatting', {
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
            const resultClass = data.score === data.max_score ? 'success' : 'error';
            
            let html = `<div class="results ${resultClass}">`;
            html += `<h3>📊 Formatting Analysis Results</h3>`;
            
            // Score display
            html += `<div class="score-display" style="background: ${data.score === data.max_score ? '#d4edda' : '#f8d7da'};">`;
            html += `<strong>Score: ${data.score}/${data.max_score} (${data.percentage}%)</strong>`;
            html += `</div>`;
            
            html += `<p><strong>${data.message}</strong></p>`;
            
            // Detailed breakdown
            html += `<div class="breakdown">`;
            html += `<h4>📋 Detailed Results:</h4>`;
            html += `<ul>`;
            
            // Bold formatting
            if (data.found_bold1) {
                html += `<li>${data.is_bold ? '✅' : '❌'} Bold1 formatting: ${data.is_bold ? 'BOLD' : 'NOT BOLD'} (${data.is_bold ? '10' : '0'}/10 points)</li>`;
            } else {
                html += `<li>❌ Bold1 text: NOT FOUND (0/10 points)</li>`;
            }
            
            // Margin check
            html += `<li>${data.margin_correct ? '✅' : '❌'} Top margin: ${data.actual_margin}" ${data.margin_correct ? '(Correct!)' : '(Should be 0.5")'} (${data.margin_correct ? '10' : '0'}/10 points)</li>`;
            
            html += `</ul>`;
            html += `</div>`;
            
            // Grade
            let grade = 'F';
            if (data.percentage >= 90) grade = 'A';
            else if (data.percentage >= 80) grade = 'B';
            else if (data.percentage >= 70) grade = 'C';
            else if (data.percentage >= 60) grade = 'D';
            
            html += `<p style="text-align: center; font-size: 18px;"><strong>Grade: ${grade}</strong></p>`;
            
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

@app.route('/check-formatting', methods=['POST'])
def check_formatting():
    """Check Bold1 and margin formatting in uploaded document"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Read the uploaded file
        file_stream = io.BytesIO(file.read())
        doc = Document(file_stream)
        
        # Check formatting (Bold1 + margins)
        results = check_document_formatting(doc)
        
        # Format response for frontend
        response = {
            "score": results["total_score"],
            "max_score": results["max_score"],
            "percentage": round((results["total_score"] / results["max_score"]) * 100, 1),
            "message": results["message"],
            "found_bold1": results["bold_check"]["found"],
            "is_bold": results["bold_check"]["is_bold"],
            "margin_correct": results["margin_check"]["correct"],
            "actual_margin": results["margin_check"]["actual_margin"],
            "debug_info": results["debug_info"]
        }
        
        return jsonify(response)
        
    except Exception as e:
        return jsonify({
            'error': 'Failed to analyze document',
            'message': str(e),
            'score': 0,
            'max_score': 20,
            'found_bold1': False,
            'is_bold': False,
            'margin_correct': False,
            'actual_margin': 0
        }), 500

if __name__ == '__main__':
    print("🚀 Starting Enhanced Word Formatting Checker...")
    print("📝 First visit: http://localhost:5000/setup")
    print("📝 Then visit: http://localhost:5000")
    print("⚙️  Make sure to set Azure environment variables!")
    app.run(debug=True, host='0.0.0.0', port=5000)