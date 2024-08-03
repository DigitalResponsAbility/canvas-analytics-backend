from flask import Flask, send_file, request
from flask_cors import CORS 
import os
import subprocess

app = Flask(__name__)
CORS(app)  # Apply CORS to your Flask app

@app.route('/get_course_completion_report', methods=['GET'])
def get_course_completion_report():
    api_url = request.args.get('api_url')
    api_key = request.args.get('api_key')
    course_id = request.args.get('course_id')
    
    if not (api_url and api_key and course_id):
        return "Missing parameters: api_url, api_key, course_id", 400
    
    script_path = 'course_completion.py'
    result = subprocess.run(['python', script_path, api_url, api_key, course_id], capture_output=True, text=True)
    
    if result.returncode == 0:
        return send_file('canvas_course_completion_analysis.xlsx', as_attachment=True)
    else:
        return result.stderr, 500

@app.route('/get_enrollment_report', methods=['GET'])
def get_enrollment_report():
    api_url = request.args.get('api_url')
    api_key = request.args.get('api_key')
    course_id = request.args.get('course_id')
    
    if not (api_url and api_key and course_id):
        return "Missing parameters: api_url, api_key, course_id", 400
    
    script_path = 'enrollment.py'
    result = subprocess.run(['python', script_path, api_url, api_key, course_id], capture_output=True, text=True)
    
    if result.returncode == 0:
        return send_file('canvas_enrollments_analysis.xlsx', as_attachment=True)
    else:
        return result.stderr, 500

if __name__ == '__main__':
    app.run(debug=True)
