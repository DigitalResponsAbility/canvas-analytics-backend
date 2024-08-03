# Canvas Analytics Backend App

This Flask application analyzes Canvas course data using Python scripts and serves the results via RESTful endpoints.

## Project Setup

### Prerequisites

- Python 3.x
- Virtual environment tool (e.g., `venv`, `virtualenv`)
- Canvas API credentials (API URL, API key, Course ID)

### Local Setup

1. Clone the repository
2. Navigate to the project directory
3. Setup virtual environment:

    ```bash
    python -m venv venv
    ```
4. Activate the virtual environment:
    - On Windows:
        ```bash
        venv\Scripts\activate
        ```
    - On macOS/Linux:
        ```bash
        source venv/bin/activate
        ```
5. Install dependencies:

    ```bash
    pip install -r requirements.txt
    ```

6. Configure Canvas API credentials (TO BE REMOVED):

    - Open Course_Completion.py and enrollment.py.
    - Replace api_url, api_key, and course_id variables with your Canvas API details.

7. Start the Flask application:

    ```bash
    python app.py
    ```
8. Access the endpoints (Replace Api Key and Course Id as 972187):

    - `http://127.0.0.1:5000/get_course_completion_report?api_url=https://utah.instructure.com:443/api/v1&api_key=apiKey&course_id=courseId`: Executes Course_Completion.py and provides a download link for the generated Excel file.
    - `http://127.0.0.1:5000/get_enrollment_report?api_url=https://utah.instructure.com:443/api/v1&api_key=apiKey&course_id=courseId`: Executes enrollment.py and provides a download link for the generated Excel file.

    
    


### Usage

- Click on the endpoints to trigger the respective Python scripts for data analysis.
- The Flask app will generate Excel files containing analysis results, which can be downloaded via the provided links.

### Future

Frontend app to integrate these URLs. 