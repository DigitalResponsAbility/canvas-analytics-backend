import requests
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

class CourseCompletionAnalyzer:
    def __init__(self, api_url, api_key, course_id):
        self.api_url = api_url
        self.api_key = api_key
        self.course_id = course_id
        self.modules = None
        self.students = None
        self.completion_data = None
        self.student_module_completion = None

    def fetch_course_modules(self):
        """
        Fetch list of modules in the course from Canvas API.
        """
        url = f"{self.api_url}/courses/{self.course_id}/modules"
        headers = {"Authorization": f"Bearer {self.api_key}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        self.modules = response.json()

    def fetch_students_in_course(self):
        """
        Fetch list of students enrolled in the course from Canvas API.
        """
        url = f"{self.api_url}/courses/{self.course_id}/students"
        headers = {"Authorization": f"Bearer {self.api_key}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        self.students = response.json()

    def get_module_items_for_student(self, module_id, student_id):
        """
        Retrieve items (including completion data) for a specific module in a course for a given student.
        """
        url = f"{self.api_url}/courses/{self.course_id}/modules/{module_id}/items"
        params = {
            "access_token": self.api_key,
            "student_id": student_id
        }
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()

    def analyze_module_completion(self):
        """
        Analyze module completion data for the course.
        """
        self.fetch_course_modules()
        self.fetch_students_in_course()

        self.completion_data = {}
        self.student_module_completion = {student['name']: {module['name']: False for module in self.modules} for student in self.students}

        for module in self.modules:
            module_id = module['id']
            module_name = module['name']
            completed_count = 0

            for student in self.students:
                student_id = student['id']
                
                try:
                    module_items = self.get_module_items_for_student(module_id, student_id)
                except requests.exceptions.HTTPError as e:
                    print(f"Error fetching module items for student {student_id}: {e}")
                    continue
                
                for item in module_items:
                    if item['completion_requirement']['completed']:
                        completed_count += 1
                        self.student_module_completion[student['name']][module_name] = True
                        break
            
            total_count = len(self.students)
            completion_percentage = (completed_count / total_count) * 100 if total_count > 0 else 0

            self.completion_data[module_name] = {
                'completed_count': completed_count,
                'total_count': total_count,
                'completion_percentage': completion_percentage
            }

        self.visualize_and_export()

    def visualize_and_export(self):
        """
        Visualize module completion data and export to Excel.
        """
        df_completion = pd.DataFrame.from_dict(self.completion_data, orient='index')
        
        plt.figure(figsize=(12, 6))
        df_completion['completion_percentage'].plot(kind='bar', rot=45)
        plt.title('Module Completion Percentage')
        plt.xlabel('Module')
        plt.ylabel('Completion Percentage (%)')
        plt.tight_layout()
        
        module_completion_img_file = 'module_completion.png'
        plt.savefig(module_completion_img_file)
        plt.close()
        
        df_student_module_completion = pd.DataFrame.from_dict(self.student_module_completion, orient='index')
        df_student_module_completion = df_student_module_completion.applymap(lambda x: 1 if x else 0)
        
        plt.figure(figsize=(12, 8))
        df_student_module_completion.plot(kind='bar', stacked=True, rot=45)
        plt.title('Student Module Completion')
        plt.xlabel('Student')
        plt.ylabel('Module Completion Status')
        plt.legend(title='Module')
        plt.tight_layout()
        
        student_module_completion_img_file = 'student_module_completion_stacked_bar.png'
        plt.savefig(student_module_completion_img_file)
        plt.close()
        
        excel_file = "canvas_course_completion_analysis.xlsx"
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            df_completion.to_excel(writer, sheet_name='Module Completion', index=True)
            df_student_module_completion.to_excel(writer, sheet_name='Student Module Completion', index=True)
            
            worksheet = writer.sheets['Module Completion']
            worksheet.insert_image('G2', module_completion_img_file, {'x_offset': 15, 'y_offset': 10, 'x_scale': 0.75, 'y_scale': 0.75})
            
            worksheet = writer.sheets['Student Module Completion']
            worksheet.insert_image('G2', student_module_completion_img_file, {'x_offset': 15, 'y_offset': 10, 'x_scale': 0.75, 'y_scale': 0.75})

        print(f"Data exported to {excel_file}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Usage: python course_completion.py <api_url> <api_key> <course_id>")
        sys.exit(1)

    api_url = sys.argv[1]
    api_key = sys.argv[2]
    course_id = int(sys.argv[3])
    
    canvas_analyzer = CourseCompletionAnalyzer(api_url, api_key, course_id)
    canvas_analyzer.analyze_module_completion()
