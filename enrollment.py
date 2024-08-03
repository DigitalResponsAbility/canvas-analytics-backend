import requests
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

class CourseEnrollmentAnalyzer:
    def __init__(self, api_url, api_key, course_id):
        self.api_url = api_url
        self.api_key = api_key
        self.course_id = course_id
        self.df = None  # Will hold the DataFrame for enrollments data
        
    def fetch_enrollments_data(self):
        """
        Fetch enrollments data from Canvas API and convert to a pandas DataFrame.
        """
        url = f"{self.api_url}/courses/{self.course_id}/enrollments"
        headers = {"Authorization": f"Bearer {self.api_key}"}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            self.df = pd.json_normalize(data)
        else:
            print(f"Failed to fetch data: {response.status_code}")
            exit()
    
    def analyze_activity_time(self):
        """
        Analyze activity time per student and visualize it.
        """
        # Convert total_activity_time to numeric
        self.df['total_activity_time'] = pd.to_numeric(self.df['total_activity_time'], errors='coerce').fillna(0)

        # Calculate total activity time per student
        total_activity_time_per_student = self.df.groupby('user.name')['total_activity_time'].sum()

        # Print total activity time per student
        print("Total Activity Time per Student:")
        print(total_activity_time_per_student)
        print()

        # Plotting total activity time per student
        plt.figure(figsize=(10, 6))
        total_activity_time_per_student.plot(kind='bar', rot=45)
        plt.title('Total Activity Time per Student')
        plt.xlabel('Student')
        plt.ylabel('Total Activity Time (seconds)')
        plt.tight_layout()

        # Save the plot as an image
        activity_time_img_file = "activity_time.png"
        plt.savefig(activity_time_img_file)
        plt.close()
    
    def analyze_last_activity(self):
        """
        Analyze last activity per student and visualize it.
        """
        # Convert last_activity_at to datetime and localize to UTC (if not already timezone-aware)
        self.df['last_activity_at'] = pd.to_datetime(self.df['last_activity_at']).dt.tz_localize(None)

        # Find the latest activity date per student
        latest_activity_per_student = self.df.groupby('user.name')['last_activity_at'].max()

        # Print latest activity per student
        print("Latest Activity per Student:")
        print(latest_activity_per_student)
        print()

        # Plotting last activity per student
        plt.figure(figsize=(10, 6))
        plt.plot(latest_activity_per_student.index, latest_activity_per_student.values, marker='o', linestyle='-', color='b')
        plt.title('Last Activity per Student')
        plt.xlabel('Student')
        plt.ylabel('Last Activity Date')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save the plot as an image
        last_activity_img_file = "last_activity.png"
        plt.savefig(last_activity_img_file)
        plt.close()
    
    def analyze_participation_frequency(self):
        """
        Analyze course participation frequency and visualize it.
        """
        # Convert 'created_at' to datetime
        self.df['created_at'] = pd.to_datetime(self.df['created_at'])

        # Set 'created_at' as datetime index
        self.df.set_index('created_at', inplace=True)

        # Calculate participation frequency
        participation_frequency = self.df.resample('W')['user_id'].count()

        # Format the index to show dates in 'MM/DD/YYYY' format
        participation_frequency.index = participation_frequency.index.strftime('%m/%d/%Y')

        # Plotting participation frequency
        plt.figure(figsize=(12, 6))
        participation_frequency.plot(kind='bar', rot=45)
        plt.title('Course Participation Frequency')
        plt.xlabel('Date')
        plt.ylabel('Frequency')
        plt.tight_layout()

        # Save the plot as an image
        participation_frequency_img_file = 'course_participation_frequency.png'
        plt.savefig(participation_frequency_img_file)
        plt.close()
    
    def export_to_excel(self):
        """
        Export data and visualizations to Excel.
        """
        # Create a Pandas Excel writer using XlsxWriter as the engine
        excel_file = "canvas_enrollments_analysis.xlsx"
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            
            # Write enrollments DataFrame to 'Enrollments' sheet
            self.df.to_excel(writer, sheet_name='Enrollments', index=False)
            
            # Write total activity time per student to 'Total Activity Time' sheet
            total_activity_time_per_student = self.df.groupby('user.name')['total_activity_time'].sum()
            total_activity_time_per_student.to_excel(writer, sheet_name='Total Activity Time', header=['Total Activity Time (seconds)'])
            
            # Write latest activity per student to 'Latest Activity' sheet
            self.df['last_activity_at'] = pd.to_datetime(self.df['last_activity_at']).dt.tz_localize(None)
            latest_activity_per_student = self.df.groupby('user.name')['last_activity_at'].max()
            latest_activity_per_student.to_excel(writer, sheet_name='Latest Activity')
            
            # Write participation frequency to 'Course Participation Frequency' sheet
            participation_frequency = self.df.resample('W')['user_id'].count()
            participation_frequency.index = participation_frequency.index.strftime('%m/%d/%Y')
            participation_frequency.to_excel(writer, sheet_name='Course Participation Frequency', header=['Frequency'])
            
            # Insert images into Excel
            worksheet = writer.sheets['Total Activity Time']
            worksheet.insert_image('D2', 'activity_time.png', {'x_offset': 15, 'y_offset': 10, 'x_scale': 0.5, 'y_scale': 0.5})
            
            worksheet = writer.sheets['Latest Activity']
            worksheet.insert_image('D2', 'last_activity.png', {'x_offset': 15, 'y_offset': 10, 'x_scale': 0.5, 'y_scale': 0.5})
            
            worksheet = writer.sheets['Course Participation Frequency']
            worksheet.insert_image('D2', 'course_participation_frequency.png', {'x_offset': 15, 'y_offset': 10, 'x_scale': 0.5, 'y_scale': 0.5})

        print(f"Data exported to {excel_file}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Usage: python enrollment.py <api_url> <api_key> <course_id>")
        sys.exit(1)

    api_url = sys.argv[1]
    api_key = sys.argv[2]
    course_id = int(sys.argv[3])
    
    canvas_analyzer = CourseEnrollmentAnalyzer(api_url, api_key, course_id)
    canvas_analyzer.fetch_enrollments_data()
    canvas_analyzer.analyze_activity_time()
    canvas_analyzer.analyze_last_activity()
    canvas_analyzer.analyze_participation_frequency()
    canvas_analyzer.export_to_excel()
