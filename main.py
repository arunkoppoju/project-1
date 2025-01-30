import os
import openpyxl
from datetime import datetime
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout


# Initialize the Excel file if it doesn't exist
def initialize_excel():
    if not os.path.exists("attendance.xlsx"):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Attendance"
        sheet.append(["Name", "Date", "Time"])  # Add headers
        workbook.save("attendance.xlsx")
        print("Initialized attendance.xlsx with headers.")


class AttendanceWindow(BoxLayout):
    def start_attendance(self):
        # Retrieve the name from the TextInput field
        name = self.ids.name_input.text.strip()

        if not name:
            self.ids.status_label.text = "Error: Please enter your name."
            return

        try:
            # Open the Excel workbook and append the new record
            workbook = openpyxl.load_workbook("attendance.xlsx")
            sheet = workbook.active

            # Get the current date and time
            now = datetime.now()
            date = now.strftime("%Y-%m-%d")
            time = now.strftime("%H:%M:%S")

            # Append the new attendance record
            sheet.append([name, date, time])
            workbook.save("attendance.xlsx")

            # Update the status label
            self.ids.status_label.text = f"Recorded: {name} at {date} {time}."

            # Clear the input field
            self.ids.name_input.text = ""

        except Exception as e:
            self.ids.status_label.text = f"Error saving to Excel: {e}"

    def stop_attendance(self):
        App.get_running_app().stop()


class AttendanceApp(App):
    def build(self):
        initialize_excel()  # Ensure the Excel file exists
        return AttendanceWindow()


if __name__ == "__main__":
    AttendanceApp().run()
