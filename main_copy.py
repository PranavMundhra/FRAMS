import streamlit as st
import subprocess
from datetime import datetime
import os

def main():
    st.title("Attendance System")
    
    st.subheader("Welcome to Attendance System")

    if st.button("Collect Data"):
        collect_data()

    if st.button("Mark Attendance"):
        mark_attendance()

def collect_data():
    try:
        today = datetime.now()
        month_year = today.strftime("%B_%Y")
        excel_file = f"data/attendance_{month_year}.xlsx"

        # Check if the Excel file exists
        if not os.path.exists(excel_file):
            print(f"Excel file '{excel_file}' not found. Running test.py...")
            subprocess.run(["python", "excel.py"])  
        
        subprocess.run(["python", "collect_data.py"])
        st.success("Data collection completed.")
    except Exception as e:
        st.error(f"Failed to execute collect_data.py: {e}")

def mark_attendance():
    try:
        today = datetime.now()
        month_year = today.strftime("%B_%Y")
        excel_file = f"data/attendance_{month_year}.xlsx"

        # Check if the Excel file exists
        if not os.path.exists(excel_file):
            print(f"Excel file '{excel_file}' not found. Running test.py...")
            subprocess.run(["python", "excel.py"])  
            
        subprocess.run(["python", "mark_attendance.py"])
        st.success("Attendance marked successfully.")
    except Exception as e:
        st.error(f"Failed to execute mark_attendance.py: {e}")

if __name__ == "__main__":
    main()
