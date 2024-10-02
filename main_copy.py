import streamlit as st
import subprocess

def main():
    st.title("Attendance System")
    
    st.subheader("Welcome to Attendance System")

    if st.button("Collect Data"):
        collect_data()

    if st.button("Mark Attendance"):
        mark_attendance()

def collect_data():
    try:
        subprocess.run(["python", "data.py"])
        st.success("Data collection completed.")
    except Exception as e:
        st.error(f"Failed to execute collect_data.py: {e}")

def mark_attendance():
    try:
        subprocess.run(["python", "mark_attendance.py"])
        st.success("Attendance marked successfully.")
    except Exception as e:
        st.error(f"Failed to execute mark_attendance.py: {e}")

if __name__ == "__main__":
    main()
