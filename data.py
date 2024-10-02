import streamlit as st
import cv2
import os
from openpyxl import load_workbook, Workbook
from datetime import datetime

def collect_data():
    st.title("Facial Recognition Attendance System")

    student_name = st.text_input("Enter student name:")
    
    if st.button("Start Collecting Data"):
        if not student_name:
            st.error("Please enter student name.")
        else:
            collect_face_data(student_name)

def collect_face_data(student_name):
    # Implementation of webcam face capture remains the same as before
    student_dir = os.path.join("data", student_name)
    if not os.path.exists(student_dir):
        os.makedirs(student_dir)

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        st.error("Unable to access webcam!")
        return

    count = 0
    while count < 150:
        ret, frame = cap.read()
        if not ret:
            st.error("Failed to capture image!")
            break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml').detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            cv2.imwrite(f"{student_dir}/face_{count}.jpg", frame[y:y+h, x:x+w])
            count += 1

        if count == 50:
            st.info("Rotate your face slightly to the right.")
        if count == 100:
            st.info("Rotate your face slightly to the left.")

    cap.release()
    cv2.destroyAllWindows()

    if count < 150:
        st.warning(f"Only {count} images were captured for {student_name}. Please retry.")
    else:
        update_excel(student_name)
        st.success(f"Data for {student_name} added successfully!")

def update_excel(student_name):
    today = datetime.now()
    month_year = today.strftime("%B_%Y")
    excel_file = f"data/attendance_{month_year}.xlsx"
    
    if not os.path.exists(excel_file):
        wb = Workbook()
        wb.save(excel_file)

    wb = load_workbook(excel_file)
    if month_year not in wb.sheetnames:
        wb.create_sheet(month_year)

    ws = wb[month_year]
    
    existing_students = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
    if student_name not in existing_students:
        row = 1
        while ws.cell(row=row, column=1).value is not None:
            row += 1
        ws.cell(row=row, column=1, value=student_name)

    wb.save(excel_file)

if __name__ == "__main__":
    collect_data()
