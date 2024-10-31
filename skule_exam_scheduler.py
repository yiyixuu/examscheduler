import pandas as pd
import streamlit as st
from ics import Calendar, Event
from datetime import timedelta

# Function to filter courses by last name
def filter_courses_by_last_name(selected_courses, last_name):
    # Filter the selected courses based on the user's last name
    filtered_courses = []
    for index, row in selected_courses.iterrows():
        surname_range = row['Surname Range']
        if pd.isna(surname_range) or "-" not in surname_range:
            filtered_courses.append(row)
        else:
            start_surname, end_surname = surname_range.split("-")
            if start_surname.strip().upper() <= last_name.upper() <= end_surname.strip().upper() or last_name.upper().startswith(end_surname.strip().upper()[0]):
                filtered_courses.append(row)
    return pd.DataFrame(filtered_courses)

# Function to create an exam calendar
def create_exam_calendar(selected_courses, user_last_name):
    # Create a new calendar
    calendar = Calendar()

    # Iterate over each row and add an event for each exam
    for index, row in selected_courses.iterrows():
        exam_event = Event()
        exam_event.name = f"{row['Course']} Final Exam"
        exam_date_time = pd.to_datetime(f"{row['Date']} {row['Time']}")
        exam_event.begin = exam_date_time.tz_localize('America/New_York')
        exam_event.end = (exam_date_time + timedelta(hours=2, minutes=30)).tz_localize('America/New_York')  # Assuming exams are 2.5 hours long
        exam_event.location = row['Location'].replace('-', '')

        # Add a description to the event
        exam_event.description = f"{row['Course Title']}\nSurname Range: {row['Surname Range']}"

        # Add the event to the calendar
        calendar.events.add(exam_event)

    # Write to an .ics file
    filename = f"{user_last_name.lower()}_final_exams_2024_fall.ics"
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix='.ics') as tmp_file:
        tmp_file.write(calendar.serialize().encode('utf-8'))
        tmp_file_path = tmp_file.name

    with open(tmp_file_path, 'rb') as f:
        st.download_button(
            label="Download Calendar File",
            data=f,
            file_name=filename,
            mime="text/calendar"
        )

# Streamlit UI
st.title("SKULE Exam Calendar Export Tool")

# Load the exam data from the xlsx file
exam_data = pd.read_excel('data.xlsx')

# User input for last name
user_last_name = st.text_input("Enter your last name:")

# Course selection using a multi-select box
unique_courses = exam_data['Course'].unique()
selected_courses = st.multiselect("Select courses to export:", unique_courses)

if user_last_name and selected_courses:
    # Filter selected courses from the data
    selected_courses_data = exam_data[exam_data['Course'].isin([course.upper() for course in selected_courses])]
    filtered_courses = filter_courses_by_last_name(selected_courses_data, user_last_name)

    # Create and download exam calendar
    create_exam_calendar(filtered_courses, user_last_name)
else:
    st.error("Please enter your last name and select at least one course.")
