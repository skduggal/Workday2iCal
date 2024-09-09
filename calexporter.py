import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime

# Specify the full path to your Excel file
excel_file_path = '/Users/sidkduggal/Downloads/View_My_Courses.xlsx'
df = pd.read_excel(excel_file_path)

# Create dictionaries to hold term 1 and term 2 events
term_1_courses = []
term_2_courses = []

# Loop through the DataFrame rows and categorize events by term
for index, row in df.iterrows():
    if pd.notna(row['Unnamed: 4']) and pd.notna(row['Unnamed: 7']):
        # Extract start date
        start_time_str = row['Unnamed: 10']  # Start date
        start_time = pd.to_datetime(start_time_str, errors='coerce')
        
        # Check the year to categorize into terms
        if pd.notna(start_time):
            if start_time.year == 2024:
                term_1_courses.append((index, row))
            elif start_time.year == 2025:
                term_2_courses.append((index, row))

# Ask the user which term they want to add to their calendar
term_choice = input("Which term would you like to add to your calendar? (Enter '1' for Term 1, '2' for Term 2): ")

# Create a calendar
cal = Calendar()

# Select the courses based on user input
if term_choice == '1':
    selected_courses = term_1_courses
    print("Adding Term 1 courses to your calendar.")
elif term_choice == '2':
    selected_courses = term_2_courses
    print("Adding Term 2 courses to your calendar.")
else:
    print("Invalid input. No courses will be added.")
    selected_courses = []

# Loop through the selected courses and create events
for index, row in selected_courses:
    event = Event()
    event.add('summary', row['Unnamed: 4'])  # Course details
    
    start_time_str = row['Unnamed: 10']  # Start date
    end_time_str = row['Unnamed: 11']  # End date
    
    start_time = pd.to_datetime(start_time_str, errors='coerce')
    end_time = pd.to_datetime(end_time_str, errors='coerce')
    
    if pd.notna(start_time) and pd.notna(end_time):
        event.add('dtstart', start_time)
        event.add('dtend', end_time)
        event.add('location', row['Unnamed: 7'])  # Using meeting patterns as location for simplicity
        cal.add_component(event)

# Write the calendar to a file
if selected_courses:
    with open('/Users/sidkduggal/Downloads/workday_schedule.ics', 'wb') as f:
        f.write(cal.to_ical())
    print("iCal file has been created successfully.")
else:
    print("No courses were added to the calendar.")
