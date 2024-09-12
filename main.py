from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import requests
from bs4 import BeautifulSoup
from icalendar import Calendar, Event, vRecur
from datetime import datetime, timedelta

driver = webdriver.Firefox()

driver.get("https://ssp.mycampus.ca/StudentRegistrationSsb/ssb/registrationHistory/registrationHistory?mepCode=UOIT")
time.sleep(5)

username_field = driver.find_element(By.NAME, "UserName")
username = input("Enter your ID: ")
username_field.send_keys(f"oncampus.local\\{username}")

password_field = driver.find_element(By.NAME, "Password")
password = input("Enter your password: ")
password_field.send_keys(password)
password_field.send_keys(Keys.RETURN)

time.sleep(5)

driver.get('https://ssp.mycampus.ca/StudentRegistrationSsb/ssb/registrationHistory/registrationHistory?mepCode=UOIT')
time.sleep(5)

driver.execute_script("document.getElementById('lookupFilter').style.display='block';")
dropdown = Select(driver.find_element(By.ID, "lookupFilter"))
options = dropdown.options

for i, option in enumerate(options):
    print(f"{i}: {option.text}")

selected_index = int(input("Enter the index of the term you want to select: "))
dropdown.select_by_index(selected_index)

selected_option = dropdown.first_selected_option
print(f"Selected term: {selected_option.text} {selected_option.get_attribute('value')}")

html_content = driver.page_source
soup = BeautifulSoup(html_content, 'html.parser')

crn_elements = soup.find_all('td', {'data-property': 'courseReferenceNumber'})
name_elements = soup.find_all('td', {'data-property': 'courseTitle'})

crn_list = [crn.get_text(strip=True) for crn in crn_elements]
name_list = [name.get_text(strip=True) for name in name_elements]

print("Extracted CRNs:", crn_list)
print("Extracted Course Names:", name_list)

if len(crn_list) != len(name_list):
    print("Warning: The number of CRNs and course names do not match.")
    driver.quit()
    exit()

meta_tag = driver.find_element(By.XPATH, "//meta[@name='synchronizerToken']")
x_synchronizer_token = meta_tag.get_attribute('content')

crn_data = {}

for crn, name in zip(crn_list, name_list):
    print(f"Processing CRN: {crn} - {name}")

    get_url = "https://ssp.mycampus.ca/StudentRegistrationSsb/ssb/searchResults/getFacultyMeetingTimes"
    params = {
        "term": selected_option.get_attribute('value'),  
        "courseReferenceNumber": crn  
    }

    get_headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "en-CA,en-US;q=0.7,en;q=0.3",
        "Connection": "keep-alive",
        "Cookie": "; ".join([f"{cookie['name']}={cookie['value']}" for cookie in driver.get_cookies()]),  
        "Host": "ssp.mycampus.ca",
        "Referer": "https://ssp.mycampus.ca/StudentRegistrationSsb/ssb/registrationHistory/registrationHistory?mepCode=UOIT",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
        "X-Requested-With": "XMLHttpRequest",
        "X-Synchronizer-Token": x_synchronizer_token, 
        "Cache-Control": "no-cache",
        "Content-Type": "application/json;charset=UTF-8"
    }

    get_response = requests.get(get_url, headers=get_headers, params=params)
    meeting_times = get_response.json() 

    meeting_info = {
        'campus': None,
        'building': None,
        'room': None,
        'scheduleType': None,
        'startTime': None,
        'endTime': None,
        'monday': False,
        'tuesday': False,
        'wednesday': False,
        'thursday': False,
        'friday': False,
        'saturday': False,
        'sunday': False,
    }

    for meeting in meeting_times.get('fmt', []):
        meeting_time = meeting.get('meetingTime', {})

        if not meeting_info['campus']:
            meeting_info['campus'] = meeting_time.get('campusDescription')
        if not meeting_info['building']:
            meeting_info['building'] = meeting_time.get('buildingDescription')
        if not meeting_info['room']:
            meeting_info['room'] = meeting_time.get('room')
        if not meeting_info['scheduleType']:
            meeting_info['scheduleType'] = meeting_time.get('meetingScheduleType')
        if not meeting_info['startTime']:
            meeting_info['startTime'] = meeting_time.get('beginTime')
        if not meeting_info['endTime']:
            meeting_info['endTime'] = meeting_time.get('endTime')

        for day in ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']:
            if meeting_time.get(day):
                meeting_info[day] = True

    crn_data[crn] = {
        'course_name': name,
        'meeting_info': meeting_info
    }

cal = Calendar()

def day_to_weekday(day):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    return days.index(day)

def parse_time(time_str):
    """Convert time in HHMM format to a datetime.time object."""
    return datetime.strptime(time_str, "%H%M").time()

now = datetime.now()
next_year = now.year + 1
end_of_next_year = datetime(next_year, 12, 31, 23, 59, 59)

for crn, data in crn_data.items():
    
    print(f"Generating ICS event for CRN: {crn} - {data['course_name']}")
    meeting = data['meeting_info']

    if meeting['startTime'] == None:
        continue

    now = datetime.now()
    start_time = parse_time(meeting['startTime'])
    end_time = parse_time(meeting['endTime'])
    
    for day, is_present in meeting.items():
        if day in ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'] and is_present:
            weekday_index = day_to_weekday(day.capitalize())
            start_date = now + timedelta(days=(weekday_index - now.weekday() + 7) % 7)
            start_datetime = datetime.combine(start_date, start_time)
            end_datetime = datetime.combine(start_date, end_time)

            event = Event()
            event.add('summary', f"{data['course_name']} (CRN {crn})")
            event.add('dtstart', start_datetime)
            event.add('dtend', end_datetime)
            event.add('location', f"{meeting['building']}, {meeting['room']}")
            event.add('description', f"Campus: {meeting['campus']}\nCRN: {crn}\nSchedule Type: {meeting['scheduleType']}")

            recurrence = vRecur()
            recurrence['freq'] = 'weekly'
            recurrence['until'] = end_of_next_year
            recurrence['byday'] = [day[:2].upper()]  
            
            event.add('rrule', recurrence)
            
            cal.add_component(event)

with open('class_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())

driver.quit()
