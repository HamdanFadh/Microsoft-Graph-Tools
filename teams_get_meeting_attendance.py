pip install requests msal

import msal
import pandas as pd
import requests
import json

client_id = 'CLIENT_ID'
client_secret = 'SECRET_KEY'
TENANT_ID = 'TenantID'
authority = f'https://login.microsoftonline.com/{TENANT_ID}'
base_url = 'https://graph.microsoft.com/v1.0/'
username = 'MEETING CREATOR EMAIL'
password = 'MEETING CREATOR PASSWORD'

endpoint = base_url + 'me'
SCOPES = ['User.Read.All','OnlineMeetings.Read','OnlineMeetingArtifact.Read.All'] #Set in azure application preview

#Get Token
client_instance = msal.ConfidentialClientApplication (client_id = client_id, client_credential = client_secret, authority = authority)
token_result = client_instance.acquire_token_by_username_password(username = username, password = password, scopes = SCOPES)
token = token_result['access_token']
url = 'https://graph.microsoft.com/v1.0/me/events'
headers = {
 'Authorization': 'Bearer {}'.format(token)
}

# Custom MeetingId
attendance_df = []
meeting_url = input('Input Meeting URL')

def get_attendance_report(meeting_url):
  #get_meeting_id = requests.get(url= 'https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=joinMeetingIdSettings/joinMeetingId+eq+' + '\''+ joinMeetingId + '\'', headers=headers)
  get_meeting_id = requests.get(url= 'https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=joinWebURL+eq+' + '\''+ meeting_url + '\'', headers=headers)
  meetingid_result = get_meeting_id.json()
  meeting_detail = pd.json_normalize(meetingid_result, record_path = ['value'])
  meeting_id = meeting_detail['id'].iloc[0]
  meeting_subject = meeting_detail['subject'].iloc[0]
  meeting_time = meeting_detail['startDateTime'].iloc[0].replace("T"," ").replace("Z","UTC")
  get_report_id = requests.get(url = f'https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}/attendanceReports', headers=headers)
  extract_report = get_report_id.json()
  report_detail = pd.json_normalize(extract_report, record_path = ['value'])
  report_detail = report_detail[report_detail['totalParticipantCount'] > 50]
  report_id = report_detail['id'].iloc[0]
  get_attendance_report = requests.get(url = f'https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}/attendanceReports/{report_id}/attendanceRecords', headers=headers)
  extract_attendance = get_attendance_report.json()
  attendance_report = pd.json_normalize(extract_attendance, record_path = ['value'])
  attendance_df = pd.DataFrame(attendance_report)
  attendance_df.to_excel(f'{meeting_time} {meeting_subject}.xlsx')
  print(f'Saved as \'{meeting_time} {meeting_subject}.xlsx\'')

def latest_meeting_report():
  get_events = requests.get(url=url, headers=headers)
  event_results = get_events.json()
  event_details = pd.json_normalize(event_results, record_path =['value'])
  meeting_url = event_details['onlineMeeting.joinUrl'].iloc[0]
  print(meeting_url)
  get_attendance_report(meeting_url)

if meeting_url == '':
  print('get latest meeting report')
  latest_meeting_report()
else:
  print ('Get report from meeting url')
  get_attendance_report(meeting_url)


