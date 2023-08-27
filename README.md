# Microsoft-Graph-Tools

## Microsoft Teams Attendance Report
Automate teams attendance report using microsoft graph api and export to excel files for easily analyze report.
Please be note, you will need meeting creator account details (email & password) to run properly.
The script only fetch meeting report if participant above 50 person. You can change it by editing 50 by report_detail variable

  report_detail = report_detail[report_detail['totalParticipantCount'] > 50]



Prequisites:
- Meeting creator account
- Azure Application (To get Client ID and Secret Key)
- Users, Events, and Online Meetings API priviledges

Features:
- Get meeting report from URL
- Get latest meeting report

