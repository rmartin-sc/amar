# AMAR (Analyze My Attendance Records)

A CLI tool for analyzing your MS Teams attendance records.  (Useful for teachers using MS Teams for online classes.)

## Installation

`npm install -g amar`

## Usage

Just run the command below with the path to a folder containing MS Teams attendance record files.  

### Assumptions made by amar:
- Attendance report file names start with the name `meetingAttendance`
- All attendance reports in the folder pertain to *one* class
- Any presence at all (e.g. joining for 1min out of a 2hr meeting) counts as presence

`amar <path>`

The output is a list of session dates found in the report files followed by a list of names sorted by attendance rate in ascending order.  Beside each name is the ratio of sessions attended by that person followed by their attendance records (`P`=present, `A`=absent).  Each column in the attendance record has a heading corresponding to the date of the session that column represents.  (E.g. The heading `J` `1` `0` means 'January 10')