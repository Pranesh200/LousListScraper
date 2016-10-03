'''
Lou's List
Enrollment/Waitlist data since 1980 for any section of any course/lecture/lab/etc.

SIS
For each class:
Average hours spent outside class
How much people felt they learned
How worthwhile people felt this course was
How well defined the course goals/requirements were
How approachable the instructor was
How effective the teacher waitlist

Data since 1980
4-Digit Semester Code
	1168 -> 1 16 8
			0 - 20th century
			1 - 21st century
			  16 - 2016
			     1 - January Semester
			     2 - Spring Semester
			     6 - Summer Semester
			     8 - Fall Semester
'''

import json, urllib.request, time, re
from bs4 import BeautifulSoup 
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series

'''
Converts from base 10 to base 26 to convert between a numerical system and a alphabet system
Give in a column index (0-)
Gives out a column letter code (A-Z,AA-ZZ,-)
'''
def convert10to26(num):
	num += 1
	alpha = ''
	while num > 0:
		d = int(num % 26)
		if d == 0:
			d = 26
		num -= d
		alpha = str(chr(65+d-1)) + alpha
		num = num / 26
	return alpha

'''
Creates a scatter plot of time (x) vs. enrollment/waitlist/interest (y)
'''
def createChart(sheet, offset, bound, title):
	chart = ScatterChart()
	chart.title = title
	chart.style = 2
	chart.x_axis.title = 'Days since enrollment start'
	chart.y_axis.title = 'Total people'
	for col in range(0, bound):
		row = 1
		while sheet[convert10to26(col*4) + str(row)].value != None:
			row += 1
		row -= 1
		xvalues = Reference(sheet, min_col=(col*4+1), min_row=2, max_row=row)
		yvalues = Reference(sheet, min_col=(col*4+1+offset), min_row=1, max_row=row)
		series = Series(yvalues, xvalues, title_from_data=True)
		chart.series.append(series)
	sheet.add_chart(chart, 'A10')

'''
Prints enrollment, waitlist, and interest data of each course in 'courses' to the Excel 'sheet'
'''
def ewCourses(sheet, semester, courses):
	initTime = -1
	for col in range(0, len(courses)):
		jsonurl = urllib.request.urlopen('http://rabi.phys.virginia.edu/mySIS/CS2/enrollmentData.php?Semester=' + semester  + '&ClassNumber=' + str(courses[col]))
		data = json.loads(jsonurl.read().decode('UTF-8'))
		if initTime == -1:
			initTime = data['enrollment'][0][0]
		sheet[convert10to26(col*4) + '1'] = str(courses[col])
		sheet[convert10to26(col*4+1) + '1'] = str(courses[col]) + ' Enrollment'
		sheet[convert10to26(col*4+2) + '1'] = str(courses[col]) + ' Waitlist'
		sheet[convert10to26(col*4+3) + '1'] = str(courses[col]) + ' Interest'
		for row in range(0, len(data['enrollment'])):
			sheet[convert10to26(col*4) + str(row+2)] = (data['enrollment'][row][0] - initTime)/1000/60/60/24
			sheet[convert10to26(col*4+1) + str(row+2)] = data['enrollment'][row][1]
			sheet[convert10to26(col*4+2) + str(row+2)] = data['waitlist'][row][1]
			sheet[convert10to26(col*4+3) + str(row+2)] = data['enrollment'][row][1] + data['waitlist'][row][1]
	createChart(sheet, 1, len(courses), 'CompSci Enrollment')
	createChart(sheet, 2, len(courses), 'CompSci Waitlist')
	createChart(sheet, 3, len(courses), 'CompSci Interest')

'''
Returns section numbers of only lecture sections of a 'course'
'''
def getLecturesFromCourse(semester, group, course):
	html = BeautifulSoup(urllib.request.urlopen('http://rabi.phys.virginia.edu/mySIS/CS2/page.php?Semester=' + semester + '&Type=Group&Group=' + group).read(), 'html.parser')
	courseAcronym = re.search("([A-Z]{2,})", html.find(text=course).parent['onclick']).group(0);
	courseNumber = re.search("(\d+)", html.find(text=course).parent['onclick']).group(0);
	courseID = courseAcronym + courseNumber
	courseNumbers = []
	print(courseAcronym + ' ' + courseNumber)
	for section in html.find_all(class_=courseID):
		if(section.find(text='Lecture') == 'Lecture'):
			courseNumbers.append(section.find(class_='Link').text)
	return courseNumbers

'''
Get course numbers of all courses of a group
'''
def getAllCoursesFromGroup(semester, group, filter_empty=True):
	courses = []
	html = BeautifulSoup(urllib.request.urlopen('http://rabi.phys.virginia.edu/mySIS/CS2/page.php?Semester=' + semester + '&Type=Group&Group=' + group + '&Print=').read(), 'html.parser')
	for link in html.find_all(class_='Link'):
		if 'EnrollmentGraph' in link['onclick'] and not ("0 /" in link.text or "1 /" in link.text):
			courses.append(re.search("\d{5}", link['onclick']).group(0))
	return courses

'''
Calculate the areas under the time (x) vs. enrollment/waitlist/interest (y) curves for a course (defined by colset)
The resulting number quantifies overall enrollment, waitlist, and interest in the course
'''
def getColumnStatistics(sheet, colset):
	# print('getWaitlistTotal')
	row = 2
	stats = [0,0,0]
	# print(convert10to26(colset*4) + str(row))
	while sheet[convert10to26(colset*4) + str(row)].value != None and sheet[convert10to26(colset*4) + str(row+1)].value != None:
		# print(convert10to26(colset*4+2) + str(row) + ' : ' + str(sheet[convert10to26(colset*4+2) + str(row)].value) + " | " + \
		# 	str(abs(sheet[convert10to26(colset*4+2) + str(row+1)].value - sheet[convert10to26(colset*4+2) + str(row)].value)) + " | " + \
		# 	str((sheet[convert10to26(colset*4) + str(row+1)].value - sheet[convert10to26(colset*4) + str(row)].value)))
		stats[0] += abs(sheet[convert10to26(colset*4+1) + str(row+1)].value - sheet[convert10to26(colset*4+1) + str(row)].value) / \
		(sheet[convert10to26(colset*4) + str(row+1)].value - sheet[convert10to26(colset*4) + str(row)].value)
		stats[1] += abs(sheet[convert10to26(colset*4+2) + str(row+1)].value - sheet[convert10to26(colset*4+2) + str(row)].value) / \
		(sheet[convert10to26(colset*4) + str(row+1)].value - sheet[convert10to26(colset*4) + str(row)].value)
		stats[2] += abs(sheet[convert10to26(colset*4+3) + str(row+1)].value - sheet[convert10to26(colset*4+3) + str(row)].value) / \
		(sheet[convert10to26(colset*4) + str(row+1)].value - sheet[convert10to26(colset*4) + str(row)].value)
		row += 1
	return stats

'''
Calculate areas under the curve for all courses in a sheet
'''
def getSheetStatistics(sheet):
	colset = 0
	stats = [0,0,0]
	while sheet[convert10to26(colset*4) + '1'].value != None:
		colStats = getColumnStatistics(sheet, colset)
		stats[0] += colStats[0]
		stats[1] += colStats[1]
		stats[2] += colStats[2]
		colset += 1
	return stats

'''
Calculate areas under the curve for all courses for all sheets
'''
def getStatisticsForGroups(book, groups):
	cnt = 1
	statSheet = book.active
	statSheet.title = 'Stats'
	for group in groups:
		print(group)
		sheet = book.create_sheet()
		sheet.title = group
		courseNumbers = getAllCoursesFromGroup('1168', group)
		ewCourses(sheet, '1168', courseNumbers)
		statSheet['A' + str(cnt)] = group
		stats = getSheetStatistics(sheet)
		statSheet['B' + str(cnt)] = stats[0]
		statSheet['C' + str(cnt)] = stats[1]
		statSheet['D' + str(cnt)] = stats[2]
		cnt += 1

def main():
	book = Workbook()
	engineeringGroups = ['APMA','BME','CHE','CEE','CompSci','ENGR','ECE','MSE','MAE','STS','SYS']
	collegeGroups = ['Anthropology','Art','Astronomy','Biology','Chemistry','Classics','Drama','EALC','Economics',\
		'English','EnviSci','French', 'German','History','Mathematics','MDST','MESA','Music','Philosophy','Physics',\
		'Politics','PHS','Psychology','ReliStu','Slavic','Sociology','SPAN','Statistics']
	getStatisticsForGroups(book, collegeGroups)
	book.save('out.xlsx')
main()
