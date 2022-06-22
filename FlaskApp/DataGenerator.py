import datetime
import itertools
from matplotlib.pyplot import connect
import pptx
from pptx.util import Pt
from pptx.util import Inches
from SQLTools import *
import requests
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from PPTXTools import *

host = 'localhost'
user_name = 'taran'
user_password = 'DESKTOP-RMV78RR'
db_name = 'seekapp'
myPath = "static/generated/"
logoPath = "SeekLogo.png"
max_table_size = 8
max_table_width = 5


logoLeft, logoTop, logoHeight, logoWidth = 8293608, 18288, 768096, 841248


def fastTranslateToEnglish(session, text):
    '''Detect the language of the text and translate it to english'''
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=" + text
    response = session.get(url)
    response = response.json()
    returnString = ""
    for item in response[0]:
        returnString += item[0]
    return returnString
    
def getCompanies():
    connection = create_server_connection(host, user_name, user_password, db_name)
    query = "SELECT id, name FROM company where status='active'"
    result = execute_query(connection, query)
    return result

def addDownloadData(prs, companyID, groupBy, connection = None):
    if connection is None:
        connection = create_server_connection(host, user_name, user_password, db_name)

    # Get all non inactive employees in the company
    # Query for appropriate columns and company ID
    query1 = "select id,{},status from employee where status !='inactive' and companyID={}".format(groupBy, companyID)
    # Attempt to execute_query and set result1 to the result of the query
    result1 = execute_query(connection, query1)
    # create dictionary of employee dictionaries
    employeeDict = {}
    for row in result1:
        employeeDict[row[0]] = {groupBy: row[1], "status": row[2]}
        # format: employeeDict[id] = [dept, status]
    #print("Employee dictionary created with {} employees".format(len(employeeDict)))

    # Create dictionary of unit dictionaries
    activeInactiveByGroup = {}
    for employee in employeeDict.values():
        if(employee[groupBy] == None):
            employee[groupBy] = "Other"
        currGroup = employee[groupBy].capitalize()
        if currGroup in activeInactiveByGroup:
            if employee["status"] == "active":
                activeInactiveByGroup[currGroup][0] += 1
            else:
                activeInactiveByGroup[currGroup][1] += 1
        else:
            if employee["status"] == "active":
                activeInactiveByGroup[currGroup] = [1, 0]
            else:
                activeInactiveByGroup[currGroup] = [0, 1]

    #print("Active/Inactive dictionary created with {} units".format(len(activeInactiveByGroup)))
    #print(activeInactiveByGroup)

    chunkCounter = 0
    while (len(activeInactiveByGroup) > chunkCounter*max_table_size):
        if chunkCounter == 0:
            slide = createBlankSlideWithTitle(prs, "Download Status")
        else:
            slide = createBlankSlideWithTitle(prs, "Download Status Continued")
        currChunk = dict(itertools.islice(activeInactiveByGroup.items(), chunkCounter*max_table_size, chunkCounter*max_table_size + max_table_size))
        groupDownloadTable = createTableWithDownloadHeaders(slide, len(currChunk) + 2, 5, Inches(1), Inches(1.5), Inches(8), Inches(2))
        currChunkKeys = list(currChunk.keys())
        currChunkValues = list(currChunk.values())
        for i in range (0, len(currChunk)):
            groupDownloadTable.rows[i+1].cells[0].text = currChunkKeys[i]
            groupDownloadTable.rows[i+1].cells[1].text = str(currChunkValues[i][0])
            groupDownloadTable.rows[i+1].cells[2].text = str(currChunkValues[i][1])
            groupDownloadTable.rows[i+1].cells[3].text = str(currChunkValues[i][0] + currChunkValues[i][1])
            if currChunkValues[i][0] + currChunkValues[i][1] == 0:
                groupDownloadTable.rows[i+1].cells[4].text = "0%"
            else:
                groupDownloadTable.rows[i+1].cells[4].text = str(round((currChunkValues[i][0])/(currChunkValues[i][0] + currChunkValues[i][1])*100)) + "%"
        groupDownloadTable.rows[len(currChunk)+1].cells[0].text = "Total"
        groupDownloadTable.rows[len(currChunk)+1].cells[1].text = str(sum(currChunkValues[i][0] for i in range(len(currChunk))))
        groupDownloadTable.rows[len(currChunk)+1].cells[2].text = str(sum(currChunkValues[i][1] for i in range(len(currChunk))))
        groupDownloadTable.rows[len(currChunk)+1].cells[3].text = str(sum(currChunkValues[i][0] + currChunkValues[i][1] for i in range(len(currChunk))))
        groupDownloadTable.rows[len(currChunk)+1].cells[4].text = str(round((sum(currChunkValues[i][0] for i in range(len(currChunk))))/(sum(currChunkValues[i][0] + currChunkValues[i][1] for i in range(len(currChunk))))*100)) + "%"
        
        chunkCounter += 1
        




def createValueTable(slide, answerDictByGroup, optionList):
    if len(answerDictByGroup) > 4:
        createDoubleValueTable(slide, answerDictByGroup, optionList)
        return
    left = Inches(3)
    top = Inches(2.5)
    width = Inches(6.75)
    height = Inches(2)
    table = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table.rows[0].cells[0].text = "Response"
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            table.rows[i+1].cells[j+1].text = str(answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])
    
    table.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        table.rows[len(optionList)+1].cells[j+1].text = str(sum(answerDictByGroupValueList[j].values()))

def createPercentValueTable(slide, answerDictByGroup, optionList):
    if len(answerDictByGroup) > 4:
        createDoublePercentValueTable(slide, answerDictByGroup, optionList)
        return
    left = Inches(3)
    top = Inches(2.5)
    width = Inches(6.75)
    height = Inches(2)
    table = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table.rows[0].cells[0].text = "Response"
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            try:
                table.rows[i+1].cells[j+1].text = str(round((answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])/sum(answerDictByGroupValueList[j].values())*100)) + "%"
            except ZeroDivisionError:
                table.rows[i+1].cells[j+1].text = "0%"

    table.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        try: 
            table.rows[len(optionList)+1].cells[j+1].text = str(round(sum(answerDictByGroupValueList[j].values())/sum(answerDictByGroupValueList[j].values())*100)) + "%"
        except ZeroDivisionError:
            table.rows[len(optionList)+1].cells[j+1].text = "0%"

def createDoublePercentValueTable(slide, answerDictByGroup, optionList):
    # Split dictionary into two dictionaries of equal length
    answerDictByGroup1 = {}
    answerDictByGroup2 = {}
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        if i % 2 == 0:
            answerDictByGroup1[answerDictByGroupKeyList[i]] = answerDictByGroupValueList[i]
        else:
            answerDictByGroup2[answerDictByGroupKeyList[i]] = answerDictByGroupValueList[i]


    answerDictByGroup = answerDictByGroup1

    left = Inches(2.75)
    top = Inches(1.5)
    width = Inches(6.75)
    height = Inches(2)
    table = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table.rows[0].cells[0].text = "Response"

    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            try:
                table.rows[i+1].cells[j+1].text = str(round((answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])/sum(answerDictByGroupValueList[j].values())*100)) + "%"
            except ZeroDivisionError:
                table.rows[i+1].cells[j+1].text = "0%"
            
    table.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        try:
            table.rows[len(optionList)+1].cells[j+1].text = str(round(sum(answerDictByGroupValueList[j].values())/sum(answerDictByGroupValueList[j].values())*100)) + "%"
        except ZeroDivisionError:
            table.rows[len(optionList)+1].cells[j+1].text = "0%"

    setTableFontSize(table, 16)

    answerDictByGroup = answerDictByGroup2

    left = Inches(2.75)
    top = Inches(4.25)
    width = Inches(7)
    height = Inches(2)

    table2 = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table2.rows[0].cells[0].text = "Response"
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table2.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table2.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            try:
                table2.rows[i+1].cells[j+1].text = str(round((answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])/sum(answerDictByGroupValueList[j].values())*100)) + "%"
            except ZeroDivisionError:
                table2.rows[i+1].cells[j+1].text = "0%"
            
    table2.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        try:
            table2.rows[len(optionList)+1].cells[j+1].text = str(round(sum(answerDictByGroupValueList[j].values())/sum(answerDictByGroupValueList[j].values())*100)) + "%"
        except ZeroDivisionError:
            table2.rows[len(optionList)+1].cells[j+1].text = "0%"
    
    setTableFontSize(table2, 16)


def createDoubleValueTable(slide, answerDictByGroup, optionList):
    # Split dictionary into two dictionaries of equal length
    answerDictByGroup1 = {}
    answerDictByGroup2 = {}
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        if i % 2 == 0:
            answerDictByGroup1[answerDictByGroupKeyList[i]] = answerDictByGroupValueList[i]
        else:
            answerDictByGroup2[answerDictByGroupKeyList[i]] = answerDictByGroupValueList[i]
    # Create tables

    answerDictByGroup = answerDictByGroup1

    left = Inches(2.75)
    top = Inches(1.5)
    width = Inches(6.75)
    height = Inches(2)
    table = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table.rows[0].cells[0].text = "Response"

    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            table.rows[i+1].cells[j+1].text = str(answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])
    
    table.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        table.rows[len(optionList)+1].cells[j+1].text = str(sum(answerDictByGroupValueList[j].values()))

    setTableFontSize(table, 16)

    answerDictByGroup = answerDictByGroup2

    left = Inches(2.75)
    top = Inches(4.25)
    width = Inches(7)
    height = Inches(2)

    table2 = slide.shapes.add_table(len(optionList) + 2, len(answerDictByGroup) + 1, left, top, width, height).table

    table2.rows[0].cells[0].text = "Response"
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    answerDictByGroupValueList = list(answerDictByGroup.values())
    for i in range(len(answerDictByGroup)):
        table2.rows[0].cells[i+1].text = answerDictByGroupKeyList[i]

    for i in range(len(optionList)):
        table2.rows[i+1].cells[0].text = optionList[i]
        for j in range(len(answerDictByGroup)):
            table2.rows[i+1].cells[j+1].text = str(answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]])
    
    table2.rows[len(optionList)+1].cells[0].text = "Total"
    for j in range(len(answerDictByGroup)):
        table2.rows[len(optionList)+1].cells[j+1].text = str(sum(answerDictByGroupValueList[j].values()))


    setTableFontSize(table2, 16)



def createValuePieChart(slide, answerDictByGroup, optionList):
    chart_data = ChartData()
    chartCategoriesList = []
    for i in range(len(optionList)):
        if len(optionList[i]) > 20:
            chartCategoriesList.append(optionList[i][:17] + "...")
        else:
            chartCategoriesList.append(optionList[i])
        
    chart_data.categories = chartCategoriesList
    #chart_data.add_series('Response', answerDictByGroup)
    answerDictByGroupKeyList = list(answerDictByGroup.keys())
    numAnswersDict = {}
    for i in range(len(optionList)):
        numAnswersDict[optionList[i]] = sum(answerDictByGroup[answerDictByGroupKeyList[j]][optionList[i]] for j in range(len(answerDictByGroup)))
    answerPercentageDict = {}
    for i in range(len(optionList)):
        try:
            answerPercentageDict[optionList[i]] = numAnswersDict[optionList[i]] / sum(numAnswersDict.values())
        except ZeroDivisionError:
            answerPercentageDict[optionList[i]] = 0
    chart_data.add_series('Response', answerPercentageDict.values())
    if len(optionList) > 4:
        height = Inches(3.5)
    else:
        height = Inches(3)
    pieChart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0), Inches(2.5), Inches(3), height, chart_data).chart
    pieChart.chart_style = 26
    pieChart.has_legend = True
    pieChart.legend.position = pptx.enum.chart.XL_LEGEND_POSITION.BOTTOM
    pieChart.legend.include_in_layout = False
    pieChart.plots[0].has_data_labels = True
    data_labels = pieChart.plots[0].data_labels
    data_labels.number_format = '0%'
    data_labels.position = pptx.enum.chart.XL_LABEL_POSITION.OUTSIDE_END


def addLearnData(prs, companyID, groupBy, startDate, endDate, connection = None):
    if connection is None:
        connection = create_server_connection(host, user_name, user_password, db_name)

    # Get all non inactive employees in the company
    # Query for appropriate columns and company ID
    query1 = "select id,{},status from employee where status !='inactive' and companyID={}".format(groupBy, companyID)
    query2 = '''select c.display_name,e.{},count(*) from score s
left join employee e on e.id = s.userid
left join course c on c.course_id = s.course_id
where s.userid in (select id from employee where companyId = {})
and s.date >='{}' and s.date <='{}'
group by c.display_name,e.{};'''.format(groupBy, companyID, str(startDate), str(endDate), groupBy)

    # Attempt to execute_query and set result1 to the result of the query
    result1 = execute_query(connection, query1)
    result2 = execute_query(connection, query2)
    # create dictionary of employee dictionaries
    employeeDict = {}
    for row in result1:
        employeeDict[row[0]] = {groupBy: row[1], "status": row[2]}
        # format: employeeDict[id] = [dept, status]
    # print("Employee dictionary created with {} employees".format(len(employeeDict)))


    totalEmployeesByGroup = {}
    for employee in employeeDict.values():
        try:
            currGroup = employee[groupBy].capitalize()
        except:
            currGroup = "Other"
        # if(employee[groupBy] == None):
        #     employee[groupBy] = "Other"
        # currGroup = employee[groupBy].capitalize()
        if currGroup in totalEmployeesByGroup:
            totalEmployeesByGroup[currGroup] += 1
        else:
            totalEmployeesByGroup[currGroup] = 1
    
    dataDict = {}
    courseList = []
    for row in result2:
        if row[0] not in courseList:
            courseList.append(row[0])
        if row[1] in dataDict:
            dataDict[row[1].capitalize()][row[0]] = row[2]
        else:
            dataDict[row[1].capitalize()] = {row[0]: row[2]}

    for group in totalEmployeesByGroup:
        if group not in dataDict:
            dataDict[group] = {}
    totalEmployeesLearningByGroup = {}
    totalEmployeesLearningByGroup["total"] = 0

    chunkCounter = 0
    while (len(dataDict) > chunkCounter*max_table_size):
        sliceCounter = 0
        while(len(courseList) > sliceCounter*max_table_width):
            if chunkCounter == 0 and sliceCounter == 0:
                slide = createBlankSlideWithTitle(prs, "Learn Status")
            else:
                slide = createBlankSlideWithTitle(prs, "Learn Status Continued")
            # currCourseList = courseList[chunkCounter*max_table_size:(chunkCounter+1)*max_table_size]
            currCourseList = courseList[sliceCounter*max_table_width:(sliceCounter+1)*max_table_width]
            currChunk = dict(itertools.islice(dataDict.items(), chunkCounter*max_table_size, chunkCounter*max_table_size + max_table_size))
            if len(currCourseList) < (max_table_width):
                groupLearnTable = createTableWithFinalLearnHeaders(slide, currCourseList, len(currChunk) + 2, len(currCourseList) + 3)
            else:
                groupLearnTable = createTableWithLearnHeaders(slide, currCourseList, len(currChunk) + 2, len(currCourseList) + 1)
            currChunkKeys = list(currChunk.keys())
            currChunkValues = list(currChunk.values())
            for i in range (0, len(currChunkKeys)):
                totalLearning = 0
                groupLearnTable.rows[i+1].cells[0].text = currChunkKeys[i]
                for j in range (0, len(currCourseList)):
                    if currCourseList[j] in currChunkValues[i]:
                        groupLearnTable.rows[i+1].cells[j+1].text = str(currChunkValues[i][currCourseList[j]])
                        totalLearning += currChunk[currChunkKeys[i]][currCourseList[j]]
                    else:
                        groupLearnTable.rows[i+1].cells[j+1].text = "0"
                if currChunkKeys[i] in totalEmployeesLearningByGroup:
                    totalEmployeesLearningByGroup[currChunkKeys[i]] += totalLearning
                else:
                    totalEmployeesLearningByGroup[currChunkKeys[i]] = totalLearning
                totalEmployeesLearningByGroup["total"] += totalLearning
                if len(currCourseList) < (max_table_width):

                    groupLearnTable.rows[i+1].cells[len(currCourseList)+1].text = str(totalEmployeesByGroup[currChunkKeys[i]] - totalEmployeesLearningByGroup[currChunkKeys[i]])
                    groupLearnTable.rows[i+1].cells[len(currCourseList)+2].text = str(totalEmployeesByGroup[currChunkKeys[i]])
            groupLearnTable.rows[len(currChunkKeys)+1].cells[0].text = "Total"

            companyLearning = 0
            for i in range (0, len(currCourseList)):
                totalLearning = 0
                for j in range (0, len(currChunkKeys)):
                    if currCourseList[i] in currChunkValues[j]:
                        totalLearning += currChunk[currChunkKeys[j]][currCourseList[i]]
                companyLearning += totalLearning
                groupLearnTable.rows[len(currChunkKeys)+1].cells[i+1].text = str(totalLearning)
            
            currLearning = 0
            for i in range (0, len(currChunkKeys)):
                currLearning += totalEmployeesLearningByGroup[currChunkKeys[i]]

            currEmployeeCount = 0
            for i in range (0, len(currChunkKeys)):
                currEmployeeCount += totalEmployeesByGroup[currChunkKeys[i]]


            if len(currCourseList) < (max_table_width):
                groupLearnTable.rows[len(currChunkKeys)+1].cells[len(currCourseList)+1].text = str(currEmployeeCount - currLearning)
                groupLearnTable.rows[len(currChunkKeys)+1].cells[len(currCourseList)+2].text = str(currEmployeeCount)
            sliceCounter += 1

        chunkCounter += 1

def addValueData(prs, companyID, groupBy, startDate, endDate, percentage=1, connection = None):
    if connection is None:
        connection = create_server_connection(host, user_name, user_password, db_name)

    # Get all questions asked for a company during a given time period
    query1 = "select id,question,option1,option2,option3,option4,option5,employees_type_data,actual_schedule_time from surveyquestions where compid = {} and actual_schedule_time >= '{}' and actual_schedule_time <= '{}' and scheduled_count<>0;".format(companyID, startDate, endDate)
    result1 = execute_query(connection, query1)

    query3 = "select {},count(*) from employee where companyId={} and status='active' group by {};".format(groupBy, companyID, groupBy)
    result3 = execute_query(connection, query3)
    employeeGroupList = []
    for row3 in result3:
        group = row3[0]
        if group == None:
            group = "Other"
        if group == "":
            group = "Other"
        group = group.capitalize()
        employeeGroupList.append(row3[0])

    # Remove duplicates from employeeGroupList
    employeeGroupList = list(set(employeeGroupList))

    allDataByQuestion = []

    s = requests.Session()
    for row1 in result1:
        questionID = row1[0]
        question = row1[1]
        if not (row1[7] == "English" or row1[7] == "FC" or row1[7] == "DP-English"):
            if question is not None:
                # question = translateToEnglishDos(question)
                question = fastTranslateToEnglish(s, question)
        # check if question is null
        if question is None:
            tempQuery = "select content_file from surveyquestions where id = {}".format(questionID)
            tempResult = execute_query(connection, tempQuery)
            question = tempResult[0][0]
            #question = "https://backoffice.seek-app.com//storage/" + question
            #question = processURL(question).replace('\\n', " ")
            
        query2 = "select e.{},n.answer from notifications n left join employee e on n.empId=e.id where message_id={} and answer is not null;".format(groupBy, questionID)
        result2 = execute_queryNoTime(connection, query2)

        optionList = []
        for i in range(2,7):
            if row1[i] != None and row1[i] != "":
                optionList.append(row1[i])

        answerDictByGroup = {}
        for group in employeeGroupList:
            answerDictByGroup[group] = {}
            for option in optionList:
                answerDictByGroup[group][option] = 0
        
        for row2 in result2:
            answerDictByGroup[row2[0].capitalize()][row2[1]] += 1

        allDataByQuestion.append([question, answerDictByGroup, optionList, row1[7], row1[8]])
            
    # Remove duplicates
    currLength = len(allDataByQuestion) - 1
    idx = 0
    while idx < currLength:
        try:
            isSame = allDataByQuestion[idx][4] == allDataByQuestion[idx+1][4] 
            # and allDataByQuestion[idx][3] != allDataByQuestion[idx+1][3]
            if isSame:
                firstAnswerDict = allDataByQuestion[idx][1]
                secondAnswerDict = allDataByQuestion[idx+1][1]
                firstAnswerDictKeyList = list(firstAnswerDict.keys())
                for i in range(len(firstAnswerDictKeyList)):
                    currGroup = firstAnswerDictKeyList[i]
                    k = 0
                    secondAnswerCurrGroupValueList = list(secondAnswerDict[currGroup].values())
                    for j in firstAnswerDict[currGroup].keys():
                        firstAnswerDict[currGroup][j] += secondAnswerCurrGroupValueList[k]
                        k += 1
                allDataByQuestion.pop(idx+1)
                currLength -= 1
                continue
        except:
            print("False duplicate found at index: " + str(idx))
            print("Confused question ids: " + str(allDataByQuestion[idx][0]) + " and " + str(allDataByQuestion[idx+1][0]))
            pass
        idx += 1

    # Add Slides
    for i in range(len(allDataByQuestion)):
        currQuestion = allDataByQuestion[i]
        question = currQuestion[0]
        answerDictByGroup = currQuestion[1]
        optionList = currQuestion[2]
        # Create a new slide for each question
        # print(questionID, question)
        if len(question) > 180:
            slide = createBlankSlideWithTitle(prs, question, 18)
        else:
            slide = createBlankSlideWithTitle(prs, question, 25)

        # Create a table for the question
        if percentage == 2:
            createPercentValueTable(slide, answerDictByGroup, optionList)
        else:
            createValueTable(slide, answerDictByGroup, optionList)
        # Create a pie chart for the question
        createValuePieChart(slide, answerDictByGroup, optionList)
    

def generateFullReport(companyID, filename, groupBy, startDate, endDate, options):
    startTime = time.time()
    prs = pptx.Presentation()
    connection = create_server_connection(host, user_name, user_password, db_name)
    if (options[0] == 1):
        if sum(options) == 4:
            titleSlide = createTitleSlide(prs, "Full Report for " + companyID, startDate, endDate)
        else:
            titleText = "Report for " + companyID + " with "
            if options[1] == 1:
                titleText += "download data, "
            if options[2] == 1:
                titleText += "learn data, "
            if options[3] == 1 or options[3] == 2:
                titleText += "value data, "
            titleText = titleText[:-2]
            titleSlide = createTitleSlide (prs, titleText, startDate, endDate)
    if (options[1] != 0):
        addDownloadData(prs, companyID, groupBy, connection)
    if (options[2] != 0):
        addLearnData(prs, companyID, groupBy, startDate, endDate, connection)
    if (options[3] != 0):
        addValueData(prs, companyID, groupBy, startDate, endDate, options[3], connection)

    print("PPTX file creation took {} seconds".format(time.time() - startTime))
    saved = False
    try:
        prs.save(myPath + filename + ".pptx")
        saved = True
        print("File saved as " + myPath + filename + ".pptx")
        return (myPath + filename + ".pptx")
    except:
        number = filename[-1]
        try:
            number = int(number)
        except:
            number = 1
            filename = filename + str(number)
    while(saved == False):
        try:
            prs.save(myPath + filename[:-1] + str(number) + ".pptx")
            print("File saved as " + myPath + filename[:-1] + str(number) + ".pptx")
            saved = True
        except Exception as e:
            print("Save failed. Error: {} Trying again.".format(e))
            number += 1
        if(number == 10):
            print("File could not be saved.")
            return None
    return (myPath + filename[:-1] + str(number) + ".pptx")

generateFullReport("92", "92Test2", "level", "2022-05-01", "2022-06-14", (1, 1, 1, 1))