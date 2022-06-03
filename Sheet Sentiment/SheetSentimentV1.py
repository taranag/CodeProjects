# This code takes in a sheet and returns a sentiment score for each sentence in a given column and row range
# Author: Taran Agnihotri
# Version: 1.0
# Last Updated: 3 June 2022

# Imports
import tkinter


# Sentiment Analysis Function
# text: the text to analyze
# returns: the sentiment of the text
def sentimentAnalysis(text):
    """Use a web service to determine whether the sentiment of text is positive""" 
    from urllib.request import urlopen
    from urllib.parse import urlencode
    from json import loads
    url = "http://text-processing.com/api/sentiment/"
    data = urlencode({"text": text}).encode("utf-8")
    response = urlopen(url, data)
    responseText = loads(response.read().decode("utf-8"))
    sentiment = responseText["label"]
    return sentiment


def SheetColorChanger(fileName):
    '''changes color of a cell of a sheet using pyxl'''
    import openpyxl as pyxl
    # open the file
    wb = pyxl.load_workbook(fileName)
    # get the sheet
    sheet = wb['Sheet1']
    # get the cell
    cell = sheet['A1']
    # change the color of the cell
    cell.fill = pyxl.styles.PatternFill(patternType='solid', fgColor='00FF00')
    # save the file
    wb.save(fileName)

#SheetColorChanger("testSheet.xlsx")

def pyxlSheetReader(fileName, row, col):
    '''reads a cell of a sheet using pyxl'''
    import openpyxl as pyxl
    # open the file
    wb = pyxl.load_workbook(fileName)
    # get the sheet
    sheet = wb['Sheet1']
    # get the cell
    cell = sheet[row][col]
    # return the cell
    return cell.value

def sheetSentimentAnalyzer():
    '''take user input and save it to start''' 
    Column = input("Enter column to analyze: ")
    startRow = input("Enter start row: ")
    endRow = input("Enter end row: ")


def fileSelector(): 
    '''Use tkinter to create gui file selector''' 
    import tkinter
    from tkinter import filedialog
    root = tkinter.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path


SheetColorChanger(fileSelector())
