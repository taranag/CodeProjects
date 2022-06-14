
import requests


    
def getNumberUsingAPI(a, b):
    url = 'http://127.0.0.1:5000/get_number'
    type = 'GET'
    data = {'val1': a, 'val2': b}
    response = requests.get(url, data)
    return response.text

#print (getNumberUsingAPI(5, 4))

def getPPTXurl(companyID, filename, groupBy):
    url = 'http://127.0.0.1:5000/get_PPTX'
    type = 'GET'
    data = {'companyID': companyID, 'filename': filename, 'groupBy': groupBy}
    response = requests.get(url, data)
    return response.text

print(getPPTXurl(92, "FirstTry2", "level"))

