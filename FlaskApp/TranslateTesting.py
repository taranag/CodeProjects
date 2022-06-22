import json
import time
import requests

def translateToEnglish(text):
    '''Detect the language of the text and translate it to english'''
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=" + text
    response = requests.get(url)
    response = response.text
    print(response)
    response = response.split("\"")
    print(response)
    for item in response:
        print(item)
    response = response[1]
    print(response)
    return response


def translateToEnglishDos(text):
    '''Detect the language of the text and translate it to english'''
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=" + text
    response = requests.get(url)
    response = response.json()
    returnString = ""
    for item in response[0]:
        returnString += item[0]
    return returnString

translateToEnglishDos('''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸಹ ಅಭಿವೃದ್ಧಿಪಡಿಸುತ್ತಾರೆ. ಇತರರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಸಂಸ್ಕೃತಿಯ ನಂತರ ನಿಮ್ಮ ಕೆಲಸದ ಸಂಬಂಧಗಳು ಸುಧಾರಿಸಿದೆಯೇ?''')


def fastTranslateToEnglish(session, text):
    '''Detect the language of the text and translate it to english'''
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=" + text
    response = session.get(url)
    response = response.json()
    returnString = ""
    for item in response[0]:
        returnString += item[0]
    return returnString

listOfTexts = [
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
    '''ಅವರ ಪ್ರಯತ್ನಗಳನ್ನು ಮೆಚ್ಚುವ ಉದ್ಯೋಗಿಗಳು ಕೆಲಸದ ಸ್ಥಳದಲ್ಲಿ ಹೆಚ್ಚಿನ ಚಾಲನೆ ಮತ್ತು ನಿರ್ಣಯವನ್ನು ಹೊಂದಿರುತ್ತಾರೆ. ಜನರು ಉತ್ತಮ ಕೆಲಸದ ಸಂಬಂಧಗಳನ್ನು ಸುಧಾರಿಸಿದೆಯೇ?''',
]   

def fastTranslateTest():
    startTime = time.time()
    s = requests.Session()
    for i in listOfTexts:
        # print(fastTranslateToEnglish(s, i))
        print(translateToEnglishDos(i))
    print("Time taken:", time.time() - startTime)

fastTranslateTest()