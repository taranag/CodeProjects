def sentimentAnalysis(text):
    """Use a web service to determine whether the sentiment of text is positive""" 
    from urllib.request import urlopen
    from urllib.parse import urlencode
    from json import loads
    url = "http://text-processing.com/api/sentiment/"
    data = urlencode({"text": text}).encode("utf-8")
    response = urlopen(url, data)
    return loads(response.read().decode("utf-8"))
    
# Take user input and print sentiment analysis result
text = input("Enter text to analyze: ")
print(sentimentAnalysis(text))