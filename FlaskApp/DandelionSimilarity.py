import requests
apiKey = "2cb51e0c171645ad9eb2a8565e1a64df"

def getSimilarityScore(text1, text2):
    url = "https://api.dandelion.eu/datatxt/sim/v1/?text1=" + text1 + "&text2=" + text2 + "&token=" + apiKey
    response = requests.get(url)
    return response.json()["similarity"]

#print(getSimilarityScore("I understand that maintaining diligence in the workplace is very important to maintain the accuracy of my work.", "I understand that diligence in the workplace is very important to maintain my work accuracy."))
