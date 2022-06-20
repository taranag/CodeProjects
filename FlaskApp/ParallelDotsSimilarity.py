def getPDSimilarity(text1, text2):
    api_key = "oNUmQRAzRHS0ZRg7GK1Q24Fb272oAmHabj1DgjKhCN8"
    import requests
    files = {
        'text_1': (None, text1),
        'text_2': (None, text2),
        'api_key': (None, "oNUmQRAzRHS0ZRg7GK1Q24Fb272oAmHabj1DgjKhCN8"),
    }

    response = requests.post('https://apis.paralleldots.com/v4/similarity', files=files)
    return response.json()["similarity_score"]

#print(getPDSimilarity('Do you like me?', 'Do you hate me?'))