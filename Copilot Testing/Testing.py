def replaceFirstLetters(sentence):
    # Split the sentence into list words
    sentenceList = sentence.split()
    # Create string where the new string will be stored
    newSentence = ""
    # Loop through the list of words (except for the last one)
    for i in range(len(sentenceList)-1):
        # Replace the first letter of the word with the first letter of the last word and add a space
        # sentenceList[i+1] is the next word and sentenceList[i+1][0] is the first letter of the next word
        # sentenceList[i] is the current word and sentenceList[i][1:] is the rest of the word (without the first letter)
        newSentence += sentenceList[i+1][0] + sentenceList[i][1:] + " "
    # Add the last word to the new string with the first letter of the first word
    newSentence += sentenceList[0][0] + sentenceList[-1][1:]
    # Return the new string
    return newSentence


# print(replaceFirstLetters("The quick brown fox jumped over the lazy dog"))

nums = [1, 2, 3, 4]
nums.insert(4, "end")   # index 4 doesn't exist
print(nums)

nums.insert(len(nums), "end")
print(nums)
