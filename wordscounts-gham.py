from docx import Document
from docx.shared import Inches
import re
import collections

def getText(filename):
    doc = Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# Store all the text in one string
myText = getText('wordscount-test.docx')

# Seperate words into a dict via pattern matching
wordPattern = '[a-zA-Z]+'
wordsDict = re.findall(wordPattern, myText)

# Count words' appearance
wordsCount = {}
for word in wordsDict:
    word = word.lower()
    if word in wordsCount.keys():
        wordsCount[word] += 1
    else:
        wordsCount[word] = 1

# Two-level Sort!
# First by key in ascending order, then by value in descending order
orderedWordCount = collections.OrderedDict(sorted(wordsCount.items(), key=lambda t: t[0]))
wordsCountList = sorted(orderedWordCount.items(), key=lambda t: t[1], reverse=True)

for kvTuple in wordsCountList:
    tupleStr = '%s : %s'%(kvTuple[0], kvTuple[1])
    print(tupleStr)

'''
for k, v in wordsCountList:
    str = '%s : %s'%(k, v)
    print(str)
'''
