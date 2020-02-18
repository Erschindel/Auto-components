import docx
import re

#create document
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

#format from list of tuples to list. the ORs in the regex result in tuples when put through re.findall
def formatList(raw, clean, size) :
    for group in range(len(raw)) :
        for dex in range(size) :
            if raw[group][dex] != "" :
                clean.append(raw[group][dex])
    return

claims = getText("test2.docx")
components = []
firstComponents = []
lowerComponents = []
firstWithSubs = []
paragraphs = []

#identify and separate components
r1 = "(?<=\na )(.*)(?=;)|(?<=\nan )(.*)(?=;)|(?<=\n)(at.*)(?=;)"
firstCompRaw = re.findall(r1, claims)
formatList(firstCompRaw, firstComponents, 3)
for comp in firstComponents:
    components.append({comp: []})

print(components)

r2 = "(?<=\nthe )(.*)(?= further comprises)|(?<=\nthe )(.*)(?= comprises)"
firstWithSubsRaw = re.findall(r2, claims)
formatList(firstWithSubsRaw, firstWithSubs, 2)

#r3 identifies all subcomponent names, including subs of subs
r3 = "(?<=comprises a )(.*?)(?=, )|(?<=comprises a )(.*?)(?= and)|(?<=comprises a )(.*?)(?=;)|(?<=comprises )(at.*?)(?=,)|(?<=comprises )(at.*?)(?= and)|(?<=comprises )(at.*?)(?=;)|(?<= an )(.*?)(?=, )|(?<=comprises an )(.*?)(?= and)|(?<=comprises an )(.*?)(?=;)|(?<= and an )(.*?)(?=;)|(?<=, )(at.*?)(?=,)|(?<=, and )(at.*?)(?=;)|(?<=, a )(.*?)(?=,)|(?<= and a )(.*?)(?=;)|(?<= and )(at.*?)(?=;)"
subCompsRaw = re.findall(r3, claims)
formatList(subCompsRaw, lowerComponents, 15)

#create list of paragraphs
r4 = "(?<=\n)(.*?)(?=\n)"
paragraphs.append(re.findall(r4, claims))

#fill components dictionary with appropriate subcomponents
r5 = "(?<=the )(.*)(?= further comprises)|(?<=the )(.*)(?= comprises)"
for line in range(len(paragraphs[0])) :
    key = re.findall(r5, paragraphs[0][line])
    cleanKey = []
    formatList(key, cleanKey, 2)
    preFormat = re.findall(r3, paragraphs[0][line])
    if key :
        chunk = []
        cleanSubs = []
        formatList(preFormat, cleanSubs, 15)
        for each in range(len(preFormat)) :
            for comp in range(len(preFormat[each])) :
                if preFormat[each][comp] :
                    chunk.append(preFormat[each][comp])
        for i in range(len(components)) :
            jkl = list(components[i].keys())
            if jkl[0] == cleanKey[0] :
                for subcomp in range(len(chunk)) :
                    components[i][cleanKey[0]].append({chunk[subcomp]:[]})
#        for i in range(len(components)) :
#            jkl = list(components[i].keys())
#            if(components[i][jkl[0]]) :
#                for j in range(len(components[i][jkl[0]])) :
#                    thirdLevelKeys = list(components[i][jkl[0]][j].keys())
#                    print(thirdLevelKeys[0])
print(components)
