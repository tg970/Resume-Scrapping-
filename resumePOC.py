import docx2txt
import json

def createTxtFromDocx(file):
	str = docx2txt.process(file + ".docx") # renders file readable
	txt = open(file + ".txt", "w") # easier to scrape than a string
	txt.write(str)
	txt.close()
	return

def createDictfromTxt(file):
    dict = {}
    
    # Scraping name
    nameArr = file.readline().strip().split()
    nameDict = {"first":nameArr[0], "last":nameArr[1], "title":None}
    if len(nameArr) > 2:
        nameDict["title"] = nameArr[2]
    dict["name"] = nameDict
    
    # Scraping introduction
    intro = file.readline().strip()
    while intro == "": # Skip through blank lines until the next field is reached
        intro = file.readline().strip()
    dict["introduction"] = intro
    
    # Scraping GT experience   
    gtCompany = file.readline().strip()
    while gtCompany == "":
        gtCompany = file.readline().strip()
    
    gtTitleDate = file.readline().strip()
    while gtTitleDate == "":
        gtTitleDate = file.readline().strip()
    [gtTitle, gtDate] = gtTitleDate[:-1].split(" (")
    
    gtEngagements = []
    for gtLine in file: # Will iterate from current line in file to end of file unless broken earlier
        gtLine = gtLine.strip()
        if gtLine == "":
            continue
        elif gtLine.find("(") < 0: # The file has moved from engagements to prior experience, stop!
            break
        else:
            gtEngagements.append(createProjectFromParagraph(gtLine,True)) # Found a specific engagement, calling sub-parsing function
    
    gtExp = {"company":gtCompany, "title":gtTitle, "date":gtDate, "engagements":gtEngagements}
    dict["gtExp"] = gtExp
    
    # Scraping prior experience
    priorExp = []
    company = gtLine
    while True:
        titleDate = file.readline().strip()
        while titleDate == "":
            titleDate = file.readline().strip()
        [title, date] = titleDate[:-1].split(" (")
        
        projects = []
        for line in file:
            line = line.strip()
            if line == "":
                continue
            elif line.find("(") < 0: # The file has moved on to another former job
                priorExp.append({"company":company, "title":title, "date": date, "projects":projects})
                company = line
                break
            else:
                projects.append(createProjectFromParagraph(line))
        
        if company == "Education": # Education has been reached
            dict["priorExp"] = priorExp
            break
    
    # keep working
    
    
    return dict

def createProjectFromParagraph(str, gt=False):
    arr = str.replace(" (","|").replace(") ","|").split("|")
    if gt:
        return {"client":arr[0], "date":arr[1], "summary":arr[2]}
    else:
        return {"name":arr[0], "date":arr[1], "summary":arr[2]}
    
if __name__ == "__main__":
    fileName = "Rohan Tomer - GT resume" # do not include extension
    createTxtFromDocx(fileName)
    txtFile = open(fileName + ".txt", "r") # opening for scraping
    
    infoDict = createDictfromTxt(txtFile) # This is the key value data structure
    
    print(infoDict["name"]["last"])