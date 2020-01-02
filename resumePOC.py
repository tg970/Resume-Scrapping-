import docx2txt
import json
import os

def createTxtFromDocx(file):
    str = docx2txt.process(file + ".docx") # renders file readable
    txt = open(file + ".txt", "w", encoding = 'utf-8') # easier to scrape than a string
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
    dict["#INTRODUCTION"] = intro
    
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
                priorExp.append({"company":company, "title":title, "date": date, 
                                 "projects":projects})
                company = line
                break
            else:
                projects.append(createProjectFromParagraph(line))
        
        if company == "Education": # Education has been reached
            dict["priorExp"] = priorExp
            break
    
    eduArr = []
    eduArr = company
    eduArr = file.readline().strip()
    while True:
        if eduArr == 'Years of Federal Experience':
            break
        elif eduArr == '':
            eduArr = file.readline().strip().split(',')
            continue
        else:
            if len(eduArr)<=1:
                print("Here")
                break
            else:
                edu =  {"degree":eduArr[0], "field":eduArr[1], "college":eduArr[2],'year':eduArr[3],
                        'minor':eduArr[4]}
                dict["peducation"] = edu
                break
    
    # List of the remaining feild headers in the source template
    fieldList = ['Years of Federal Experience','Training and Certifications','Language Skills',
                 'International Experience','Computer Skills','Software','Hardware','Affiliations',
                 'Military Service','Awards','Research','Teaching','Publications',
                 'Security Clearance','']
    # Keys for the corresponding headers
    keyList = ['#YEARSEXPERIENCE','#CERTIFICATIONS','#LANGUAGES','#INTERNATIONAL','#COMPUTERSKILLS',
               '#SOFTWARE','#HARDWARE','#AFFILIATIONS','#MILSERV','#AWARDS','#RESEARCH','#TEACHING',
               '#PUBLICATIONS','#SECURITYCLEARANCE']
     
    # Scrapes everything after education
    # set operating line
    line = eduArr
    #Loop through the fields in the list
    for x in range(0,len(fieldList)-1):
        #Get the key from the key list and read the next line
        key = keyList[x]
        line = file.readline().strip()
        while True:
            # read the next line
            line = file.readline().strip()
            # stop if you've hit the next field name
            if line == fieldList[x+1]:
                    break
            # otherwise if the line is blank, or the feild name were looking for, read the next line
            elif line == '' or line == fieldList[x]:
                line = file.readline().strip()
                continue
            #Storing the correct line with the appropriate key
            else:         
                dict[key] = line.strip()
                break 
           
    return dict

def createProjectFromParagraph(str, gt=False):
    arr=[]
    arr.append(str.split('–')[0].rsplit('(',1)[0].strip())
    arr.append(str.split('–')[0].rsplit('(',1)[1].strip() + ' – ' + str.split('–')[1].split(')')[0].strip())
    arr.append(str.split('–')[1].split(')',1)[1].strip())
    #arr = str.replace(" (","|").replace(")","|").split("|")
    if gt:
        return {"#AGENCY":arr[0], "#PROJECTDATES":'(' + arr[1] + ')', "#EXPERIENCE":arr[2]}
    else:
        return {"name":arr[0], "date":arr[1], "summary":arr[2]}
    
if __name__ == "__main__":
    #fileName = 'sample_resumes/' +"Rohan Tomer - GT resume" # do not include extension
    fileName = 'sample_resumes/' + "BSullivan_Resume"
    createTxtFromDocx(fileName)
    txtFile = open(fileName + ".txt", "r",encoding='utf-8') # opening for scraping
    infoDict = createDictfromTxt(txtFile) # This is the key value data structure
    
    #print(infoDict["gtExp"]["engagements"][0])
    #print(infoDict)
    exec(open("DocxTesting.py").read())

  
# ------------------------------------------ Graveyard ---------------------------------------------
    '''
    fYrs =  eduArr
    fYrs = file.readline().strip()
    while True:
        fYrs = file.readline().strip()
        if fYrs == 'Training and Certifications':
            break
        elif fYrs == '' or fYrs =='Years of Federal Experience':
            fYrs = file.readline().strip()
            continue
        else:            
            dict["#YEARSEXPERIENCE"] = fYrs
            break
        
    train =  fYrs
    train = file.readline().strip()
    while True:
        train = file.readline().strip()
        if train == 'Language Skills':
            break
        elif train == '' or train =='Training and Certifications':
            train = file.readline().strip()
            continue
        else:            
            dict["#TRAININGS"] = train
            break
    
    lang =  train
    lang = file.readline().strip()
    while True:
        lang = file.readline().strip()
        if lang == 'International Experience':
            break
        elif lang == '' or lang =='Language Skills':
            lang = file.readline().strip()
            continue
        else:            
            dict["#LANGUAGES"] = lang
            break    
        
        
    intern =  lang
    intern = file.readline().strip()
    while True:
        intern = file.readline().strip()
        if intern == 'Computer Skills':
            break
        elif intern == '' or intern =='International Experience':
            intern = file.readline().strip()
            continue
        else:            
            dict["#INTERNATIONAL"] = intern
            break  
        
    compSkill =  intern
    compSkill = file.readline().strip()
    while True:
        compSkill = file.readline().strip()
        if compSkill == 'Software':
            break
        elif compSkill == '' or compSkill =='Computer Skills':
            compSkill = file.readline().strip()
            continue
        else:            
            dict["#COMPUTERSKILLS"] = compSkill
            break  
        
    software =  compSkill
    software = file.readline().strip()
    while True:
        software = file.readline().strip()
        if software == 'Hardware':
            break
        elif software == '' or software =='Software':
            software = file.readline().strip()
            continue
        else:            
            dict["#SOFTWARE"] = software
            break 
        '''
    
