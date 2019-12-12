# --------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------
# --------------------------------------- Docx Testing ---------------------------------------------
# --------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------

import docx

# Identifying the file location of the template
file = r'C:\Users\us59114\Desktop\Grant Thornton\Resume\Reformat\Target Resume Template2.docx'
# Identifying an output file location
fileout = r'C:\Users\us59114\Desktop\Grant Thornton\Resume\Output Testing.docx'

# Example of data.. would be replaced by the output from the scraping
new = {'#NAME':'Brian C. Sullivan', 
       '#EDUCATION':'B.B.A., Economics, James Madison University, 2018, Minor / Concentration: Concentration in Financial Economics',
       '#TRAININGS':'Microsoft Office Specialist Excel 2013, Bloomberg Market Concept',
       '#CLEARANCE':'',
       '#YEARSEXPERIENCE':'> 1 Year',
       '#INTRODUCTION': 'Mr. Sullivan is an Advisory Associate with Grant Thornton LLP’s Public Sector Practice in Arlington, VA and is aligned with the Advanced Digital Technology and Analytics service line within the practice. He holds a Bachelor of Business Administration (B.B.A.) degree in Economics with a concentration in Financial Economics from James Madison University. During his time as Business Analyst, he has gained experience in data analysis and visualization and supporting evaluations of financial management systems and processes. Through his education and work experience, Mr. Sullivan has developed skills in data visualization, data analytics, econometric modeling, program evaluation, and client services. He has experience with computer programs such as Microsoft PowerBI, Tableau, Qlik Sense, SAS, and the MS Office Suite, as well as open-source programming languages such as Python, R, SQL, VBA and HTML.',
       '#EXPERIENCE': 'Mr. Sullivan is a member of the Tableau Visualization Support team within OPIA. This effort includes utilizing Tableau, a data visualization tool, to create various reports required both internally for OPIA and USPTO stakeholders at various leadership levels as well as for the external consumption by the public. Mr. Sullivan leads the requirements gathering, development and refresh processes of all reports and more than 12 dashboards from the data validation and staging process to the final deliverable. Mr. Sullivan also developed numerous VBA Macros and Python scripts to include ETL scripts, semi-automated reporting tools, and a PDF scraping tool in order to support dashboard refresh processes, data collection efforts and ad hoc reporting requests from OPIA stakeholders.',
       '#AGENCY':'Department of Commerce, United States Patent and Trademark Office (USPTO), Office of Policy and International Affairs (OPIA)',
       '#PROJECTDATES': '(February 2019 – Present)',
       '#TITLES':'',
       '#SUMMARY':'Experienced data analytics professional with demonstrated ability developing Key Performance Indicators (KPI), data mining, data preparation, and data visualization.',
       '#CERTIFICATIONS':'',
       '#LASTNAME':'Sullivan',
       '#TITLE':''}

def replace_string(file):
    # Open file
    doc = docx.Document(file)
    # Loop through fields in dictionary
    for field in new:
        #Loop through paragraphs in document
        for p in doc.paragraphs:
            # If field is in paragraph, loop through the runs
            if field in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    #Replace the key with the desired text
                    if field in inline[i].text:
                        text = inline[i].text.replace(field, new[field])
                        inline[i].text = text 
        # Loop through the tables in the document
        for table in doc.tables:
            # Loop through the rows in the table
            for row in table.rows:
                # Loop through the cells in the row
                for cell in row.cells:
                    # Loop through the paragraphs in the cell
                    for p in cell.paragraphs:
                        # If feild is in paragraph, loop through the runs
                        if field in p.text:
                            inline = p.runs
                            for i in range(len(inline)):
                                # Replace the key with the desired text
                                if field in inline[i].text:
                                    text = inline[i].text.replace(field, new[field])
                                    inline[i].text = text                       
    # Save to fileout
    doc.save(fileout)
    return 

replace_string(file)

# --------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------
# ------------------------------------------ End Script --------------------------------------------
# --------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------
# ---------------------------------------- Merging two files ---------------------------------------
# --------------------------------------------------------------------------------------------------

# combines files while maintaining formatting
## Needs to be adjusted to use all files that were created. Start each resume on new page?
doc1 = docx.Document(r'C:\Users\us59114\Desktop\Grant Thornton\Resume\Output Testing.docx')
doc2 = docx.Document(r'C:\Users\us59114\Desktop\Grant Thornton\Resume\Output Testing2.docx')
for element in doc2.element.body:
    doc1 .element.body.append(element)
doc1.save('new.docx')

# --------------------------------------------------------------------------------------------------
# --------------------------------- Duplicates tags for repetitive fields --------------------------
# --------------------------------------------------------------------------------------------------

testing = {'EXPERIENCE1':{
               '#AGENCY':'Agency 1',
               '#PROJECTDATES':'Dates 1',
               '#EXPERIENCE': 'Description 1'},
           'EXPERIENCE2':{
               '#AGENCY':'Agency 2',
               '#PROJECTDATES':'Dates 2',
               '#EXPERIENCE': 'Description 2'},
          'EXPERIENCE3':{
               '#AGENCY':'Agency 3',
               '#PROJECTDATES':'Dates 3',
               '#EXPERIENCE': 'Description 3'}
        }

# returns the numbers for paragraphs that contain an experience related field
def getParaNumbers(file):
    # Open file
    doc = docx.Document(file)
    para = []
    # Loop through fields in dictionary
    for field in testing['EXPERIENCE1']:
        count = -1
        #Loop through paragraphs in document
        for p in doc.paragraphs:
            count = count + 1
            # If field is in paragraph, loop through the runs
            if field in p.text:
                para.append(count)
    return para
para = getParaNumbers(file)

# removing duplicates
para = list(dict.fromkeys(para))

# Gets formatting properties of template runs - will need to add more
def clone_run_props(tmpl_run, this_run):
    this_run.bold = tmpl_run.bold
    this_run.italic = tmpl_run.italic
    this_run.underline = tmpl_run.underline
    this_run.font.name = tmpl_run.font.name
    this_run.font.size = tmpl_run.font.size
    this_run.font.color.rgb = tmpl_run.font.color.rgb

doc = docx.Document(file)
# For each experience in the dictionary
for x in range(0, len(testing)-1):
    # Add a blank paragraph
    doc.add_paragraph()
    # For each paragraph in the copy range
    for i in para:
        # point to the copy paragraph
        template_paragraph = doc.paragraphs[i]
        # add a new paragraph
        new_paragraph = doc.add_paragraph()
        count =0
        # for each run in the copy paragraph
        for run in template_paragraph.runs:
            # add a new run in the new paragaph
            cloned_run = new_paragraph.add_run()
            # assign the properties of the copy to the new
            clone_run_props(run, cloned_run)
            # assigns the text of the run
            cloned_run.text = doc.paragraphs[i].runs[count].text
            count = count +1

doc.save(r'C:\Users\us59114\Desktop\Grant Thornton\Resume\Output Testing.docx')    
    