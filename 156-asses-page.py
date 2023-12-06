

from PyPDF2 import PdfMerger
import csv
#import os
#import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue, black

# for Thonny
# Need to go to Tools >> Manage Packages
# Download and install PyPDF2 and Reportlab packages

# Make 3 folders in same directory as this program called:
# "headers", "exam_pages", and "final_results"

# Fill "exam_pages" folder with the pages of the exam in
# PDF format with names matching the objectives below such as "A1.pdf"

assess_num = input('Which assessment is this?')
outcomes = ['EQ1', 'EQ2', 'EQ3','EQ4', 'EQ5','FUN1','FUN2','FUN3','FUN4', 'FUN5', 
            'FUN6','GRAL1','GRAL2','GRAL3','GRTR1','GRTR2','GRTR3','APP1','APP2','APP3','APP4',
            'EQ1(2)', 'EQ2-2(2)', 'EQ3(2)','EQ4(2)', 'EQ5(2)','FUN1(2)',
            'FUN2(2)','FUN3(2)','FUN4(2)', 'FUN5(2)','FUN6(2)','GRAL1(2)',
            'GRAL2(2)','GRAL3(2)','GRTR1(2)','GRTR2(2)','GRTR3(2)','APP1(2)',
            'APP2(2)','APP3(2)','APP4(2)']
#score_threshold_to_include_in_final_exam = 2.0


# function to pull data from CSV files
def csv_download(name):
    data=[]
    with open(name, encoding="utf8") as f:
        reader = csv.reader(f)
        #Add each row in the csv file as a list
        for row in reader:
            data.append(row)
    return data

# setup and running CSV function
raw_data = csv_download('a'+ assess_num +'_responses.csv')
grades_data = csv_download('grades.csv')
requested = {}
grades = {}

# builds the set of objectives requested by students
i = 0
for row in raw_data:
    if i==0:
        i=1
        continue
    else:
        requested[row[1]] = {}
        requested[row[1]]['objs'] = row[4]


# builds the grades set, going through the objectives and determining what is required
### IMPORTANT
### This is also where any mandatory objectives are set (like C1 and C2)
i = 0
for row in grades_data:
    if i==0:
        i=1
        continue
    else:
        pid = row[1]
        email = row[2]
        fname = row[3]
        lname = row[4]
        status = row[5]
        # if status == 'D':
        #     continue
        # cg = row[12]
        # scn = row[0]
        if pid not in grades: 
            grades[pid] = {'email': email, 'fname': fname, 'lname': lname, 'objs': {}}
        obj = row[7]
        # print(row[11])
        try: 
            score = float(row[11])
        except: 
            score = float(0)
        grades[pid]['objs'][obj] = score
#        if score < score_threshold_to_include_in_final_exam: 
#            grades[pid]['final'][obj] = 'required'
        
        #grades_data[row[4]] = {}
        #classlist[row[4]]['pid'] = row[2]
        #classlist[row[4]]['scn'] = row[1]
        #classlist[row[4]]['name'] = row[3]


# joins the objectives that are requested and those that are required
for pid in grades: 
    email = grades[pid]['email']
    if email in requested: 
        requested_objs = requested[email]['objs']
        for i in outcomes: 
            obj = i
            if obj in requested_objs: 
                grades[pid]['objs']['Outcome '+obj] = 'requested'
                print(pid + ' requested ' + obj + ' to be in Assessment' + assess_num)

# formats objective lists to be used on the front page table
# should combine this with some previous code I'm sure but we are not worried about efficiency
for pid in grades: 
    grades[pid]['list'] = []
    objs = grades[pid]['objs']
    for obj in outcomes:
        if "Outcome "+obj in objs:
            grades[pid]['list'].append(obj)
    if len(grades[pid]['list']) > 11:
        grades[pid]['obj_list'] = " | ".join(grades[pid]['list'])
    else: 
        grades[pid]['obj_list'] = "  |  ".join(grades[pid]['list'])        


# Now build all the cover pages

#stopper= 0
for pid in grades:
    #stopper+=1
    #if stopper > 60: 
    #    break 
    email = grades[pid]['email']
    pdfs =[] 
    can = canvas.Canvas('headers/h'+assess_num+'-'+email+'.pdf', pagesize=letter)
    width, height = letter
    can.setLineWidth(.3) 
    
    can.setFillColor(blue)
    can.setFont('Helvetica', 66)
    can.drawString(50,700,'UAF')
    can.setFont('Helvetica', 30)
    can.drawString(185,730,'Department of Mathematics')
    can.drawString(185,700,'and Statistics')
    
    can.setFillColor(black)
    can.setFont('Helvetica', 24)
    can.drawString(50,660,'Math 156X / Assessment '+assess_num)
    can.line(50,656,350,656)
    can.setFont('Helvetica', 12)
    can.drawString(50,645,'Dr. Latrice Bowman / Fall 2023')
    # 
    can.setFont('Helvetica', 14)
    can.drawString(50,615,'Name: ')
    can.setFont('Helvetica', 20)
    can.drawString(100,615,grades[pid]['fname']+' '+grades[pid]['lname'])
    
    can.setFont('Helvetica', 14)
    can.drawString(50,590,"The specific Learning Outcomes in your packet are listed in the table below.")
    can.drawString(50,578,"ALG1, ALG2, and ALG3 do not have specific problems, they are based on")
    can.drawString(50,566,"the assessment in its entirety.")
 
    can.setFont('Helvetica', 16)
    can.drawString(50,530,'OUTCOMES:')

    if len(grades[pid]['list']) <= 5:
        can.saveState()
        can.translate( -20, 0 )
        can.line(70,510,70,430) # first vertical line
        for i in range(0, len(grades[pid]['list'])):
            can.drawString(80,490, 'ALG1')
            can.drawString(140,490, 'ALG2')
            can.drawString(200,490, 'ALG3')
            can.setFont('Helvetica', 10)
            can.drawString(260+60*(i),490, grades[pid]['list'][i])
            can.setFont('Helvetica', 16)
            can.line(128.5,510,128.5,430) # vertical line
            can.line(188.5,510,188.5,430) # vertical line
            can.line(248.5,510,248.5,430) # vertical line
            can.line(308.5+60*i,510,308.5+60*i,430) # vertical line
            can.line(70,510,128.5+60*(i+3),510) #horizontal line top of box
            can.line(70,480,128.5+60*(i+3),480) #horizontal line middle of box
            can.line(70,430,128.5+60*(i+3),430) #horizontal line bottom of box       
        can.restoreState()  
        
        can.setFont('Helvetica', 16)
        can.drawString(50,400,'Assessment Specific Instructions:')
        can.setFont('Helvetica', 14)
        can.drawString(50,386,'*You will have 1 hour to complete this assessment.')
        can.drawString(50,372,'*The assessment is closed books and closed notes.')
        can.drawString(50,358,'*Calculators/Phones are not allowd on this assessment.')
        can.drawString(50,344,'*In order to receive P or M on an outcome you must complete requirements')
        can.drawString(50,330,' for that outcome.')
        can.drawString(50,316,'*All answers should be completely simplified with the correct units where necessary.')
        can.drawString(50,302,'*All graphs/diagrams should be completely labeled.')
        can.drawString(50,288,'*If a specific method is asked for, credit will not be given if another method is used.')

    elif len(grades[pid]['list']) <= 13:
        can.saveState()
        can.translate( -20, 0 )
        can.line(70,510,70,310) # first vertical line
        for i in range(0, 5):
            can.drawString(80,490, 'ALG1')
            can.drawString(140,490, 'ALG2')
            can.drawString(200,490, 'ALG3')
            can.setFont('Helvetica', 10)
            can.drawString(260+60*i,490, grades[pid]['list'][i])
            can.setFont('Helvetica', 16)
            can.line(128.5,510,128.5,410) # vertical line
            can.line(188.5,510,188.5,410) # vertical line
            can.line(248.5,510,248.5,410) # vertical line
            can.line(308.5+60*(i),510,308.5+60*i,410) # vertical line
            can.line(70,510,128.5+60*(i+3),510) #horizontal line top of box
            can.line(70,470,128.5+60*(i+3),470) #horizontal line middle of box
            can.line(70,410,128.5+60*(i+3),410) #horizontal line bottom of box  
            for i in range(5, len(grades[pid]['list'])):
                can.setFont('Helvetica', 10)
                can.drawString(80+60*(i-5),390, grades[pid]['list'][i])
                can.setFont('Helvetica', 16)
                can.line(128.5+60*(i-5),410,128.5+60*(i-5),310) # vertical line
                can.line(70,370,128.5+60*(i-5),370) #horizontal line middle of box
                can.line(70,310,128.5+60*(i-5),310) #horizontal line bottom of box    
        can.restoreState()  
        
        can.setFont('Helvetica', 16)
        can.drawString(50,280,'Assessment Specific Instructions:')
        can.setFont('Helvetica', 14)
        can.drawString(50,266,'*You will have 1 hour to complete this assessment.')
        can.drawString(50,244,'*The assessment is closed books and closed notes.')
        can.drawString(50,222,'*Calculators/Phones are not allowd on this assessment.')
        can.drawString(50,200,'*In order to receive P or M on an outcome you must complete requirements')
        can.drawString(50,178,' for that outcome.')
        can.drawString(50,156,'*All answers should be completely simplified with the correct units where necessary.')
        can.drawString(50,134,'*All graphs/diagrams should be completely labeled.')
        can.drawString(50,112,'*If a specific method is asked for, credit will not be given if another method is used.')

    else:
        can.drawString(80,490, grades[pid]['obj_list'])  

        can.setFont('Helvetica', 16)
        can.drawString(50,400,'Assessment Specific Instructions:')
        can.setFont('Helvetica', 14)
        can.drawString(50,386,'*You will have 1 hour to complete this assessment.')
        can.drawString(50,372,'*The assessment is closed books and closed notes.')
        can.drawString(50,358,'*Calculators/Phones are not allowd on this assessment.')
        can.drawString(50,344,'*In order to receive P or M on an outcome you must complete requirements')
        can.drawString(50,330,' for that outcome.')
        can.drawString(50,316,'*All answers should be completely simplified with the correct units where necessary.')
        can.drawString(50,302,'*All graphs/diagrams should be completely labeled.')
        can.drawString(50,288,'*If a specific method is asked for, credit will not be given if another method is used.')

    #c.showPage()
    can.save()        
    pdfs.append('headers/h'+assess_num+'-'+email+'.pdf') 
    #pdfs.append('exam_pages/H1.pdf')
    
    
    ####
    ####
    # Now that cover page is made go through and add the other pages needed for this
    # student's exam
    ####
    for obj in outcomes:
        ehh_name = 'Outcome '+obj
        if ehh_name in grades[pid]['objs']:
                pdfs.append('a'+assess_num+'_exam_pages/'+obj+'.pdf')
    merger = PdfMerger()
    for pdf in pdfs:
        merger.append(open(pdf, 'rb'))
    #with open('final_results/'+classlist[email]['scn']+'--'+classlist[email]['name']+'--'+email+'.pdf', 'wb') as fout:
    with open('assessment_'+assess_num+'/'+grades[pid]['fname']+'_'+grades[pid]['lname']+'.pdf', 'wb') as fout:       
        merger.write(fout)            
    

# Now we make a CSV file with what exam pages are printed for each student

output = [['Email', 'First Name', 'Last Name', '#Pages', 'Assessment '+assess_num+' Outcomes']]

for pid in grades:
   #     continue
    my_row = []
    my_row.append(grades[pid]['email'])
    my_row.append(grades[pid]['fname'])
    my_row.append(grades[pid]['lname'])
    my_row.append(1+len(grades[pid]['list']))
    my_row.append(", ".join(grades[pid]['list']))
    output.append(my_row)

with open('assessment_'+assess_num+'_output.csv', 'w', newline='') as f:
    writer = csv.writer(f)     
    writer.writerows(output)     


