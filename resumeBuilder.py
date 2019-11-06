from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
import json

def create_list(paragraph, list_type):
    p = paragraph._p #access to xml paragraph element
    pPr = p.get_or_add_pPr() #access paragraph properties
    numPr = OxmlElement('w:numPr') #create number properties element
    numId = OxmlElement('w:numId') #create numId element - sets bullet type
    numId.set(qn('w:val'), list_type) #set list type/indentation
    numPr.append(numId) #add bullet type to number properties list
    pPr.append(numPr) #add number properties to paragraph


def loadJsonData(jsonFile):
    input_file = open(jsonFile)
    json_array = json.load(input_file)
    input_file.close()
    return json_array

def replace_in_line(paragraph,replace,content):
    toSkip = len(replace)
    pos = paragraph.text.index(replace)
    paragraph.text = paragraph.text[:pos]+content+paragraph.text[pos+toSkip:]


def check_new_section(currentOrder,possibleSection):
    try: 
        if currentOrder[1] == possibleSection:
            #Means we hit a new section
            currentOrder.pop(0)
    except IndexError:
        pass

    return currentOrder
    
def parse_json_basic(json,block,content):
    for i in json[block]:
        return i[content]

def order_of_blocks(document):
    order = []
    #Allows us to easily assess what section we are in 
    for paragraph in document.paragraphs:
        text = paragraph.text.lower()
        if text == "education":
            order.append("education")
        if text == "certifications":
            order.append("certifications")
        if text == "skills" or text == "skills & abilities":
            order.append("skills")
        if text == "links":
            order.append("links")
        if text == "experience":
            order.append("experience")
        if text == "projects":
            order.append("projects")
    return order

json = loadJsonData('data/NoSqlSchema.json')
f = open('templates/template1.docx', 'rb')
#opens the document
document = Document(f)

usedDoubleMajor = False
styles = document.styles
paragraph_styles = [
    s for s in styles if s.type == WD_STYLE_TYPE.PARAGRAPH
 ]
style = styles['Heading 2']
        
#First Step,
order = order_of_blocks(document)

for paragraph in document.paragraphs:
    #rint(paragraph.text)
    text = paragraph.text.lower()
    order = check_new_section(order, text)
    #Name
    if 'firstname' in text: 
        paragraph.text = json["FirstName"]+ " " +json["LastName"]

    if 'address' in text:
        try:
            addressPos = text.index('address')
            #There is an address in the database
            address = json["Address"]
            paragraph.text = paragraph.text[:addressPos]+address+paragraph.text[addressPos+7:]
        except KeyError:
            #there is no address provided
            address = ""
            paragraph.text = paragraph.text[:addressPos]+address+paragraph.text[addressPos+10:]

    if 'phone' in text:
        replace_in_line(paragraph,"Phone",json['Phone']) 

    if 'email' in text:
        replace_in_line(paragraph,"Email",json['Email'])

    if 'summarytext' == text:
        paragraph.text = json["SummaryText"]

    if 'degree' in text:
        replace_in_line(paragraph,"Degree",parse_json_basic(json,"Education","Degree"))

    if 'graddate' in text and order[0] == "education":
        replace_in_line(paragraph,"GradDate",parse_json_basic(json,"Education","GradDate"))
    
    if 'school' in text and order[0] == "education":
        replace_in_line(paragraph,"School",parse_json_basic(json,"Education","School"))
    
    if 'major' in text and order[0] == "education":
        #just a check to see if they double majored in anything
        major = parse_json_basic(json,"Education","Major")
        paragraph.text += " "+major[0] 

    if 'minor' in text and order[0] == "education":
        minorList = parse_json_basic(json,"Education","Minor")
        for minors in range(len(minorList)):
            paragraph.text += minorList[minors] + ", "
        paragraph.text = paragraph.text[:-2]

    if "education" not in order and usedDoubleMajor == False:
        doubleMajor = parse_json_basic(json,"Education","Major")
        if len(doubleMajor) > 1:
            doubleDegree = True
        if doubleDegree == True:
            for degree in range(1,len(doubleMajor)):
                newDegree = paragraph.insert_paragraph_before(parse_json_basic(json,"Education","Degree")+" | "+parse_json_basic(json,"Education","GradDate")+" | " +parse_json_basic(json,"Education","School"))
                newDegree.style = style
                font = style.font
                font.name = 'Cambria'
                font.bold = True
                font.size = Pt(11)
                newMajor = paragraph.insert_paragraph_before("Major: "+doubleMajor[degree],style='List Bullet')
                create_list(newMajor, "1")
        usedDoubleMajor = True

    


        
                
            
            

            
            
        
        

    
        #replace_in_line(paragraph,"Degree",)



document.save("demo.docx")
        


"""

    if "skills & abilities" in text:
        if len(degrees) > 1:
            for degree in range(1,len(degrees)):
                newDegree = paragraph.insert_paragraph_before(degrees[degree].strip())
                newDegree.style = style
                font = style.font
                font.name = 'Cambria'
                font.bold = True
                font.size = Pt(11)

                newMajor = paragraph.insert_paragraph_before("Major:",style='List Bullet')
                create_list(newMajor, "1")
                
                newMinor = paragraph.insert_paragraph_before("Minor:",style='List Bullet')
                create_list(newMinor,"1")




For columns
section = document.sections[0]

sectPr = section._sectPr
cols = sectPr.xpath('./w:cols')[0]
cols.set(qn('w:num'),'2')

    #input()
    ## For Editing 
    #text = paragraph.text.lower()
    #if "project name" in text:
     #   print(paragraph.text)
     #   paragraph.text = "moo"


    #For removing stuff

    #if paragraph.text == "Project2":
     #   p = paragraph._element
      #  p.getparent().remove(p)
       # p._p = p._element = None
"""