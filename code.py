import io
from PIL import Image
import pytesseract
from wand.image import Image as wi
import re
import xlwt
import glob
import nltk
from nltk.corpus import stopwords





# Get the pdf files in the current directory
def get_pdf_files():
    files=glob.glob("./*.pdf")
    return files





#Function for parsing the pdf
def extract_data(path):
 pdf = wi(filename = path, resolution = 300)
 pdfImage = pdf.convert('jpeg')

 imageBlobs = []

 img=pdfImage.sequence[0]
 imgPage = wi(image = img)
 imageBlobs.append(imgPage.make_blob('jpeg'))

 recognized_text = []

 for imgBlob in imageBlobs:
  im = Image.open(io.BytesIO(imgBlob))
  text = pytesseract.image_to_string(im, lang = 'eng')
  recognized_text.append(text)

 return  str(recognized_text)




def pdf_excel():

    #fetching all the pdf files
    files=get_pdf_files()


    #opening a new excel book
    wb=xlwt.Workbook()
    ws=wb.add_sheet("Sheet 1")

    #setting the width of excel sheet columns for accomodating long fields
    ws.col(0).width=7000
    ws.col(2).width=7000


    #iterating through all the pdf files
    for i in range(len(files)):
        print ("PROCESSING "+str(i+1)+" OUT OF "+str(len(files))+"PDF'S")

        #the next function gets the text from the pdf
        parsed_text=extract_data(files[i])



        #THE REMAINING CODDE EXTRACTS DATA FROM PDF
        formatted_text=parsed_text.split("Detailed",1)[0]
        inser_data=extract_fields(formatted_text)
        for j in range(len(inser_data)):
            ws.write(i,j,inser_data[j])   #WRITTING TO EXCEL SHEET


    print ("FILE SAVED AS  \'parsed.xls\'")
    wb.save('parsed.xls')   #SAVING TO EXCEL SHEET





def extract_fields(text):

    text = text.replace('\\n', " ")  #removing \\n character from parsed text

    text_values=text.split("Total",2)[2]
    values=re.findall(r'\d+\.\d+|\d\d|\d',text_values)  #finding scores



    raw_tokens=nltk.word_tokenize(str(text))
    stop_words=set(stopwords.words('english'))

    tokens=[]
    for w in raw_tokens:
        if w not in stop_words:
            tokens.append(w)

    tokens_with_pos=nltk.pos_tag(tokens)   #getting all details with stop words removed


    nouns=[]

    for word,pos in tokens_with_pos:
        if(pos=='NNP'or pos=='NN'):
            nouns.append(word)       #getting nouns




    name=nouns[2]    # getting name of teacher ( which is a noun)


    subject=""
    i=5                #SHORTCUT FOR GETTING SUBJECT NAME IN FULL
    while(nouns[i]!='Semester'):
        subject+=nouns[i]
        i+=1



    year=re.findall(r'20\d\d',text)  # regex for finding year of teaching


    ret=[]
    ret.append(name)
    ret.append(subject)
    ret.append(year[0])
    ret=ret+values
    return ret

# the above sequence of codes returns a list with all data to be inserted into excel






pdf_excel()