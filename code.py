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

    files=get_pdf_files()

    wb=xlwt.Workbook()
    ws=wb.add_sheet("Sheet 1")

    ws.col(0).width=7000
    ws.col(2).width=7000


    for i in range(len(files)):
        print ("PROCESSING "+str(i+1)+" OUT OF "+str(len(files))+"PDF'S")
        parsed_text=extract_data(files[i])
        formatted_text=parsed_text.split("Detailed",1)[0]
        inser_data=extract_fields(formatted_text)
        for j in range(len(inser_data)):
            ws.write(i,j,inser_data[j])


    print ("FILE SAVED AS  \'parsed.xls\'")
    wb.save('parsed.xls')





def extract_fields(text):

    text = text.replace('\\n', " ")

    text_values=text.split("Total",2)[2]
    values=re.findall(r'\d+\.\d+|\d\d|\d',text_values)



    raw_tokens=nltk.word_tokenize(str(text))
    stop_words=set(stopwords.words('english'))

    tokens=[]
    for w in raw_tokens:
        if w not in stop_words:
            tokens.append(w)

    tokens_with_pos=nltk.pos_tag(tokens)


    nouns=[]

    for word,pos in tokens_with_pos:
        if(pos=='NNP'or pos=='NN'):
            nouns.append(word)




    name=nouns[2]


    subject=""
    i=5
    while(nouns[i]!='Semester'):
        subject+=nouns[i]
        i+=1



    year=re.findall(r'20\d\d',text)


    ret=[]
    ret.append(name)
    ret.append(subject)
    ret.append(year[0])
    ret=ret+values
    return ret







pdf_excel()