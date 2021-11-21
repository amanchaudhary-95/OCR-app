import numpy as np
import cv2
import pytesseract
from PIL import Image
import pandas as pd
import streamlit as st
import easyocr
import io
import docx 
import re
import fitz

buffer_excel = io.BytesIO()
buffer_word = io.BytesIO()

icon = Image.open('AP.jpg')
st.set_page_config(
     page_title="OCR APP",
     page_icon=icon,
     layout="wide",
     initial_sidebar_state="expanded",
     menu_items={
         'About': "## OCR APP\n **Developed by** : Aman Chaudhary"
     })
        

def save_excel():
    data = { "RM Name" : [],
    "RM Percentage" : []
    }
    #data2 = ["Aman Chaudhary 90.2%", "Name abc 0.4%", "Chaudhary def 0.6%"]
    z = ""
    k = ""
    
    for x in st.session_state.OCR_text:
        if "\n" in x:
            #k = Decimal(k)
            data["RM Percentage"].append(k)
            data["RM Name"].append(z)
            z = ""
            k = ""
        else:
            if x.isdigit() == True:
                k = k + x
            else:
                if x == ".":
                    k = k + x
                #elif x == "%":
                #   k = k + x
                else:
                    z = z + x     
    df = pd.DataFrame(data, columns=["RM Name", "RM Percentage"])
    df=df.applymap(lambda a : a.encode('unicode_escape').decode('utf-8', 'replace') if isinstance(a, str) else a)
    #return df.to_csv(index=False).encode('utf-8')
    return df

@st.cache
def tess_ocr(img):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
    reader = pytesseract.image_to_string(img)
    return reader

@st.cache
def easy_ocr(img):
    reader = easyocr.Reader(['en','hi'], gpu=True) # this needs to run only once to load the model into memory
    result = reader.readtext(img, detail=0, paragraph=True, text_threshold=0.8)
    return result


def pdf_image(pdf_file,image_list):
    c=0
    for image_number in image_list:
        base_image=pdf_file.extractImage(image_number[0])
        st.write("Image Number", c+1)
        image_bytes = base_image['image']
        img = Image.open(io.BytesIO(image_bytes))
        image_ext = base_image['ext']
        st.image(img)
        c+=1
    


## Main code starts From here
st.title("OCR APP")
st.sidebar.selectbox("Choose OCR Engine",options=["Google Tesseract","Easy OCR"],index=0, key='engine')
files_upload=st.sidebar.file_uploader("Uploade file to perform OCR on it.",key='uploaded_file', accept_multiple_files=True, type=['jpg','jpeg','png','pdf'])

if len(files_upload)==0:
    st.info('### Upload file(s). \nTip: Select multiple files by pressing "ctrl" key')

else:
    for file in files_upload:
        
        if file.type == "image/jpeg" or file.type == "image/png":
            img = Image.open(file.name)
            st.write("### Uploaded Image")
            
        elif file.type == "application/pdf":
            pdf_file = fitz.open(file.name)
            image_list = []
            for page_number in range(pdf_file.page_count):
                # get the page itself
                page = pdf_file.load_page(page_number)
                image_list = image_list + page.getImageList()
                text_list = page.get_text('blocks')
            pdf_image(pdf_file,image_list)
            st.number_input("Enter the image number to perform OCR on it",min_value=1, max_value=len(image_list), step=1, value=1, key='image_number')
            image_number = image_list[st.session_state.image_number-1]
            base_image=pdf_file.extractImage(image_number[0])
            image_bytes = base_image['image']
            img = Image.open(io.BytesIO(image_bytes))      
        

        st.image(img, channels="BGR", caption="")
        if st.session_state.engine == 'Google Tesseract':
            result = tess_ocr(img)
        elif st.session_state.engine == 'Easy OCR':
            result1 = easy_ocr(img)
            result = ""
            for j in result1:
                result = result + "\n" + j
        
        st.text_area("Text Extracted after PERFORMING OCR by {}".format(st.session_state.engine),result,height=250, key='OCR_text')
        col1, col2, col3= st.columns([1.45,1.45,1])
        col1.download_button("Download as Text file",st.session_state.OCR_text, file_name=file.name[:-4]+".txt",mime="text/plain")
        excel_data= save_excel()
        excel_data.to_excel(buffer_excel,index=False)
        col2.download_button("Download as Excel file", data=buffer_excel, file_name=file.name[:-4]+".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        doc = docx.Document()
        temp_text = st.session_state.OCR_text
        if st.session_state.engine =='Google Tesseract':
            temp_text = re.sub(r'[^\x00-\x7f]',r'', temp_text)
            temp_text = temp_text[:-1]
        doc.add_paragraph(temp_text)
        doc.save(buffer_word)
        col3.download_button("Download as Word file", data=buffer_word, file_name=file.name[:-4]+".docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.template")
        
