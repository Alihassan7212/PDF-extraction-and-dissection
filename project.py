import fitz
import io
from PIL import Image
import tabula
import pdfplumber
import csv
import os
from fpdf import FPDF
import time
from textblob import TextBlob

os.system ("color")
#Create a nice looking message using colours
def intro():
    print("\033[1m" + "\t\t\t****************************")
    print("\t\t\t*\t\t\t   *")
    print("\t\t\t*\t"+"\033[4;31m"+"PDF"+ "\033[0m" + "\033[1m" +"DISSECTOR\t   *")
    print("\t\t\t*\t\t\t   *")
    print("\t\t\t****************************\n" + "\033[0m")
intro()

print("This program extracts text, tables, document info, and can split or rotate pdf pages.")
print("Enter the name of the pdf you want to dissect e.g" + "\033[1;33m" + " example.pdf")
print("[+] NOTE: The pdf file must be in the same directory as the program and have correct case")
print("[+] NOTE 2: Sometimes the program fails to find the file even if it is present. In that case simply retry again" + "\033[0m")

#Take the filename as input and pass it on. In case of wrong input; exit.
try:
    file = input("->FileName: ")
    file2 = file
    pdf_file = fitz.open(file)
except:
    os.system('cls')    #Clear bloated screen
    intro()
    print("Sorry the file does not exist")
    input("Press enter to exit the program")
    exit()

#Code to extract images:
print("Extracting images from pdf.....")
# iterate over PDF pages
percentages_list = []
for page_no in range(len(pdf_file)):
    # get the page itself
    page = pdf_file[page_no]
    total_area = page.mediabox.get_area()
    percentage = 0
    image_list = page.get_images()
    # printing number of images found in this page
    for image_index, img in enumerate(page.get_images(), start=1):
        # get the XREF of the image
        xref = img[0]

        image_data = page.get_image_rects(xref)
        for object in image_data:
            percentage += ((object.get_area()) / total_area) * 100
        
        # extract the image bytes
        base_image = pdf_file.extract_image(xref)
        image_bytes = base_image["image"]
        # get the image extension
        image_ext = base_image["ext"]
        # load it to PIL
        image = Image.open(io.BytesIO(image_bytes))
        # save it to local disk
        image.save(open(f"image{page_no+1}_{image_index}.{image_ext}", "wb"))
    percentages_list.append(percentage)

#Code to convert the extracted images to pdf:
folder = os.getcwd()
files = os.listdir(folder)

list = []
for file in files:
    if file.endswith(".jpeg"):
        i = Image.open(file)
        n = i.convert('RGB')
        list.append(n)
    if file.endswith(".png"):
        i = Image.open(file)
        n = i.convert('RGB')
        list.append(n)
try:
    list.pop()
    n.save('All_Images.pdf', save_all=True, append_images = list)
except:
    print('')
time.sleep(1)

#Secret code to hide images xD
for file in files:
    if file.endswith(".jpeg"):
        os.remove(file)
    if file.endswith(".png"):
        os.remove(file)



#Code for MetaDeta/Document info:
print("Extracting document information from pdf.....")
out = open(file + "_metadata" + ".txt", "w")  # open text output
data = pdf_file.metadata    #Store the metadata in a variable
for key,value in data.items():     #Access the key and its value since data is in a dictionary
    out.write(str(key) + " ----> " + str(value) + "\n")
out.close()

time.sleep(1)

#Code for converting doc_info to pdf:
doc_info = file + "_metadata" + ".txt"
with open(doc_info, "r") as info:
    content = info.read()
pdf = FPDF()
pdf.add_page()
pdf.set_font('Arial', size=12)
pdf.multi_cell(0,5,content)
pdf.output("MetaData.pdf")
#Code to remove doc_info and text:
os.remove(doc_info)


print("Extracting tables.....") 
try:
    tables = tabula.read_pdf(file, pages="all", )
    folder_name = "tables"
    if not os.path.isdir(folder_name):
        os.mkdir(folder_name)
    # iterate over extracted tables and export as excel individually
    for i, table in enumerate(tables, start=1):
        table.to_excel(os.path.join(folder_name, f"table_{i}.xlsx"), index=False)
except:
    print('')
#Perform Sentiment analysis on text
print("There are " + str(len(pdf_file)) + " Pages in the pdf file")
number = int(input("Write the page number of the pdf on which you want to perform Sentiment Analysis: "))
while ((number < 1) or (number > len(pdf_file))):
    number = int(input("You have written a wrong number. Please write correct page number: "))
page = pdf_file[number - 1]
text = page.get_text("text")
textobject = TextBlob(text)
textSentiment = textobject.sentiment.polarity

if textSentiment > 0:
    sentiment = "Positive"
# If the sentiment is negative, return "Negative"
elif textSentiment < 0:
    sentiment = "Negative"
# If the sentiment is neutral, return "Neutral"
else:
    sentiment = "Neutral"


#Code for extracting word count:
print("Extracting Text.....")
out2 = open(file + ".txt", "wb")  # open text output
WordCount = 0
for page in pdf_file:
    text = page.get_text("text")
    words = text.split(' ')
    WordCount += len(words) 
    out2.write(text.encode("utf-8"))
out2.close()


#Code for writing summary:
os.system('cls')        #Clear bloated screen
intro()
print("\033[1;33m" + "\n\t\tSummary" + "\033[0m")
image_count = 0
for page_no in range(len(pdf_file)):
    # get the page itself
    page = pdf_file[page_no]
    image_list = page.get_images()
    # printing number of images found in this page
    if image_list:
        print(f"[+] Found {len(image_list)} images in page {page_no}")
        image_count += len(image_list)
    else:
        print("[!] No images found on page", page_no)
print("Total no of images in the PDF: " + str(image_count))
print("Document info has been extracted.")
print("Text has been extracted")
try:
    print("Total number of tables in PDF: " + str(len(tables)))
except:
    print("No tables found")
print("Sentiment obtained from page " + str(number) + " was: " + sentiment)
print("Total Word count was: " + str(WordCount))

for i in range(len(pdf_file)):
    print("Percentage occupied by graphics in page " + str(i+1) + " = " + str(percentages_list[i]))

input("\nPress enter to exit the program....")    #Wait for user input to end program
quit()