
#Written By: Michael Giombetti
#Creation date: 04/01/2022
#Last update: 05/01/2022

from typing_extensions import OrderedDict
from PIL import Image
import pytesseract
import cv2  
import pyexcel_ods3 
import os
from pyexcel_ods3 import save_data

#Input paths: Tessract.exe / path to folder to be scanned

input_folder_path = r'C:\\Users\\micha\\Documents\\New Py project'
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'


#Amount of Files in specified path

list_files = os.listdir(input_folder_path)

#Loop through the files, searchs for .jpg files

files_toScan = []

for file_names in list_files:
    if '.jpg' in file_names:
     #print(file_names)
     files_toScan.append(file_names) #fixed, where it would only process 1 document
     #files_to_scan[]=file_names

print(files_toScan) #Print in console files .jpg

# Open Images  and change them to tesseract apt format
#image_arr = []
#P_image_arr = []
#for q in range(len(files_toScan)):
      #image_arr.append(cv2.imread(files_toScan[q]))
      #P_image_arr.append( cv2.cvtColor(image_arr,cv2.COLOR_BGR2RGB))

#for r in range(len(image_arr)):
  #  print(image_arr[r])

def openNanalyzefile(file):
    image_i = cv2.imread(file)
    P_image = cv2.cvtColor(image_i, cv2.COLOR_BGR2RGB)
    Data_raw = pytesseract.image_to_data(P_image)
    return Data_raw

print(openNanalyzefile('pago.jpg'))

with open('dump.txt', 'w') as FileWrite:
    FileWrite.write(openNanalyzefile("pago.jpg"))



#print(Data_raw)

#initialise lists
test_list = [""]
main_list = [""]

# loop through the data 
for i,j in enumerate(Data_raw.splitlines()):
    if i!=0:
        j = j.split()
        #print(j)
        if len(j)==12:   #if the lenght if greater than 11, the 12th element is the text [12-1]
            word_raw = j[11] #text 
            test_list.append(word_raw) #append text to the empty list
           # print(word_raw) #debug print

for word in range (len(test_list)): #Loop through the text list 
    #lowerCase = word.lower()
    if test_list[word] == "TOTAL": #Search for word total and append de next value total= "price"
        main_list.append(test_list[word+1])
        
    

#for x in range(len(main_list)): #print data of the collection list
 #   print(main_list[x])

Data_to_exel = OrderedDict() #Writting to ods

Or_Data = [         #data to write to the ods file
    [test_list[0]],
    [test_list[1]],
    [test_list[2]],
    [test_list[3]],
    [test_list[4]],
    [test_list[5]],
    [test_list[6]],
    [test_list[7]],
    [test_list[8]],
    
]
MainSheet = {
    "MainData":Or_Data
}
#Data_to_exel.update({"Sheet 1" : [["Proveedor","Fecha","Nº Referencia"],[word_raw[i]]]  })

save_data("NewProjectDocumentSpreadsheet.ods", MainSheet) #Save data to the ods file

#cv2.imshow("a", P_image)
#cv2.waitKey(0)
#cv2.destroyAllWindows()