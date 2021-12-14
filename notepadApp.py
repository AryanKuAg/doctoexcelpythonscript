#import docx
from typing import final
from openpyxl import Workbook


doc = open("zero.txt","r") # file to import


mydata = ''

for para in doc.readlines():
   
   mydata = mydata + ' ' + para
  
 
   # print('---------------------------')


listofEntries = []


for i, elem in enumerate(str(mydata).split("~~")):
    listofEntries.append([])
    for j in str(elem).split('~'):
        data = str(j).strip()
        
        if data.find('$') != -1:
            listofEntries[i].append(" ")
            continue
        
        if data == '0':
            listofEntries[i].append(data.replace("0", " "))
        else:
            justanother = data.replace("\n", " ")
            listofEntries[i].append(justanother.replace("\t", " "))
            

#print(listofEntries)
# Program is okay uptil here ---------------------------------
            


finalfinalList = []

for i in listofEntries:
    if len(i) < 12:
        continue
    else:
        finalfinalList.append(i)


print(finalfinalList)

#purifiedData = []

#for i in mydata:
#    purifiedData.append(str(i).replace("\t", " "))

#print(purifiedData)

# Converting to excel sheet--------------------------------------------------------------------------
try:
    wb = Workbook()
    # f = open("extraction" + str(i) +".txt", "r")
    # arrangedData = dataManipulation.dataManipulation('extraction' + str(i) + '.txt') #set file name here
    # print(arrangedData)
    # Get Arranged data in list of list


    ws = wb.active

    # OUR MAIN ISSUE
    for j in finalfinalList: # purified data should be like [['','','','',''],['','','','',''],['','','','','']]

      # upper case naming and replace ' ' with double spaces 
      j[3] = str(j[3]).upper()
      j[3] = j[3].replace(" ", "  ")

      # phone number to a format
      if len(j[7]) == 10:
          j[7] = '+91 ' + j[7][0:5] + ' ' + j[7][5:]
    
      if len(j[9]) == 10:
          j[9] = '+91 ' + j[9][0:5] + ' ' + j[9][5:]


      # capitalize
      j[10] = str(j[10]).capitalize()
      j[11] = str(j[11]).capitalize()

      if len(j) > 15:
        j[15] = str(j[15]).capitalize()

      # open close upper
      j[5] = j[5].upper()
    
      
      ws.append(j)

    
    wb.save('herohero.xlsx')
    
    
 
except Exception as e:
  print("An exception occurred: ", e)
