#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May 19 09:54:42 2020

@author: giorgo
"""
# you need to install pdfrw first.
# "pip3 install pdfrw" should do the job.


import pandas as pd
import pdfrw
from pdfrw import PdfReader, PdfWriter
from datetime import datetime
import os
import sys

timestr = datetime.now().strftime("%d%m%Y") #day, month, year, hour, minutes, seconds, milliseconds
timestr_c = datetime.now().strftime("%d%m%Y_%H%M%S")


if len(sys.argv)==1:
    print('USAGE:  fill_form.py csv_path')
    print('   e.g. fill_form.py "/home/username/mystudy/payment/results_survey.csv"')
    print('Code assumes that the folder containing fill_form.py also contains the unmodified template KostenerstattungOnline.pdf.')
    sys.exit()

folder_template = os.path.dirname(os.path.realpath(__file__))
invoice_path = sys.argv[1]
folder_invoice = os.path.dirname(invoice_path)


#read the csv of the questionnaries
df_total = pd.read_csv(invoice_path)
df_total_2 = df_total

df_total["A[IBAN]"] = df_total["IBAN[IBANC_SQ001]"].astype(str) + df_total["IBAN[IBANC_SQ002]"].astype(str) + df_total["IBAN[IBANC_SQ003]"].astype(str) + df_total["IBAN[IBANC_SQ004]"].astype(str) + df_total["IBAN[IBANC_SQ005]"].astype(str)  + df_total["IBAN[IBANC_SQ006]"].astype(str)


if df_total['A[address]'].isnull().any():
    print("\n\nWARNING!!!\n\nThere are missing addresses in your input file!\n" + 
          "It will not be possible to process the payment unless all addresses are provided.\n" +
          "In your output folder you can find a csv file reporting participants for whom no address was provided." 
          + "\n\n----------------------")
    df_total.fillna("/", inplace=True)
    idx = df_total.apply(lambda ts: any(ts == '/'), axis=1)
    

duplicateRowsDF_name = df_total[df_total.duplicated(['A[name]'])]
duplicateRowsDF_IBAN = df_total[df_total.duplicated(['A[IBAN]'])]

if not duplicateRowsDF_name.empty:
    print("\n\nWARNING!!!\n\nThere are duplicates in your input file! " + 
          "Please check that everything is correct.\n" +
          "In your output folder you can find csv files reporting duplicates" +
          " based on names and/or IBAN.\n")

if not duplicateRowsDF_IBAN.empty:
     print("\n\nWARNING!!!\n\nThere are duplicates in your input file! " + 
          "Please check that everything is correct.\n" +
          "In your output folder you can find csv files reporting duplicates" +
          " based on names and/or IBAN.\n")



#define function for splitting the dataframe into multiple dataframes
#having maximum 4 rows each
def splitDataFrameIntoSmaller(df, chunkSize = 4):
    listOfDf = list()
    numberChunks = len(df) // chunkSize + 1
    for i in range(numberChunks):
        listOfDf.append(df[i*chunkSize:(i+1)*chunkSize])
    return listOfDf

#aplly function
df_list = splitDataFrameIntoSmaller(df_total)

#create a folder for the invoices if it does not exist
output_folder = folder_invoice + "invoice_" + timestr_c + "/"



if not os.path.exists(output_folder):
    os.makedirs(output_folder)

a = 0
for x in df_list:
    a = a + 1
    #read one df at the time
    df = pd.DataFrame(x)
    d={}
    INVOICE_OUTPUT_PATH = folder_invoice + "invoice_" + timestr_c + "/" + "invoice"
    INVOICE_TEMPLATE_PATH = folder_template + "/KostenerstattungOnline.pdf"


    for i in range(0, len(df)):

        #extract values for each df
        name = df.iloc[i]["A[name]"]
        address = df.iloc[i]["A[address]"]
        EUR =  str(df.iloc[i]["A[EUR]"])

        #date in which the survey is submitted
        #datum = str(df.iloc[i]["submitdate"])
        #convert date EU-like format DD-MM_YYYY
        #datum = datum[8:10] + "-" + datum[5:7]+  "-"+  datum[0:4]

        datum = df.iloc[i]["A[date]"]
        description  = df.iloc[i]["A[description]"]

        #split the IBNA in 6 parts
        IBAN = df.iloc[i]["A[IBAN]"]
        IBAN_1 = IBAN[0:4]
        IBAN_2 = IBAN[4:8]
        IBAN_3 = IBAN[8:12]
        IBAN_4 = IBAN[12:16]
        IBAN_5 = IBAN[16:20]
        IBAN_6 = IBAN[20:22]

        d["name_{0}".format(i)]=name
        d["EUR_{0}".format(i)]=EUR
        d["address_{0}".format(i)]=address
        d["date_{0}".format(i)]=datum
        d["IBAN{0}_1".format(i)]=IBAN_1
        d["IBAN{0}_2".format(i)]=IBAN_2
        d["IBAN{0}_3".format(i)]=IBAN_3
        d["IBAN{0}_4".format(i)]=IBAN_4
        d["IBAN{0}_5".format(i)]=IBAN_5
        d["IBAN{0}_6".format(i)]=IBAN_6
        d["description_{0}".format(i)]=description

    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_ALTERNATE_FILED_KEY = "/TU"
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'

    def write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict):
        template_pdf = pdfrw.PdfReader(input_pdf_path)
        template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

        annotations = template_pdf.pages[0][ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        annotation.update(
                            pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                        )
        pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

    #assign values to the empty boxes
    #if there are still 4 subject to be inserted
    if len(df) == 4:
        data_dict = {
           'name1': d["name_0"],
           "address1" : d["address_0"],
           "amount1" : d["EUR_0"],
           "iban1_part1" :  d["IBAN0_1"],
           "iban1_part2" :  d["IBAN0_2"],
           "iban1_part3" :  d["IBAN0_3"],
           "iban1_part4" :  d["IBAN0_4"],
           "iban1_part5" :  d["IBAN0_5"],
           "iban1_part6" :  d["IBAN0_6"],
           "description1" : d["description_0"],
           "date1" : d["date_0"],
           "name2": d["name_1"],
           "address2" : d["address_1"],
           "amount2" : d["EUR_1"],
           "iban2_part1" :  d["IBAN1_1"],
           "iban2_part2" :  d["IBAN1_2"],
           "iban2_part3" :  d["IBAN1_3"],
           "iban2_part4" :  d["IBAN1_4"],
           "iban2_part5" :  d["IBAN1_5"],
           "iban2_part6" :  d["IBAN1_6"],
           "description2" : d["description_1"],
           "date2" : d["date_1"],
           'name3': d["name_2"],
           "address3" : d["address_2"],
           "amount3" : d["EUR_2"],
           "iban3_part1" :  d["IBAN2_1"],
           "iban3_part2" :  d["IBAN2_2"],
           "iban3_part3" :  d["IBAN2_3"],
           "iban3_part4" :  d["IBAN2_4"],
           "iban3_part5" :  d["IBAN2_5"],
           "iban3_part6" :  d["IBAN2_6"],
           "description3" : d["description_2"],
           "date3" : d["date_2"],
           "name4": d["name_3"],
           "address4" : d["address_3"],
           "amount4" : d["EUR_3"],
           "iban4_part1" :  d["IBAN3_1"],
           "iban4_part2" :  d["IBAN3_2"],
           "iban4_part3" :  d["IBAN3_3"],
           "iban4_part4" :  d["IBAN3_4"],
           "iban4_part5" :  d["IBAN3_5"],
           "iban4_part6" :  d["IBAN3_6"],
           "description4" : d["description_3"],
           "date4" : d["date_3"],
        }
    #if there are still 3 subject to be inserted
    elif len(df) == 3:
        data_dict = {
           'name1': d["name_0"],
           "address1" : d["address_0"],
           "amount1" : d["EUR_0"],
           "iban1_part1" :  d["IBAN0_1"],
           "iban1_part2" :  d["IBAN0_2"],
           "iban1_part3" :  d["IBAN0_3"],
           "iban1_part4" :  d["IBAN0_4"],
           "iban1_part5" :  d["IBAN0_5"],
           "iban1_part6" :  d["IBAN0_6"],
           "description1" : d["description_0"],
           "date1" : d["date_0"],
           "name2": d["name_1"],
           "address2" : d["address_1"],
           "amount2" : d["EUR_1"],
           "iban2_part1" :  d["IBAN1_1"],
           "iban2_part2" :  d["IBAN1_2"],
           "iban2_part3" :  d["IBAN1_3"],
           "iban2_part4" :  d["IBAN1_4"],
           "iban2_part5" :  d["IBAN1_5"],
           "iban2_part6" :  d["IBAN1_6"],
           "description2" : d["description_1"],
           "date2" : d["date_1"],
           'name3': d["name_2"],
           "address3" : d["address_2"],
           "amount3" : d["EUR_2"],
           "iban3_part1" :  d["IBAN2_1"],
           "iban3_part2" :  d["IBAN2_2"],
           "iban3_part3" :  d["IBAN2_3"],
           "iban3_part4" :  d["IBAN2_4"],
           "iban3_part5" :  d["IBAN2_5"],
           "iban3_part6" :  d["IBAN2_6"],
           "description3" : d["description_2"],
           "date3" : d["date_2"],
        }
    #if there are still 2 subject to be inserted
    elif len(df) == 2:
        data_dict = {
           'name1': d["name_0"],
           "address1" : d["address_0"],
           "amount1" : d["EUR_0"],
           "iban1_part1" :  d["IBAN0_1"],
           "iban1_part2" :  d["IBAN0_2"],
           "iban1_part3" :  d["IBAN0_3"],
           "iban1_part4" :  d["IBAN0_4"],
           "iban1_part5" :  d["IBAN0_5"],
           "iban1_part6" :  d["IBAN0_6"],
           "description1" : d["description_0"],
           "date1" : d["date_0"],
           "name2": d["name_1"],
           "address2" : d["address_1"],
           "amount2" : d["EUR_1"],
           "iban2_part1" :  d["IBAN1_1"],
           "iban2_part2" :  d["IBAN1_2"],
           "iban2_part3" :  d["IBAN1_3"],
           "iban2_part4" :  d["IBAN1_4"],
           "iban2_part5" :  d["IBAN1_5"],
           "iban2_part6" :  d["IBAN1_6"],
           "description2" : d["description_1"],
           "date2" : d["date_1"],
        }
    #if there is still 1 subject to be inserted
    elif len(df) == 1:
        data_dict = {
           'name1': d["name_0"],
           "address1" : d["address_0"],
           "amount1" : d["EUR_0"],
           "iban1_part1" :  d["IBAN0_1"],
           "iban1_part2" :  d["IBAN0_2"],
           "iban1_part3" :  d["IBAN0_3"],
           "iban1_part4" :  d["IBAN0_4"],
           "iban1_part5" :  d["IBAN0_5"],
           "iban1_part6" :  d["IBAN0_6"],
           "description1" : d["description_0"],
           "date1" : d["date_0"],
        }
     
     
    if __name__ == '__main__':
        write_fillable_pdf(INVOICE_TEMPLATE_PATH, INVOICE_OUTPUT_PATH + str(a) + "_"+ timestr + ".pdf", data_dict)    

if not duplicateRowsDF_name.empty:
    duplicateRowsDF_name= duplicateRowsDF_name.sort_values('A[name]')
    duplicateRowsDF_name.to_csv(output_folder + "name.csv", index = False)

if not duplicateRowsDF_IBAN.empty:  
    duplicateRowsDF_IBAN = duplicateRowsDF_IBAN.sort_values('A[IBAN]')
    duplicateRowsDF_IBAN.to_csv(output_folder + "IBAN.csv", index = False)

if df_total_2['A[address]'].isnull().any():
    df_total[idx].to_csv(output_folder + "no_address.csv", index = False)

#this creates a concatenated file of all pdf filesif df_total['A[address]'].isnull().any():

def concatenate(paths, output):
    writer = PdfWriter()
    
    for path_x in paths:
        reader = PdfReader(path_x)
        writer.addpages(reader.pages)
        
    writer.write(output)  

paths = []
len_df_total = len(df_total)

if (__name__ == '__main__') and (len_df_total % 4 == 0):
    for y in range(1,a):
            path_x = INVOICE_OUTPUT_PATH+ str(y) + "_" + timestr + ".pdf"
            paths.append(path_x)
elif (__name__ == '__main__') and (len_df_total % 4 != 0):
    for y in range(1,a+1):
            path_x = INVOICE_OUTPUT_PATH+ str(y) + "_" + timestr + ".pdf"
            paths.append(path_x)
            
conc_file = output_folder + "/conc_inv" + timestr_c + ".pdf"
concatenate(paths, conc_file)

#this will remove the single pdf files leaving only the concatenated one
if (len_df_total % 4 != 0):
    for path_x in paths:
            os.remove(path_x)
elif (len_df_total % 4 == 0):
    for path_x in paths:
            os.remove(path_x)
    y = y+1
    os.remove(INVOICE_OUTPUT_PATH+ str(y) + "_" + timestr + ".pdf")
