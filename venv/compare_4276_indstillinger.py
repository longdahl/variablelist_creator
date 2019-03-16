import pandas as pd
import math
import os
import xlsxwriter
import numpy as np
import pickle

#make sure path is to most recent version!!


arbejde = 0

if arbejde == 1:
    save_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\'
    load_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018_newname.xlsx'
    dst_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\"
else:
    save_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\variable_lists\\"
    load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018.xlsx"
    dst_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\"
def create_register_list(save_path,load_path):
    path_4276 = load_path
    dict_4276 = pd.read_excel(path_4276,None)

    f = open(save_path + "out.txt", "w+")
    Register_list = dict_4276['Oversigt']["Unnamed: 1"]
    for reg in Register_list:
        print(reg)
        try:
            if math.isnan(reg):
                continue
        except:
            pass

        if reg == "Register":
            continue

        register_lookup = dict_4276[reg]
        num_rows = register_lookup.shape[0]

        num_cols = register_lookup.shape[1]

        for row in range(0,num_rows):
            var = register_lookup.iloc[row]['Dataset='+reg]

            try:
                if math.isnan(var):
                    continue
            except:
                pass
            if var == reg or var == "Variabelnavn":
                continue
            for col in range(1,num_cols):

                try:
                    year = register_lookup.iloc[row]['Unnamed: ' + str(col)]
                except:
                    print(year)
                    print(type(year))
                    continue
                if row == 4: #correct issues with a double row thats always present here.
                    year = register_lookup.iloc[row-1]['Unnamed: ' + str(col)]
                try:
                    if math.isnan(year):
                        continue
                except:
                    pass
                if type(year) == str:
                    if year != ".":
                        print("a cell value is type string but not . ")
                    continue
                year = int(year)
                year = str(year)
                with open(save_path + 'out.txt', 'a') as the_file:
                    the_file.write(reg + " " + var + " " + year + '\n')

    f.close()
create_register_list(save_path,load_path)


def renamer(load_path):
    for file in os.listdir(load_path):
        if file == "LPR_FOEDSLER2.xlsx":
            os.rename(load_path+file,load_path + "LPRFOEDS.xlsx")
            continue
        if file == "LPRSKSOPR.xlsx":
            os.rename(load_path + "LPR_SKSOPR2.xlsx",load_path + "LPRSKSOP.xlsx")
            continue
        if file == "LPRSKSUBE.xlsx":
            os.rename(load_path + "LPR_SKSUBE2.xlsx",load_path + "LPRSKSUB.xlsx")
            continue
        if file == "LPRUDTILSGH.xlsx":
            os.rename(load_path + "LPR_UDTILSGH2.xlsx",load_path + "LPRUDTIL.xlsx")
            continue
        if file == "DRGPSYK_AMB2.xlsx":
            os.rename(load_path + "DRGPSYK_AMB2.xlsx",load_path + "DRGPAMB.xlsx")
            continue
        if file == "DRGSOMA_AMB2.xlsx":
            os.rename(load_path + "DRGSOMA_AMB2.xlsx",load_path + "DRGSAMB.xlsx")
            continue
        if file == "DRGSOMA_HEL2.xlsx":
            os.rename(load_path + "DRGSOMA_HEL2.xlsx",load_path + "DRGSHEL.xlsx")
            continue
        new_file = file.replace("2","")
        new_file = new_file.replace("_","")
        os.rename(load_path + file,load_path + new_file)
#renamer(dst_load_path)


def dst_list_creator(load_path,save_path):
    f = open(save_path + "dst_out.txt", "w+")
    for file in os.listdir(load_path):

        excel = pd.read_excel(load_path+file)

        #temp = pickle.loads(pickle.dumps(excel.iloc[2,3::])) hacky deep copy


        excel.to_excel(load_path + file)
        reg = file[0:-5]
        print(reg)
        for k in range(2,excel.shape[0]):
            #print(k)
            var = excel.iloc[k,1]

            for i in range(3,excel.shape[1]):
                #print(i)
                year = excel.iloc[k, i]
                if year == "." or pd.isnull(year):
                    continue
                year = str(int(year))
                f.write(reg + " " + var + " " + year + '\n')
    f.close()
#dst_list_creator(dst_load_path,save_path)

reg = "aefb"
def create_var_list(reg,list_load_path,save_path):
    #Todo første række i excel 4276 er ikke med
    reg = reg.upper()
    file_dst = open(list_load_path + "dst_out.txt", "r")
    file_4276 = open(list_load_path + "out.txt", "r")

    temp_dst = open(list_load_path + "temp_dst.txt","w")
    for line in file_dst:
        if line.startswith(reg):
            temp_dst.write(line)
    temp_dst.close()

    temp_4276 = open(list_load_path + "temp_4276.txt", "w")
    for line in file_4276:

        if line.startswith(reg):
            temp_4276.write(line)
    temp_4276.close()

    temp_dst = open(list_load_path + "temp_dst.txt", "r")
    temp_4276 = open(list_load_path + "temp_4276.txt", "r")

    file_match = set(temp_4276).intersection(temp_dst)


    with open(list_load_path + 'result.txt', 'w') as file_out:
        for line in sorted(file_match):
            file_out.write(line)
    file_out.close()
    temp_dst.close()
    file_out.close()

    temp_dst = open(list_load_path + "temp_dst.txt", "r")
    workbook = xlsxwriter.Workbook(save_path + 'demo.xlsx')
    worksheet = workbook.add_worksheet(reg)
    dst_format = workbook.add_format()
    dst_format.set_bg_color('red')
    ivoe_format = workbook.add_format()
    ivoe_format.set_bg_color('yellow')
    prev_line = ""
    row = 1
    lowest = 10000

    for line in temp_dst:
        if lowest > int(line.split(" ")[2]):
            lowest = int(line.split(" ")[2])
    temp_dst = open(list_load_path + "temp_dst.txt", "r")
    for line in temp_dst:
        if prev_line == line.split(" ")[1]:
            col = int(line.split(" ")[2]) - lowest + 1
            worksheet.write(row, col, line.split(" ")[2], dst_format)
        else:
            col = int(line.split(" ")[2]) - lowest + 1
            row = row + 1
            worksheet.write(row, 0, line.split(" ")[1], dst_format)
            worksheet.write(row, col, line.split(" ")[2], dst_format)


        prev_line = line.split(" ")[1]
    temp_dst.close()
    match = open(list_load_path + 'result.txt')
    prev_line = ""
    row = 1
    for line in match:
        if prev_line == line.split(" ")[1]:
            col = int(line.split(" ")[2]) - lowest + 1
            worksheet.write(row, col, line.split(" ")[2], ivoe_format)
        else:
            col = int(line.split(" ")[2]) - lowest + 1
            row = row + 1
            worksheet.write(row, col, line.split(" ")[2], ivoe_format)
        prev_line = line.split(" ")[1]

        prev_line = line.split(" ")[1]
    workbook.close()



#create_var_list(reg,list_load_path,save_path)