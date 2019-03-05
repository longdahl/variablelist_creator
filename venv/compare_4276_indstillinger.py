import pandas as pd
import math
import os
import xlsxwriter


#make sure path is to most recent version!!


save_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\'
load_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018_newname.xlsx'
dst_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\from_dst\\"
list_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\"

def create_register_list(save_path,load_path):
    path_4276 = load_path
    dict_4276 = pd.read_excel(path_4276,None)

    f = open(save_path + "out.txt", "w+")
    Register_list = dict_4276['Oversigt']["Unnamed: 1"]
    for reg in Register_list:

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
        var_list = register_lookup['Dataset='+reg]

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

                # print("Unnamed: " + str(col))
                year = register_lookup.iloc[row]['Unnamed: ' + str(col)]
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
                # float to int

        # with pd.option_context('display.max_rows', None, 'display.max_columns', None):
        #     print(register_lookup)

    f.close()
#create_register_list(save_path,load_path)


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
        reg = file[0:-5]
        print(reg)
        for k in range(3,excel.shape[0]):
            var = excel.iloc[k,1]
            for i in range(4,excel.shape[1]):
                year = excel.iloc[k, i]
                if year == ".":
                    continue
                year = str(int(year))
                f.write(reg + " " + var + " " + year + '\n')
    f.close()
#dst_list_creator(dst_load_path,save_path)

reg = "aefv"
def create_var_list(reg,list_load_path,save_path):
    #Todo første række i excel 4276 er ikke med
    reg = reg.upper()
    file_dst = open(list_load_path + "dst_out.txt", "r")
    file_4276 = open(list_load_path + "out_allsheets.txt", "r")

    temp = open(list_load_path + "temp.txt","w")
    for line in file_dst:
        if line.startswith(reg):
            temp.write(line)
    temp.close()
    temp = open(list_load_path + "temp.txt", "r")
    file_match = set(file_4276).intersection(temp)
    with open(list_load_path + 'result.txt', 'w') as file_out:
        for line in file_match:
            file_out.write(line)
    file_out.close()

    # open register from DST

    #recolor according to result

    #create list of var names and years from excel file

    #loop through result and get index in the lists of each var and year. recolor that field.

    #save register


    #which input should the script take: format: reg: var(year-year), var(year-year)... \n reg: var...





    k = True

    if False == k:
        workbook = xlsxwriter.Workbook(save_path + 'demo.xlsx')
        worksheet = workbook.add_worksheet()
        dst_format = workbook.add_format()

        dst_format.set_bg_color('yellow')
        worksheet.write(2, 0, 123,dst_format)

        workbook.close()
create_var_list(reg,list_load_path,save_path)

    #change name -2
    # change DRGPSYK_AMB2,DRGSOMA_AMB2,DRGSOMA_HEL2, lpradm,lprbes,lprdiag,
    #lprfoeds -> lprfoedsler, lprsksop --> lprsksopr, lprsksub --> lprssksube, lprudtil, lprudtilsgh,
