import pandas as pd
import math
"""
script ide:

load fil hvor svar gemmes:

Første side har variabel oversigt. Loop gennem disse.
    For hver:
	    Gå til sheet i excel:
	    #Variabelnavn A5. dvs første variabelnavn er  i A6:
		    Loop ned gennem variabelnavne:
			    Loop gennem årene:
                For hver: Gem register + variabelnavn + år + newline
                #blanke år har ”.”
                #DREAM (variabelnavn) virker anderledes

"""


#make sure path is to most recent version!!


save_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\'
load_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018.xlsx'

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
create_register_list(save_path,load_path)