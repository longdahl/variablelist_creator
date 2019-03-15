import pandas as pd
from docx import Document
import math
path_7217 = "C:\\Users\\Mikkel\\Desktop\\arbejde\\Project database\\Projects\\707217 - Arbejdstilsyn\\707217_DST_variabelliste_20180917.xlsx"
save_path = "C:\\Users\\Mikkel\\Desktop\\arbejde\\Project database\\datamanager\\script\\word_documentation\\"


dict_7217 = pd.read_excel(path_7217,None)

for key in dict_7217.keys():
    if key.startswith("E_"):
        continue
    reg = dict_7217[key]
    num_rows = reg.shape[0]

    if key != "VNDS":
        continue
    document = Document()
    document.add_heading("Documentation of " +key, 0)
    my_var_list = []

    for row in range(4,num_rows):

        var_name = reg.iloc[row, 0]
        label = reg.iloc[row, 2]
        if key == "AT_ulykker" or key == "AT_erhvervssygdomme": #fixes a  weird index bug with AT_ulykker
            var_name = reg.iloc[row, 1]
            label = reg.iloc[row, 3]

        if var_name == "TIMES":
            continue
        if pd.isnull(var_name) and pd.isnull(label):
            continue
        if pd.isnull(label) == True:
            label = "No label"

        var_name = str(var_name)
        label = str(label)

        if my_var_list.__contains__(var_name): #continue if the variable has been present before
            continue
        if var_name != "nan":
            my_var_list.append(var_name)

        paragraph = document.add_paragraph()
        paragraph.add_run(var_name + ":").bold = True
        paragraph.add_run("    " + label)

    document.save(save_path + key + ".docx")
#Register_list = dict_7217['Oversigt']["Unnamed: 1"]


