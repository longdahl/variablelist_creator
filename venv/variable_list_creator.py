import os
import xlsxwriter


arbejde = 1
if arbejde == 1:
    save_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\variable_lists\\'
    input_list = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\var_list_input"
    load_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018_newname.xlsx'
    dst_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\"
else:
    input_list = "C:\\Users\\Mikkel\\Desktop\\arbejde\\Project database\\datamanager\\script\\var_list_input"
    save_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\variable_lists\\"
    dst_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\"


def var_name_fixer(line):
    var_name = line.split(" ")[1]
    condition_list = ["_1","_2","_3","_4","_5","_6","_7","_8","_9"]
    for condition in condition_list:
        if var_name.endswith(condition):
            var_name = var_name[:-2]
            break
    return_name = line.split(" ")[0] + " " + var_name + " " + line.split(" ")[2]
    return return_name

def create_var_list(filename,reg_list,list_load_path,save_path):
    #Todo problemer med FIRM
    workbook = xlsxwriter.Workbook(save_path + filename + '.xlsx')
    for reg in reg_list:
        try:
            reg = reg.upper()
            reg = reg.rstrip("\n\r")
            file_dst = open(list_load_path + "dst_out.txt", "r")
            file_4276 = open(list_load_path + "var4276.txt", "r")
            temp_dst = []

            for line in file_dst:

                if line.startswith(reg):
                    line = line.upper()
                    line = var_name_fixer(line)
                    temp_dst.append(line)

            temp_4276 = []

            for line in file_4276:

                if line.startswith(reg):
                    line = line.upper()
                    line = var_name_fixer(line)
                    temp_4276.append(line)


            file_match = [value[:-1] for value in temp_dst if value in temp_4276]
            file_match = sorted(file_match)

            if reg == "DODSAASG":
                for line in file_match:
                    print(line)

            worksheet = workbook.add_worksheet(reg)
            dst_format = workbook.add_format()
            dst_format.set_bg_color('red')
            ivoe_format = workbook.add_format()
            ivoe_format.set_bg_color('green')
            prev_line = ""
            row = 1
            lowest = 10000

            for line in temp_dst:
                if lowest > int(line.split(" ")[2]):
                    lowest = int(line.split(" ")[2])

            worksheet.write(0, 0, reg)
            worksheet.write(1, 0, "Variable name")
            worksheet.write(0, 1, "IVØ leverance", ivoe_format)
            worksheet.write(0, 2, "DST leverance", dst_format)
            var_list = []

            for line in temp_dst:
                if prev_line == line.split(" ")[1]:
                    col = int(line.split(" ")[2]) - lowest + 1
                    worksheet.write(row, col, line.split(" ")[2], dst_format)
                else:
                    col = int(line.split(" ")[2]) - lowest + 1
                    row = row + 1
                    worksheet.write(row, 0, line.split(" ")[1])
                    var_list.append(line.split(" ")[1])
                    worksheet.write(row, col, line.split(" ")[2], dst_format)
                prev_line = line.split(" ")[1]

            for line in file_match:
                row = var_list.index(line.split(" ")[1]) + 2
                col = int(line.split(" ")[2]) - lowest + 1
                worksheet.write(row, col, line.split(" ")[2], ivoe_format)

        except:
            print("there was an error with register: " + reg)
    workbook.close()

if __name__ == '__main__':

    for filename in os.listdir(input_list):
        reg_list = []
        for line in open(input_list + "\\" + filename):
            reg_list.append(line)
        create_var_list(filename,reg_list,list_load_path,save_path)

