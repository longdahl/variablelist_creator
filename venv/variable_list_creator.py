import pandas as pd
import math
import os
import xlsxwriter
import numpy as np
import pickle
import tkinter

arbejde = 0
if arbejde == 1:
    save_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\variable_lists\\'
    load_path = 'C:\\Users\\mikkel-bj\\Desktop\\datamanager\\sharepoint_backup\\704276_Opdateringsoversigt30082018_newname.xlsx'
    dst_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\mikkel-bj\\Desktop\\datamanager\\script\\"
else:
    save_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\variable_lists\\"
    dst_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\from_dst\\"
    list_load_path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\"


def create_var_list(reg,list_load_path,save_path):
    #Todo problemer med FIRM
    reg = reg.upper()
    file_dst = open(list_load_path + "dst_out.txt", "r")
    file_4276 = open(list_load_path + "out_allsheets.txt", "r")

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
    workbook.close()


if __name__ == '__main__':

    reg = "aefb"
    create_var_list(reg,list_load_path,save_path)

