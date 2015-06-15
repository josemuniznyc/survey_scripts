import os
import xlrd

part_1_dir = "/Users/jose.muniz/Dropbox/Documents/SPS/evals/s15/online_courses/sm/part_1"
part_2_dir = "/Users/jose.muniz/Dropbox/Documents/SPS/evals/s15/online_courses/sm/part_2"
excel_dir = "/Users/jose.muniz/Dropbox/Documents/SPS/evals/s15/online_courses/sm/excel_files"


def rename_excel_file(old, new):
    new_filename = os.path.join(excel_dir, new + ".xls")
    os.rename(old, new_filename)


def parse_excel_folders(directory):
    os.chdir(directory)
    excel_file = os.listdir()
    if excel_file:
        excel_file = excel_file[0]
        wb = xlrd.open_workbook(excel_file)
        ws = wb.sheet_by_index(0)
        parse_text = ws.row_values(0)
        parse_text = parse_text[0]
        parse_text = parse_text.replace(" ", "_")
        rename_excel_file(excel_file, parse_text)
    else:
        os.remove(directory)
        pass


def process_dir(directory):
    os.chdir(directory)
    folders = [x for x in os.listdir(".") if x[-4:] != ".zip"]
    folders = [os.path.abspath(x) for x in folders]
    [parse_excel_folders(x) for x in folders]

# Handle part 1
process_dir(part_1_dir)

# Handle part 2
process_dir(part_2_dir)
