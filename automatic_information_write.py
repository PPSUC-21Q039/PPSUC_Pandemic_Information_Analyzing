import xlrd
import xlwt
import time

information_workbook = 'information.xls'
input_workbook = 'input.xls'
output_workbook = 'output.xls'
current_time = time.asctime()
input_workbook_col = 0 # Could be changed frequently

def process_read (workbook_name, index_num, col_num):
    raw = xlrd.open_workbook (workbook_name)
    information_sheet = raw.sheet_by_index (index_num)
    information_sheet_col = information_sheet.col_values (col_num)
    return information_sheet_col

# 提前读取出来，避免无谓的重复读取，可将CPU time 从 5s 降到 0.1s
to_be_compared = process_read ('input.xls', 0, input_workbook_col)

# 提前打开Workbook，避免多次创建文件，导致最后只被覆盖来只剩下数二的
wk = xlwt.Workbook()
sheet = wk.add_sheet ("result", cell_overwrite_ok=False)

# 定义Style（指表头说明）
style_bold = xlwt.easyxf ('font: bold on')
sheet.write_merge (0, 0, 0, 3, current_time)

def compare_data (workbook_name, index_num, col_num, input_workbook_name, input_index_num, input_col_num):
    result = [temp for temp in process_read (workbook_name, index_num, col_num) if temp in str(to_be_compared).strip()]
    return result

def process_write (result_data, workbook_name, index_num, col_num, title_name):
    sheet.write (1, col_num, title_name, style_bold)
    #i = 1
    i = 2 # 保留第一行来放时间
    for temp in result_data:
        sheet.write (i, col_num, temp)
        i = i+1
    wk.save (workbook_name)

def main():
    result_shi_1 = compare_data (information_workbook, 0, 0, input_workbook, 0, input_workbook_col)
    result_shi_2 = compare_data (information_workbook, 1, 0, input_workbook, 0, input_workbook_col)
    result_shu_1 = compare_data (information_workbook, 2, 0, input_workbook, 0, input_workbook_col)
    result_shu_2 = compare_data (information_workbook, 3, 0, input_workbook, 0, input_workbook_col)

    process_write (result_shi_1, output_workbook, 0, 0, "视听一区")
    process_write (result_shi_2, output_workbook, 0, 1, "视听二区")
    process_write (result_shu_1, output_workbook, 0, 2, "数据一区")
    process_write (result_shu_2, output_workbook, 0, 3, "数据二区")

if __name__ == "__main__":
    main()
