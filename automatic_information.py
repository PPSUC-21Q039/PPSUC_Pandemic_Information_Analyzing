import xlrd

# 打开视听数据中队的花名册
raw = xlrd.open_workbook('information.xls')

# 视听一区
information_sheet_1 = raw.sheet_by_index(0)
information_sheet_1_col_1 = information_sheet_1.col_values(0) # 读取花名册的第一列
# information_sheet_1_col_1 = str(information_sheet_1_col_1).strip()

# 视听二区
information_sheet_2 = raw.sheet_by_index(1)
information_sheet_2_col_1 = information_sheet_2.col_values(0) # 读取花名册的第一列
# information_sheet_2_col_1 = str(information_sheet_2_col_1).strip()

# 数据一区
information_sheet_3 = raw.sheet_by_index(2)
information_sheet_3_col_1 = information_sheet_3.col_values(0) # 读取花名册的第一列
# information_sheet_3_col_1 = str(information_sheet_3_col_1).strip()

# 数据二区
information_sheet_4 = raw.sheet_by_index(3)
information_sheet_4_col_1 = information_sheet_4.col_values(0) # 读取花名册的第一列
# information_sheet_4_col_1 = str(information_sheet_4_col_1).strip()


### 读取疫情信息未打卡用户
to_be_compared = xlrd.open_workbook('input.xls')
input_sheet_1 = to_be_compared.sheet_by_index(0)
to_be_compared_col_1 = input_sheet_1.col_values(0) # 未打卡名单（第一列）

# 比对
result_shi_1 = [x for x in information_sheet_1_col_1 if x in str(to_be_compared_col_1).strip()]
result_shi_2 = [x for x in information_sheet_2_col_1 if x in str(to_be_compared_col_1).strip()]
result_shu_1 = [x for x in information_sheet_3_col_1 if x in str(to_be_compared_col_1).strip()]
result_shu_2 = [x for x in information_sheet_4_col_1 if x in str(to_be_compared_col_1).strip()]

print ("视听一区：", result_shi_1)
print ("\n视听二区：", result_shi_2)
print ("\n数据一区：", result_shu_1)
print ("\n数据二区：", result_shu_2)
