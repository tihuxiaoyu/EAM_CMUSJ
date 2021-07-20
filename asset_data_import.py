
# import sys type = sys.getfilesystemencoding()

import openpyxl as op, pprint as pp

print('Opening workbook...')

wb = op.load_workbook('本溪南湖20210120.xlsx')
sheet = wb['Sheet1']

# 字典asset_data存储固定资产数据
asset_data = {}
print('Reading rows...')
for row in range(2, sheet.max_row + 1):
    # 表格中每行数据读取出来
    asset_number = sheet['A' + str(row)].value
    contract_number = sheet['B' + str(row)].value
    sn = sheet['C' + str(row)].value
    start_data = sheet['D' + str(row)].value
    asset_name = sheet['E' + str(row)].value
    asset_type = sheet['F' + str(row)].value
    asset_values = sheet['G' + str(row)].value
    number = sheet['H' + str(row)].value
    nationality = sheet['I' + str(row)].value
    factory = sheet['J' + str(row)].value
    agency = sheet['K' + str(row)].value
    user_department = sheet['L' + str(row)].value
    stor_place = sheet['M' + str(row)].value
    st_stor_place = sheet['N' + str(row)].value
    place_number = sheet['O' + str(row)].value
    engineer = sheet['P' + str(row)].value
    officer = sheet['Q' + str(row)].value
    ser_years = sheet['R' + str(row)].value

    # 设置字典中的默认值
    asset_data.setdefault(asset_number, {})
    asset_data[asset_number] = {'固定资产编号': asset_number,'合同号': contract_number, 
        '序列号': sn, '开始使用日期': start_data, '固定资产名称': asset_name,
        '标准型号': asset_type, '原值': asset_values, '数量': number,
        '国别': nationality, '标准生产厂家': factory, '标准经销商': agency,
        '使用部门': user_department, '存放地点': stor_place,
        '标准存放地点': st_stor_place, '实际地点编码': place_number,
        '负责工程师': engineer, '负责人': officer, '使用年限': ser_years, 
        '设备维修': {}}


# asset_data.decode('UTF-8').encode(type)
print("Writing results...")
result_file = open('asset_database.py', 'w')
result_file.write('#-*-coding:gbk-*-\n''all_data = ' + pp.pformat(asset_data))
print('Done')
print(asset_data)