import openpyxl

class Asset:
    """固定资产的类"""

    def __init__(self, asset_info, order_number):
        self.form_name = '医疗设备维修及配件采购申请新.xlsx'
        self.sheet_name = 'Sheet1'
        self.asset_info = asset_info
        self.order_number = order_number

    def creat_form(self):
        """创建维修及配件申请单"""
        wb = openpyxl.load_workbook(self.form_name)
        sheet = wb[self.sheet_name]

        sheet['B2'] = self.asset_info['使用部门']
        sheet['F2'] = self.asset_info[self.order_number]['报修内容']['报修时间']
        sheet['B3'] = self.asset_info['固定资产名称']
        sheet['F3'] = self.asset_info['固定资产编号']
        sheet['B4'] = self.asset_info['原值']
        sheet['F4'] = self.asset_info['标准生产厂家']
        sheet['B5'] = self.asset_info['标准型号']
        sheet['F5'] = self.asset_info['标准经销商']
        sheet['B6'] = self.asset_info['序列号']
        sheet['B7'] = self.asset_info[self.order_number]['报修内容']['故障描述']

        parts_info = self.asset_info[self.order_number]['parts_info']
        row_n = 10
        for key in parts_info:
            sheet.cell(row=row_n, column=1).value = parts_info[key]['配件名称']
            sheet.cell(row=row_n, column=2).value = parts_info[key]['配件型号']
            sheet.cell(row=row_n, column=4).value = parts_info[key]['配件价格']
            sheet.cell(row=row_n, column=5).value = parts_info[key]['配件数量']
            sheet.cell(row=row_n, column=6).value = \
sheet.cell(row=row_n, column=4).value * sheet.cell(row=row_n, column=5).value
            row_n += 1

        sheet['F15'] = '=SUM(F10:F14)'

        sheet['B17'] = self.asset_info[self.order_number]['维修代理商']['维修代理商']
        sheet['F17'] = self.asset_info[self.order_number]['维修代理商']['联系方式']



        form_name_n = f"{self.asset_info['固定资产编号']}_{self.order_number}.xlsx"
        form_name_n = f'医疗设备维修及配件采购申请新_{form_name_n}'
        wb.save(form_name_n)
        print('维修申请单已经创建！')