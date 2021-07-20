import openpyxl

class RepairRequisitionForm:
    """维修及配件采购申请单的类"""

    def _init_(self, asset_info, order_number):
        self.form_name = '医疗设备维修及配件采购申请新.xlsx'
        self.sheet_name = 'Sheet1'
        self.asset_info = asset_info
        self.order_number = order_number

    def creat_form(self):
        """创建维修及配件申请单"""
        wb = openpyxl.load_workbook(self.form_name)
        sheet = wb[self.sheet_name]

        sheet['B2'] = self.asset_info['使用部门']
        sheet['F2'] = self.asset_info[self.order_number]['故障内容']['报修时间']
        sheet['B3'] = self.asset_info['固定资产名称']
        sheet['F3'] = self.asset_info['固定资产编号']

        form_name_n = '''医疗设备维修及配件采购申请新_'
        f'{self.asset_info['固定资产编号']}_'
        f'{self.asset_info[self.order_number]}.xlsx'''
        ws.save(form_name_n)

        print('维修申请单已经创建！')