import datetime

import openpyxl
import asset_database
from create_repair_requisition_form import RepairRequisitionForm

asset_info = {}
maintain_order = {}

anwser = input('是否创建一个新的维修工单？Y/N\n')
if anwser == 'Y' or 'y':
    asset_number = input('请输入设备资产编号：\n')
    asset_info = asset_data[asset_number]
    print(asset_info)
