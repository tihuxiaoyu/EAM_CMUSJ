import datetime
import pprint as pp

import openpyxl
import asset_database
import order_numbers
from asset import Asset
from settings import Settings

class Maintain():
    """固定资产维修的类"""

    def __init__(self):
        """初始化"""
        self.settings = Settings(order_numbers)
        self.asset_info = {}
        self.order_nummber = self.settings.order_number
        self.maintain_order = {}

    def main(self):
        while True:
            anwser = input('是否创建一个新的维修工单？Y/N\n')
            if anwser == 'Y' or anwser == 'y':
                # 更新维修单号
                self.settings.renew_order_number()
                self.order_nummber = self.settings.order_number

                # 确定固定资产信息
                while True:
                    try:
                        asset_number = input('请输入设备资产编号：\n')
                        self.asset_info = asset_database.all_data[asset_number]
                        print(self.asset_info)
                        self.asset_info.setdefault(self.order_nummber, {})
                        break
                    except KeyError:
                        print('没有找到该资产！')

                # 输入表格所用信息
                self._input_repairs_info()
                self._input_parts_info()
                self._input_maintain_agency_info()

                # 创建一个固定资产实例
                asset = Asset(self.asset_info, self.order_nummber)
                # 生成维修申请表
                asset.creat_form()
            else:
                break
            
            

    def _input_repairs_info(self):
        """输入报修相关信息"""
        # 故障内容
        fault_describe = input('请输入故障描述：\n')

        while True:
            try:
                repair_time = input('''请输入报修时间（默认为当前时间）\
（格式：YYYY-MM-DD）：\n''')
                if len(repair_time) == 0:
                    repair_time = datetime.date.today()
                else:
                    repair_time = datetime.datetime.strptime(repair_time, \
                        '%Y-%m-%d')
                break
            except ValueError:
                print('时间格式错误！')

        repair_man = input('请输入报修联系人：\n')
        repair_man_tel = input('请输入报修人联系方式：\n')
        # 存储报修信息的字典repairs
        repairs = {'故障描述': fault_describe, 
                 '报修时间': repair_time,
                 '报修人': repair_man,
                 '报修人联系方式': repair_man_tel,}
        self.asset_info[self.order_nummber].setdefault('报修内容', repairs)
        print(self.asset_info)
        
     # 配件信息
    def _input_parts_info(self):
        """输入配件信息"""
        self.asset_info[self.order_nummber].setdefault('parts_info', {})
        num = 0
        active = True
        while active:
            anwser = input('是否有新的配件信息？（Y/N）')
            if anwser.upper() == 'Y':
                num += 1
                parts_info_n = '配件信息' + str(num)
                part_name = input('请输入配件名称：\n')
                part_type = input('请输入配件型号：\n')
                while True:
                    try:
                        part_price = float(input('请输入配件价格：\n'))
                        break
                    except ValueError:
                        print('格式错误！请输入数字！')
                while True:
                    try:
                        part_number = int(input('请输入配件数量：\n'))
                        break
                    except ValueError:
                        print('格式错误！请输入数字！')
                # 存储配件信息的字典parts
                parts = {'配件名称': part_name,
                        '配件型号': part_type,
                        '配件价格': part_price,
                        '配件数量': part_number,}
                self.asset_info[self.order_nummber]['parts_info'].\
                setdefault(parts_info_n, parts)
            else:
                active = False

        print(self.asset_info)

    def _input_maintain_agency_info(self):
        """输入维修代理商信息"""
        agency_name = input('请输入维修代理商名称：\n')
        agency_tel = input('请输入维修代理商联系方式（联系人名称，电话）：\n')
        # 存储代理商信息的字典maintain_agency
        maintain_agency = {'维修代理商': agency_name,
                            '联系方式': agency_tel,}
        self.asset_info[self.order_nummber].setdefault('维修代理商', 
            maintain_agency)

        print(self.asset_info)

    def _input_maintain_info(self):
        """输入维修信息"""
        maintain_details = input('请输入维修内容：\n')
        while True:
            try:
                end_time = input('请输入维修结束时间（YYYY-MM-DD）：\n')
                end_time =datetime.datetime.strptime(end_time)
                break
            except ValueError:
                print('时间格式错误！')
        maintain_engineer = input('请输入维修工程师：\n')
        engineer_tel = input('请输入维修工程师联系方式：\n')
        # 存储维修信息的类maintain_info
        maintain_info = {'维修内容': maintain_details,
                        '维修结束时间': end_time,
                        '维修工程师': maintain_engineer,
                        '工程师联系方式': engineer_tel,}
        self.asset_info[self.order_nummber].setdefault('维修信息', 
            maintain_info)

        print(self.asset_info)

if __name__ == '__main__':
    """创建一个维修实例并执行主程序"""
    maintain = Maintain()
    maintain.main()