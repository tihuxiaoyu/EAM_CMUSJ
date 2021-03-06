import datetime
import pprint as pp
import json

# import order_numbers

class Settings:
    """固定资产管理系统的参数设置"""
    def __init__(self):
        """初始化系统的设置"""
        # self.order_numbers = order_numbers.all_data
        with open('order_numbers.json') as f:
            self.order_numbers = json.load(f)
        self._order_number = self.order_numbers[-1]

    def renew_order_number(self):
        """更新维修工单号order_number"""
        date = str(datetime.date.today().strftime('%Y%m%d'))
    
        if date == self._order_number[:8]:
            order_number = int(self._order_number)
            order_number += 1
            self._order_number = str(order_number)
        else:
            self._order_number = f'{date}001'

        print(f'本次维修工单号为：{self._order_number}')
        self.order_numbers.append(self._order_number)
        with open('order_numbers.json', 'w') as f:
            json.dump(self.order_numbers, f, indent=4)
        #     f.write('all_data = ' + pp.pformat(self.order_numbers))
        print('新的工单号已存储！')

    # 访问器 - getter方法
    @property
    def order_number(self):
        """读取维修工单号order_number"""
        return self._order_number

if __name__ == '__main__':
    settings = Settings()
    settings.renew_order_number()