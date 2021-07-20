import datetime
import pprint as pp

import order_numbers

class Settings:
    """固定资产管理系统的参数设置"""
    def __init__(self, order_numbers):
        """初始化系统的设置"""
        self.order_numbers = order_numbers.all_data
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

        print(self._order_number)
        self.order_numbers.append(self._order_number)
        file = open('order_numbers.py', 'w')
        file.write('all_data = ' + pp.pformat(self.order_numbers))

    # 访问器 - getter方法
    @property
    def order_number(self):
        """读取维修工单号order_number"""
        return self._order_number

if __name__ == '__main__':
    settings = Settings(order_numbers)
    settings.renew_order_number()