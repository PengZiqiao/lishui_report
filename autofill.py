from query import NeiAPI
from tushare import MailMerge

# 输出项目映射
ITEM_MAPPING = {'price': '已售均价', 'price_mom': '均价环比', 'price_yoy': '均价同比', 'sale_rate': '上市面积占比', 'sold_rate': '成交面积占比'}
for k, v in {'sale': '上市面积', 'sale_set': '上市套数', 'sold': '已售面积', 'sold_set': '已售套数', 'money': '已售金额'}.items():
    ITEM_MAPPING.update({k: v, f'{k}_mom': f'{v}环比', f'{k}_yoy': f'{v}同比'})

# 物业类型常量
ZHUZHAI = ['住宅']
BIESHU = ['联排别墅', '双拼别墅', '独立别墅', '叠加别墅']
SPZZ = [*ZHUZHAI, *BIESHU]
SHANGYE = ['商业']
BANGONG = ['办公']
SHANGBAN = [*SHANGYE, *BANGONG]
SPF = [*SPZZ, *SHANGBAN, '车库', '其它']

USAGE_MAPPING = {'R_': ZHUZHAI, 'V_': BIESHU, '': SPZZ, 'C_': SHANGYE, 'O_': BANGONG, 'CO_': SHANGBAN, 'A_': SPF}


def wan(series):
    return series.apply(lambda x: f'{float(x)/1e4:.2f}')


def rate(x):
    return f'增长{x:.1f}' if x >= 0 else f'下降{-x:.1f}'


def adjust(df):
    # 面积换算成万㎡
    for col in ['sale', 'sold']:
        if col in df:
            df[col] = wan(df[col])

    # 金额换算成亿元
    if 'money' in df:
        df['money'] = df['money'].apply(lambda x: f'{float(x)/1e8:.2f}')

    # 套数、单价为整数
    if col in ['sale_set', 'sold_set', 'price']:
        if col in df:
            df[col] = df[col].astype('int')

    # 占比保留一位小数
    for col in ['sale_rate', 'sold_rate']:
        if col in df:
            df[col] = round(df[col], 1)

    # 增长率
    for col in df:
        if ('mom' in col) | ('yoy' in col):
            df[col] = df[col].apply(rate)

    return df


class Report:

    def __init__(self):
        self.data = dict(
            month=int(input('请输入当前月份：'))
        )
        self.now = f'2018年{self.data["month"]:0>2d}月'
        self.nei = NeiAPI('彭子乔', 'password02!')
        self.doc = MailMerge('template.docx')

    def query(self, item, usage, cum=False, block='溧水', stat='按板块/片区'):
        # 选项
        options = {
            'start': '2018年01月' if cum else self.now,
            'end': self.now,
            'block': block,
            'stat': stat,
            'usg': USAGE_MAPPING[usage],
            'item': [ITEM_MAPPING[each] for each in item]
        }

        # 查询
        df = self.nei.gongxiao('month', **options)
        df.columns = item

        # 数据处理
        df = adjust(df)

        # 导出到字典
        for each in item:
            self.data.update({
                f'{usage}{"cum" if cum else ""}{each}': df.at['合计' if block == '溧水' else '溧水', each]
            })


if __name__ == '__main__':
    pass
