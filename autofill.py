from query import NeiAPI
from tushare import MailMerge
import pandas as pd

# 输出项目映射
ITEM_MAPPING = {'price': '已售均价', 'price_mom': '均价环比', 'price_yoy': '均价同比', 'sale_rate': '上市面积占比', 'sold_rate': '已售面积占比',
                'rolling_rate': '滚动一年供销比'}
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
        if 'mom' in col or 'yoy' in col:
            df[col] = df[col].apply(rate)

    return df


def ershou_adjust(df):
    for col in ['sold', 'money']:
        df[col] = wan(df[col])

    for col in ['set', 'price']:
        df[col] = df[col].astype('int')

    for col in df:
        if 'mom' in col or 'yoy' in col:
            df[col] = df[col].apply(rate)

    return df


def ershou_cum_adjust(df):
    for col in ['S_cumsold', 'SR_cumsold']:
        df[col] = df[col].apply(lambda x: f'{float(x):.2f}')

    for col in ['S_cummoney', 'SR_cummoney']:
        df[col] = wan(df[col])

    for col in ['S_cumset', 'SR_cumset', 'S_cumprice', 'SR_cumprice']:
        df[col] = df[col].astype('int')

    for col in df:
        if 'mom' in col or 'yoy' in col:
            df[col] = df[col].apply(rate)

    return df


def gen_item(*args):
    return sum([[x, f'{x}_mom', f'{x}_yoy'] for x in args], [])


def rolling_rate():
    df = rpt.nei.gongxiao(by='month', start=rpt.now, end=rpt.now, stat='按板块/片区', usg=SPZZ, item=['滚动一年供销比'], add='逐月')
    rpt.data.update({'rolling_rate': round(df.loc['溧水'][0] / 100, 2)})


def stock_speed():
    df = rpt.nei.gongxiao(by='month', start=f'2017年{rpt.data["month"]+1:0>2d}月', end=rpt.now, stat='按板块/片区', usg=SPZZ,
                          item=['已售面积', '已售套数'])
    sold, sold_set = df.loc['溧水']

    df = pd.read_excel('data.xlsx', '库存')
    df['街道'] = df['街道'].str.strip()
    df.set_index('街道', inplace=True)
    stock, stock_set = df.loc['合计']

    speed = round(stock / sold * 12, 2)
    speed_set = round(stock_set / sold_set * 12, 2)

    rpt.data.update({
        'stock': f'{stock/1e4:.2f}',
        'speed': speed,
        'stock_set': int(stock_set),
        'speed_set': speed_set
    })

    df.reindex(['永阳街道', '开发区', '石湫镇', '洪蓝镇', '白马镇', '和凤镇', '东屏镇', '晶桥镇', '合计'], inplace=True)
    print((df['stock'] / 1e4).round(2))


class Report:

    def __init__(self):
        self.data = {'month': int(input('请输入当前月份：'))}
        self.now = f'2018年{self.data["month"]:0>2d}月'
        self.nei = NeiAPI('彭子乔', 'password02!')
        self.doc = MailMerge('template.docx')

    def query(self, item, usage='', cum=False, block='溧水', stat='按板块/片区'):
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

    def ershou(self):
        # 当月数据 GIS二手房当月交易情行，分别拉二手房与手手住宅
        df = pd.read_excel('data.xlsx', '二手房当月', index_col=0)
        df = ershou_adjust(df)
        for i, usage in enumerate(['S', 'SR']):
            row = df.iloc[i]
            for col in row.index:
                self.data.update({f'{usage}_{col}': row[col]})

        # 当年数据 内网-存量房2012-月报表-累计同比
        df = pd.read_excel('data.xlsx', '二手房当年')
        df = ershou_cum_adjust(df)
        self.data.update(df.iloc[0].to_dict())


if __name__ == '__main__':
    rpt = Report()

    # 当月商品房
    rpt.query(item=gen_item('sale', 'sold', 'money'), usage='A_')

    # 当月商品住宅
    rpt.query(item=gen_item('sale', 'sale_set', 'sold', 'sold_set', 'price', 'money'))
    rpt.query(item=['sale_rate', 'sold_rate'], block='全市')
    for usage in ['R_', 'V_']:
        rpt.query(item=['price'], usage=usage)

    # 当月商办物业
    rpt.query(item=gen_item('sale', 'sold', 'price', 'money'), usage='CO_')

    # 当年商品房
    rpt.query(item=['sale', 'sale_yoy', 'sold', 'sold_yoy', 'money', 'money_yoy'], usage='A_', cum=True)

    # 当年商品住宅
    rpt.query(item=['sale', 'sale_rate', 'sale_yoy', 'sale_set', 'sale_set_yoy',
                    'sold', 'sold_rate', 'sold_yoy', 'sold_set', 'sold_set_yoy',
                    'price', 'price_yoy', 'money', 'money_yoy'], block='全市', cum=True)
    for usage in ['R_', 'V_']:
        rpt.query(item=['price'], usage=usage, cum=True)
    rolling_rate()

    # 当年商办物业
    rpt.query(item=['sale', 'sale_yoy', 'sold', 'sold_yoy',
                    'price', 'price_yoy', 'money', 'money_yoy'], usage='CO_', cum=True)

    # 二手房
    rpt.ershou()

    # 库存
    stock_speed()

    # 商、办量价
    for usage in ['C_', 'O_']:
        rpt.query(item=gen_item('sale', 'sold', 'price'), usage=usage)

    # 输出到模板
    rpt.doc.merge(**rpt.data)
    rpt.doc.write('output.docx')
