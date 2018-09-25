from functools import partial

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


class Nei:
    def __init__(self, username, password):
        self.driver = webdriver.Chrome()
        self.wait = partial(WebDriverWait, self.driver)
        self.__login(password, username)

    def __login(self, password, username):
        self.get('main.php')
        self._input('username', username)
        self._input('password', password)
        self._click_button('submit')
        self.wait(60).until(EC.title_is('研究部数据管理系统'))
        print('登陆成功！')

    def get(self, url):
        base = 'http://192.168.108.16/realty/admin'
        self.driver.get(f'{base}/{url}')

    def _input(self, ele, text):
        self.driver.find_element_by_name(ele).send_keys(text)

    def _click_button(self, name):
        self.driver.find_element_by_name(name).click()

    def _select(self, ele, value):
        """下拉列表
        :param
            ele:表单控件的name
            value:表单的 value 或visible_text
        """
        s = Select(self.driver.find_element_by_name(ele))
        try:
            s.select_by_visible_text(value)
        except NoSuchElementException:
            s.select_by_value(value)

    def _multiselect(self, ele, value_list):
        """多选
        :param
            ele:表单控件的name
            value_list:表单的 value 或visible_text 组成的列表
        """
        s = Select(self.driver.find_element_by_name(ele))
        s.deselect_all()

        for value in value_list:
            try:
                s.select_by_visible_text(value)
            except NoSuchElementException:
                s.select_by_value(value)


class NeiAPI(Nei):
    def gongxiao(self, by, **kwargs):
        """供销查询
        :param
            by: 'week', 'month', 'year'
            start, end: 2017年第1周 => '201701'; 2017年1月 => '2017-01-00'
            block: 板块 default:'全市'
            stat: 输出方式
            usg: 物业类型
            item：输出项
            add: 累计,逐周,逐月
        """
        self.get(f'ol_new_block_{by}.php')
        self.wait(60).until(EC.presence_of_element_located((By.NAME, 'block')))

        # 设置开始、结束时间
        kwargs[f'{by}1'], kwargs[f'{by}2'] = kwargs['start'], kwargs['end']

        # 下拉列表的选项们
        for key in [f'{by}1', f'{by}2', 'block', 'stat', 'add']:
            if key in kwargs:
                self._select(key, kwargs[key])

        # 物业类型
        if 'usg' in kwargs:
            self._multiselect('usage[]', kwargs['usg'])

        # 输出项
        if 'item' in kwargs:
            if by == 'month':
                self._multiselect('Litem1[]', kwargs['item'])
            else:
                self._multiselect('Litem2[]', kwargs['item'])

        # 查询
        self._click_button('Submit')

        # 读取结果
        try:
            self.wait(120).until(EC.presence_of_element_located((By.TAG_NAME, 'caption')))

            bs = BeautifulSoup(self.driver.page_source, 'lxml')
            table = bs.table.find('table').prettify()
            df = pd.read_html(table, index_col=0, header=1)[0]
            print('查询成功!')
            return df
        except TimeoutException:
            print('查询失败！')


