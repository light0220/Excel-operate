# -*- encoding: utf-8 -*-
'''
@File    :   sheet_comparison.py
@Author  :   北极星光 
@Contact :   light22@126.com
'''

from excel_operate import ExcelOperate
from sheet_copy import SheetCopy
from list_operate import *
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Side, Border, Font


class SheetComparison:
    def __init__(self, src_excel, cmp_excel, report_path: str = None) -> None:
        '''===================================\n
        传入两个Excel对象，进行比较。生成报告
        src_excel: 原Excel对象，可传入ExcelOperate类对象
        cmp_excel: 待比较的Excel对象，可传入ExcelOperate类对象
        report_path: 报告结果保存路径
        '''
        self.src_excel = src_excel
        self.cmp_excel = cmp_excel
        self.report_path = report_path

    def set_title_row(self, src_title_row: int, cmp_title_row: int):
        '''===================================\n
        此方法可以为原工作表及待比较工作表设置表头行
        src_title_row: 原Excel工作表的表头行，可传入整数类型
        cmp_title_row: 待比较Excel工作表的表头行，可传入整数类型
        '''
        self.src_title_row = src_title_row
        self.cmp_title_row = cmp_title_row

    def set_key_col(self, src_key_col: int, cmp_key_col: int):
        '''===================================\n
        此方法可以为原工作表及待比较工作表设置关键列
        src_key_col: 原Excel工作表的关键列，可传入整数类型
        cmp_key_col: 待比较Excel工作表的关键列，可传入整数类型
        '''
        self.src_key_col = src_key_col
        self.cmp_key_col = cmp_key_col

    def font_color(self, cell, style):
        '''===================================\n
        此方法为目标单元格设置样式
        cell: 目标单元格，请传入单元格对象
        style: 目标单元格将要设置成的样式，只支持'zeng','shan'和'gai'这三个参数。
        '''
        if style == 'zeng':
            cell.font = Font(color='0000FF')
        if style == 'shan':
            cell.font = Font(color='FF0000',strike=True)
        if style == 'gai':
            cell.font = Font(color='FF00FF')

    def compare(self):
        '''===================================\n
        对比工作表：将原工作表及目标工作表的表头行和关键列设置好之后即可使用此方法对比工作表，并生成对比报告。
        '''
        # 如果目标工作表的表头行与原工作表的表头行不在同一行
        if self.src_title_row != self.cmp_title_row:
            if self.src_title_row > self.cmp_title_row:
                self.cmp_excel.insert_rows(
                    1, self.src_title_row-self.cmp_title_row)
                self.cmp_title_row = self.src_title_row
            else:
                self.src_excel.insert_rows(
                    1, self.cmp_title_row-self.src_title_row)
                self.src_title_row = self.cmp_title_row

        # 如果目标工作表的关键列与原工作表的关键列不在同一列
        if self.src_key_col != self.cmp_key_col:
            if self.src_key_col > self.cmp_key_col:
                self.cmp_excel.insert_rows(
                    1, self.src_key_col-self.cmp_key_col)
                self.cmp_key_col = self.src_key_col
            else:
                self.src_excel.insert_rows(
                    1, self.cmp_key_col-self.src_key_col)
                self.src_key_col = self.cmp_key_col

        # 如果目标工作表的表头行与原工作表的表头行不相同
        src_title_list = [
            i.value for i in self.src_excel.ws[self.src_title_row]]
        cmp_title_list = [
            i.value for i in self.cmp_excel.ws[self.cmp_title_row]]
        # print(cmp_title_list)
        if src_title_list != cmp_title_list:
            self.src_excel.insert_rows(1)
            self.src_title_row += 1
            self.cmp_excel.insert_rows(1)
            self.cmp_title_row += 1
            src_title_list = duplicate_to_only(
                src_title_list)  # 通过重命名的方式给列表去重。
            cmp_title_list = duplicate_to_only(
                cmp_title_list)  # 通过重命名的方式给列表去重。
            title_insert_info = is_insert(src_title_list, cmp_title_list)
            title_delete_info = is_insert(cmp_title_list, src_title_list)
            title_append_info = is_appand(src_title_list, cmp_title_list)
            if title_insert_info != None:
                n = 1
                for i in title_insert_info:
                    self.src_excel.insert_cols(i + n, title_insert_info[i])
                    n += title_insert_info[i]
                src_title_list
            if title_delete_info != None:
                n = 1
                for i in title_delete_info:
                    self.cmp_excel.insert_cols(i + n, title_delete_info[i])
                    n += title_delete_info[i]
            if title_append_info != None:
                self.src_excel.insert_cols(
                    self.src_excel.ws.max_column + 1, title_append_info)
            src_title_list = [
                i.value for i in self.src_excel.ws[self.src_title_row]]
            cmp_title_list = [
                i.value for i in self.cmp_excel.ws[self.cmp_title_row]]
            for idx, i, j in zip(range(len(src_title_list)), src_title_list, cmp_title_list):
                if i == None:
                    self.src_excel.ws[1][idx].value = '临'
                    self.cmp_excel.ws[1][idx].value = '增'
                if j == None:
                    self.src_excel.ws[1][idx].value = '删'
                    self.cmp_excel.ws[1][idx].value = '临'

        # 如果目标工作表的关键列与原工作表的关键列不相同
        src_key_list = [i.value for idx, i in enumerate(
            self.src_excel.ws[get_column_letter(self.src_key_col)], 1) if idx > self.src_title_row]
        cmp_key_list = [i.value for idx, i in enumerate(
            self.cmp_excel.ws[get_column_letter(self.cmp_key_col)], 1) if idx > self.cmp_title_row]
        if src_key_list != cmp_key_list:
            src_key_list = duplicate_to_only(src_key_list)  # 通过重命名的方式给列表去重。
            cmp_key_list = duplicate_to_only(cmp_key_list)  # 通过重命名的方式给列表去重。
            insert_info = is_insert(src_key_list, cmp_key_list)
            delete_info = is_insert(cmp_key_list, src_key_list)
            append_info = is_appand(src_key_list, cmp_key_list)
            if insert_info != None:
                n = self.src_title_row + 1
                for i in insert_info:
                    self.src_excel.insert_rows(i + n, insert_info[i])
                    n += insert_info[i]
            if delete_info != None:
                n = self.cmp_title_row + 1
                for i in delete_info:
                    self.cmp_excel.insert_rows(i + n, delete_info[i])
                    n += delete_info[i]
            if append_info != None:
                self.src_excel.insert_rows(
                    self.src_excel.ws.max_row + 1, append_info)
            src_key_list = [i.value for idx, i in enumerate(
                self.src_excel.ws[get_column_letter(self.src_key_col)], 1) if idx > self.src_title_row]
            cmp_key_list = [i.value for idx, i in enumerate(
                self.cmp_excel.ws[get_column_letter(self.cmp_key_col)], 1) if idx > self.cmp_title_row]
            # 处理在同一位置删除行同时又插入行可能导致出现的空行问题
            for idx, i, j in zip(range(len(src_key_list)), src_key_list, cmp_key_list):
                if i == j == None:
                    r = 0  # 通过循环自加来测出在该处出现了多少空行
                    for n in range(1, idx + 1):
                        if src_key_list[idx - n] == cmp_key_list[idx - n]:
                            break
                        r += 1
                    for n in range(1, r + 1):
                        # 删除原工作表及目标工作表中的当前空行
                        self.src_excel.delete_rows(
                            self.src_title_row + idx + n)
                        self.cmp_excel.delete_rows(
                            self.cmp_title_row + idx + n)
                        # 将原工作表前一项的下方插入空白行
                        self.src_excel.insert_rows(
                            self.src_title_row + idx + 2 - n)
                        # 将对比工作表前一顶的上方插入空白行，从而形成错行的效果
                        self.cmp_excel.insert_rows(
                            self.cmp_title_row + idx + 1 - n)

            print(src_key_list)
            print(cmp_key_list)
            self.font_color(src_excel.ws['B3'], 'shan')


# 调试
if __name__ == '__main__':
    src_excel = ExcelOperate('tests\比较示例 - 原.xlsx')
    cmp_excel = ExcelOperate('tests\比较示例 - 对比.xlsx')
    report_path = 'D:/Desktop/对比报告.xlsx'
    cmper = SheetComparison(src_excel, cmp_excel, report_path)
    cmper.set_title_row(2, 2)
    cmper.set_key_col(2, 2)
    cmper.compare()
    src_excel.save('D:/Desktop/1111src.xlsx')
    cmp_excel.save('D:/Desktop/1111cmp.xlsx')
