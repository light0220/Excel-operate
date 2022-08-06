from excel_operate import ExcelOperate
from copy import copy
from openpyxl.utils import get_column_letter

class SheetCopy:
    def __init__(self,src_file_path:str,tag_file_path:str=None,sheet_name:str=None,column_adjust:float=0) -> None:
        '''
        src_file_path: 此参数为源Excel文件路径
        tag_file_path: 此参数为目标Excel文件路径，不指定的情况下默认新建工作薄
        sheet_name: 此参数为被复制的工作表名称，不指定的情况下默认为源工作表中的当前激活工作表
        column_adjust: 此参数为列宽修正系数，作用为修正列宽误差
        '''
        self.src_file = ExcelOperate(src_file_path,sheet_name)
        if tag_file_path == None:
            self.tag_file = ExcelOperate()
            self.tag_file.ws.title = self.src_file.ws.title
        else:
            self.tag_file = ExcelOperate(tag_file_path)
            if self.src_file.ws.title not in self.tag_file.wb.sheetnames:
                self.tag_file.wb.create_sheet(self.src_file.ws.title)
                self.tag_file.ws = self.tag_file.wb[self.src_file.ws.title]
            else:
                n = 1
                while True:
                    if f'{self.src_file.ws.title} ({n})' not in self.tag_file.wb.sheetnames:
                        self.tag_file.wb.create_sheet(f'{self.src_file.ws.title} ({n})')
                        self.tag_file.ws = self.tag_file.wb[f'{self.src_file.ws.title} ({n})']
                        break
                    n += 1
        self.column_adjust = column_adjust # 列宽误差修正系数，默认为0，由于openpyxl在设置列宽时与实际存在出入，因此可以给定此参数来调整误差

    # 定义工作表复制模块
    def copy_sheet(self):
        for row in self.src_file.ws:
            # 遍历源xlsx文件制定sheet中的所有单元格
            for cell in row:  # 复制数据
                self.tag_file.ws[cell.coordinate].value = cell.value
                if cell.has_style:  # 复制样式
                    self.tag_file.ws[cell.coordinate].font = copy(cell.font)
                    self.tag_file.ws[cell.coordinate].border = copy(cell.border)
                    self.tag_file.ws[cell.coordinate].fill = copy(cell.fill)
                    self.tag_file.ws[cell.coordinate].number_format = copy(cell.number_format)
                    self.tag_file.ws[cell.coordinate].protection = copy(cell.protection)
                    self.tag_file.ws[cell.coordinate].alignment = copy(cell.alignment)

        wm = list(zip(self.src_file.ws.merged_cells))  # 开始处理合并单元格
        if len(wm) > 0:  # 检测源xlsx中合并的单元格
            for i in range(0, len(wm)):
                cell2 = (str(wm[i]).replace("(<MergedCellRange ", "").replace(">,)", ""))  # 获取合并单元格的范围
                self.tag_file.ws.merge_cells(cell2)  # 合并单元格
        # 开始处理行高列宽
        for i in range(1, self.src_file.ws.max_row + 1):
            self.tag_file.ws.row_dimensions[i].height = self.src_file.ws.row_dimensions[i].height
        for i in range(1, self.src_file.ws.max_column + 1):
            self.tag_file.ws.column_dimensions[get_column_letter(i)].width = self.src_file.ws.column_dimensions[get_column_letter(i)].width + self.column_adjust  # 修正列宽误差
        
        return self.tag_file


# 调试
if __name__ == '__main__':
    src_file_path = r'D:\codes\Python Projects\excel_operate_root\tests\示例.xlsx'
    copyer = SheetCopy(src_file_path,r'D:\Desktop\1111.xlsx')
    tag_file = copyer.copy_sheet()
    tag_file.save(r'D:\Desktop\1111.xlsx')