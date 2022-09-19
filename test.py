from excel_operate import ExcelOperate
from excel_operate import SheetCopy
from excel_operate import SheetComparison
from excel_operate import is_insert

excel = ExcelOperate('tests/示例.xlsx')
# excel.ws['C15'].value = '测试通过！！'
for row in excel.ws.iter_rows(min_row=3,max_row=5,min_col=2,max_col=5):
    for cell in row:
        print(cell.value)
excel.insert_rows(5,3,35)
excel.insert_cols(5,3,35)
excel.delete_rows(5,5)
excel.delete_cols(3,4)
excel.save('D:/desktop/1111.xlsx')

stcper = SheetCopy('tests/示例.xlsx','D:/desktop/1111.xlsx')
tag_file = stcper.copy_sheet()
tag_file.save('D:/desktop/1111.xlsx')

src_excel = ExcelOperate()
# cmp_excel = SheetComparison()

l1 = [1, 3, 5, 6, 7]
l2 = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
ins_info = is_insert(l1, l2)
print(ins_info)