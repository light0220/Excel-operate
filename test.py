from excel_operate import ExcelOperate

excel = ExcelOperate('tests/示例.xlsx')
excel.ws['C15'].value = '测试通过！！'
for row in excel.ws.iter_rows(min_row=3,max_row=5,min_col=2,max_col=5):
    for cell in row:
        print(cell.value)
excel.insert_rows(5,3,35)
excel.insert_cols(5,3,35)
excel.delete_rows(5,5)
excel.delete_cols(3,4)
excel.save('tests/示例 - 修改.xlsx')