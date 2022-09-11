# 欢迎使用

## 一、概述

本模块为方便大家更便捷使用Python来操作Excel文件而编写，基于openpyxl进行编写，所以在使用本模块前请确保您的计算机中已安装有Python环境以及openpyxl库。

+ 作者：北极星光
+ E-mail：light22@126.com
+ Pypi官方库地址：https://pypi.org/project/excel-operate-light22
+ Github地址：https://github.com/18513233125/Excel-operate

## 二、功能介绍

+ 简化openpyxl库的导入流程，无需分别导入Workbook和load_workbook来创建和打开工作薄，只需要使用ExcelOperate类实例化对象时，指定file_path参数即打开已有的Excel文件，不指定即为新建工作薄。同时实例化打开EXcel文件时可以直接指定sheet_name参数来快速打开工作表，不指定则打开当前激活工作表。
+ 方便快捷插入和删除行（列）。解决直接使用openpyxl中工作表对象的insert_rows/insert_cols/delete_rows/delete_cols方法时，插入或删除的行（列）的位置后方行（列）的行高（列宽）以及如果存在已合并单元格的情况下，已合并的单元格信息等格式不会因插入或删除的行（列）而下（后）移或上（前）移，从而导致表格的整体格式甚至数据受到影响。通过本模块的方法插入或删除行（列）后的效果与在Excel程序中插入或删除行（列）的效果可以做到近乎完全相同。同时为insert_rows和insert_cols方法增加height和width参数，方便插入行（列）时快捷设定行高和列宽。
+ 工作表复制，解决openpyxl只能在工作薄内复制工作表，无法跨工作薄复制工作表的问题。
+ 工作表对比，可以快速对比两个工作表发生的改动，并生成对比报告。
+ 可以使用openpyxl库中的所有操作。
+ 列表去重，列表对比等列表相关的操作。
+ 其它功能敬请期待。

## 三、使用方法

+ 导入本模块方法：

  + 在需要导入本模块的代码中写入导入语句：`from excel_operate import ExcelOperate`。
+ 对象实例化方法：

  + 通过语句 `excel = ExcelOperate(file_path=None,sheet_name=None)`，获得一个名为 `excel`的对象。
    + 参数 `file_path`*为文件路径，指定后打开文件路径对应的Excel文件（相当于* `openpyxl`*中的* `load_workbook`*方法）不指定则会新建一个空白Excel文件（相当于* `openpyxl`*中的* `Workbook`*类的实例化）；*
    + *参数* `sheet_name`*为工作表名，指定后打开对应的工作表(如果打开的工作薄中不存在指定名称的工作表则以指定的工作表为名创建一个工作表)，不指定则打开当前激活的工作表（相当于工作薄对象的* `active`*方法）。*
+ 插入或删除行（列）：

  + 插入行 `excel.insert_rows(idx,amount=1,height='before')`；
  + 删除行 `excel.delete_rows(idx,amount=1)`；
  + 插入列 `excel.insert_cols(idx,amount=1,width='before')`；
  + 删除列 `excel.delete_cols(idx,amount=1)`。
    + 参数 `idx`为插入或删除的起始行（列）
    + 参数 `amount`为插入或删除的行（列）数，不指定的情况下默认插入或删除1行（列）
    + 参数 `height`和 `width`仅在插入行和列时使用，可指定为整数或者浮点数来设置插入的行高或列宽，不指定的情况下默认参数为 `'before'`，该参数可以让插入的行或列继承插入前上一行（列）的行高或列宽。另外插入行时，`height`参数也可指定为 `None`，此参数可以设置插入行的行高为自动行高。
+ 工作表复制：

  + 首先通过语句 `from excel_operate import SheetCopy`导入模块。
  + 然后通过实例化对象 `copyer = SheetCopy(src_file_path, tag_file_path=None, sheet_name=None, column_adjust=0)`来获得实例化对象 `copyer`
    + 参数 `src_file_path`为源文件路径；
    + 参数 `tag_file_path`为目标文件路径，不指定的话将新建一个工作薄；
    + 参数 `sheet_name`为被复制的工作表名称，不指定的话为源文件的当前工作表；
    + 参数 `column_adjust`为列宽修正系数，原因为 `openpyxl`中设置的列宽与实际列宽存在误差，一般为0~0.9之间。如果复制的工作表与源工作表列宽不相同的话可以修改此参数从而使复制的工作表与源工作表列宽相等。
  + 然后通过 `tag_file = copyer.copy_sheet()`来获得一个 `ExcelOperate`对象 `tag_file`。
  + 当然最后不要忘记保存目标对象 `tag_file.save(目标路径)`。
  + 示例代码如下：
    ```
    from excel_operate import SheetCopy  # 导入模块

    src_file_path = r'D:\codes\Python Projects\Excel-operate\tests\示例.xlsx'  # 原Excel文件地址
    copyer = SheetCopy(src_file_path)  # 实例化对象，此处工作表名称未指定默认为当前激活工作表，目标文件路径未指定则会新建一个新工作薄。
    tag_file = copyer.copy_sheet(origin_col=5, origin_row=8)  # 复制工作表，返回目标ExcelOperate对象，此处目标工作表以第5列，第8行为起始点进行写入。
    tag_file.save(r'D:\Desktop\1111.xlsx')  # 保存目标对象
    ```
+ 工作表对比：

  + 首先通过语句 `from excel_operate import SheetComparison`导入模块。
  + 实例化比较对象 `cmper=SheetComparison(src_excel, cmp_excel)`。

    + 参数 `src_excel`为原工作表的 `ExcelOperate`对象；
    + 参数 `cmp_excel`为目标工作表的 `ExcelOperate`对象。
  + 为比较对象设置表头行和关键列 `cmper.set_title_row(src_title_row, cmp_title_row)`和 `cmper.set_key_col(src_key_col, cmp_key_col)`。

    + 参数 `src_title_row` 和 `cmp_title_row`分别为原工作表和目标工作表的表头行的行号，传入数据为整数类型。
    + 参数 `src_key_col` 和 `cmp_key_col`分别为原工作表和目标工作表的关键列对应的列号，传入数据为整数类型。
  + 进行比较 `report=cmper.compare()`可将比较结果返回给report变量，返回值为 `ExcelOperate`对象。
  + 最后保存对比报告 `report.save(report_path)`即可。
  + 示例代码如下：

    ```
    from excel_operate import ExcelOperate   # 导入模块
    from excel_operate import SheetComparison   # 导入模块

    src_excel = ExcelOperate('tests\比较示例 - 原.xlsx')  # 原工作表对象实例化
    cmp_excel = ExcelOperate('tests\比较示例 - 对比.xlsx')  # 目标工作表对象实例化
    report_path = 'D:/Desktop/对比报告.xlsx'  # 报告保存路径
    cmper = SheetComparison(src_excel, cmp_excel)  # 实例化比较对象
    cmper.set_title_row(2, 2)  # 设置表头行为原工作表第2行，目标工作表第2行
    cmper.set_key_col(2, 2)  # 设置关键列为原工作表第2列，目标工作表第2列
    report = cmper.compare()  # 比较并返回结果
    report.save(report_path)  # 保存结果
    print('对比完成！对比报告已保存至：', report_path)  # 打印结果
    ```
+ 其它 `openpyxl`操作：

  + 使用 `wb = excel.wb`可以*获得 `openpyxl`中的工作薄对象 `wb`，从而进行* `openpyxl`中关于工作薄对象的一切操作；
  + 使用 `ws = excel.ws`*可以获得 `openpyxl`中的工作表对象 `ws，`从而进行* `openpyxl`中关于工作表对象的一切操作。
+ Excel文件的保存：

  + 可直接使用 `excel.save（file_path=None）`*的方法来保存。*
    + 参数 `file_path`为文件保存路径，不指定的情况下为覆盖原文件保存，指定后可以另存为新的路径（如果之前实例化 `ExcelOperate`对象时没有指定 `file_path`，即创建空白Excel文件的情况下，保存时不指定路径的话则会报错。）
+ 其它功能：

  + 列表去重：使用 `list_operate`模块下的 `duplicate_to_only(l, remove=False)`方法可以将传入的列表去重，并返回一个新列表。
    + 参数 `l`为传入的列表。
    + 参数 `remove`可传入布尔值类型数据，默认为 `False`即以重命名的方式去重，如果传入 `True`则会将列表中第二个及之后出现的重复元素直接删除。
  + 列表比较：
    + 使用 `list_operate`模块下的 `is_insert(srcl, tagl)`可以判断目标列表是否为原列表插入元素所得。如果判断为是则返回一个以插入位置索引为键，该索引位置插入的元素个数为值的一个字典；否则返回`None`。
    + 使用 `list_operate`模块下的 `is_delete(srcl, tagl)`可以判断目标列表是否为原列表删除元素所得。如果判断为是则返回一个以删除位置索引为键，该索引位置删除的元素个数为值的一个字典；否则返回`None`。
    + 使用 `list_operate`模块下的 `is_appand(srcl, tagl)`可以判断目标列表是否为原列表后添加元素所得。如果判断为是则返回添加的元素个数；否则返回`None`。

## 四、版本历史

+ V 2.0.2
  + 修复部分BUG。
+ V 2.0.1
  + 新增工作表对比sheet_comparison模块。
  + sheet_copy模块增加可以指定目标工作表的插入起始点功能。
  + 优化sheet_copy模块的代码逻辑，方便从外部直接调用SheetCopy类下的copy_sheet方法。
  + 新增列表操作list_operate模块，可以对列表做一些有趣的操作。
+ V 1.0.3
  + 优化代码逻辑，增加在ExcelOperate类的实例化中如果打开的Excel工作薄中没有sheet_name参数指定的工作表则会以sheet_name参数为名，创建一个工作表。
+ V 1.0.2
  + 优化代码结构，增加函数参数注释。
+ V 1.0.1
  + 正式版本，新增工作表复制模块SheetCopy，优化README文档结构。
+ V 0.2.0
  + 测试版本，优化README描述，正式在pypi.org官方上线版本。
+ V 0.1.0
  + 初始版本、测试版本。

---
