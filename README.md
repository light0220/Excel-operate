# 欢迎使用

## 一、概述

本模块为方便大家更便捷使用Python来操作Excel文件而编写，基于openpyxl进行编写，所以在使用本模块前请确保您的计算机中已安装有Python环境以及openpyxl库.

作者：北极星光

E-mail：light22@126.com

## 二、功能介绍

+ 简化openpyxl库的导入流程，可以使用openpyxl库中的所有操作。
+ 方便快捷插入和删除行（列）。解决直接使用工作表对象的insert_rows方法时，插入或删除的行（列）的位置后方行（列）的行高（列宽）以及如果存在已合并单元格的情况下，已合并的单元格信息等格式不会因插入或删除的行（列）而下（右）移或上（左）移，从而导致表格的整体格式甚至数据受到影响。以及插入行（列）的行高（列宽）的快捷设定。通过本模块的方法插入或删除行（列）后的效果与在Excel程序中插入或删除行（列）的效果可以做到近乎完全相同。
+ 快速复制工作表（未实装）。
+ 其它功能敬请期待

## 三、使用方法

+ 导入本模块方法：将 *excel_operate*文件夹复制到当前工程目录下，或者您的用户主目录（通常情况下为C:\users\用户名）下的AppData/Roaming/Python/Python310/site-packages目录下，并在要导入本模块的代码中写入导入语句：*from excel_operate import ExcelOperate*
+ 对象实例化方法：*excel = ExcelOperate(file_path=None,sheet_name=None)*，此方法可获得一个名为*excel*的对象，其中参数*file_path*为文件路径，指定后打开文件路径对应的Excel文件（相当于*openpyxl*中的*load_workbook*方法）默认不指定则会新建一个空白Excel文件（相当于*openpyxl*中的*Workbook*类的实例化）；参数*sheet_name*为工作表名，默认不指定为打开当前激活的工作表（相当于工作薄对象的*active*方法）。
+ 插入或删除行（列）：插入行*excel.insert_rows(idx,amount=1,height='before')*，删除行*excel.delete_rows(idx,amount=1)*，插入列*excel.insert_cols(idx,amount=1,width='before')*，删除列*excel.delete_cols(idx,amount=1)*。其中参数*idx*为插入或删除的起始行（列），参数*amount*为插入或删除的行（列）数，不指定的情况下默认插入或删除1行（列），参数height和width仅在插入行或列时使用，可指定为整数或者浮点数来设置插入的行高或列宽，不指定的情况下默认参数为'before'，该参数可以让插入的行或列继承插入前上一行（列）的行高或列宽。另外插入行时，height参数也可指定为None，此参数可以设置插入行的行高为自动行高。
+ 其它*openpyxl*操作：*wb = excel.wb*可以获得工作薄对象，从而进行*openpyxl*中关于工作薄对象的一切操作；*ws = excel.ws*可以获得工作表对象，从而进行*openpyxl*中关于工作表对象的一切操作。
+ Excel文件的保存：可直接使用*excel.save（file_path=None）*的方法来保存，*file_path*参数默认不指定的情况下为覆盖原文件保存，指定后可以另存为新的路径（当然如果之前实例化*excel*对象时没有指定*file_path*，即创建空白Excel文件的情况下，保存时不指定路径的话则会报错。）

---

---
