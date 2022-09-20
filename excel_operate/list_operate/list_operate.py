# -*- encoding: utf-8 -*-
'''
@File    :   list_operate.py
@Author  :   北极星光 
@Contact :   light22@126.com
'''


from readline import insert_text


def duplicate_to_only(l: list, remove=False):
    '''---
    ### 传入一个列表，去除列表中的重复项
    ---
    + l: 目标列表
    + remove: 默认参数False将第二个及之后重复的项目重命名为name(1),name(2)...的格式，传入True则直接删除列表中第二次及以后出现的所有重复项。
    '''
    tmp_list = [] 
    if remove == False:
        for i in l:
            if i not in tmp_list:
                tmp_list.append(i)
            else:
                tmp_list.append(f"{i}({tmp_list.count(i) + 1})")
        return tmp_list
    else:
        for i in l:
            if i not in tmp_list:
                tmp_list.append(i)
        return tmp_list


def is_insert(srcl: list, tagl: list):
    '''---
    ### 传入两个列表，判断目标列表是否为原列表插入元素所得。如果判断为是则返回一个以插入位置索引为键，该索引位置插入的元素个数为值的一个字典；否则返回None。
    ---
    + srcl: 原列表
    + tagl: 目标列表
    '''
    if len(srcl) >= len(tagl):
        return None
    # 一般认为，插入元素为前插，最后一项不一样，则不能通过插入使二者一样
    if srcl[-1] != tagl[-1]:
        return None
    insert_info = {}
    # [1, 2, 3]
    # [0, 1, 3, 3]
    target_index = 0
    for i in range(len(tagl)):
        if tagl[i] != srcl[target_index]:
            insert_info[target_index] = insert_info.get(target_index, 0) + 1
        else:
            target_index += 1
    if target_index == len(srcl) - 1:
        return None
    return insert_info

def is_appand(srcl: list, tagl: list):
    '''---
    ### 传入两个列表，判断目标列表是否为原列表后添加元素所得。如果判断为是则返回添加的元素个数；否则返回None。
    ---
    + srcl: 原列表
    + tagl: 目标列表
    '''
    if len(tagl) <= len(srcl):
        return None
    for i in range(len(srcl)):
        if srcl[i] != tagl[i]:
            return None
    return len(tagl) - len(srcl)

def is_delete(srcl: list, tagl: list):
    '''---
    ### 传入两个列表，判断目标列表是否为原列表删除元素所得。如果判断为是则返回一个以删除位置索引为键，该索引位置删除的元素个数为值的一个字典；否则返回None。
    ---
    + srcl: 原列表
    + tagl: 目标列表
    '''
    if len(tagl) <= len(srcl):
        return None
    dic_info = {}
    for i in range(len(tagl)):
        if tagl[i] != srcl[i - sum(dic_info.values())]:
            dic_info[i] = dic_info.get(i, 0) + 1
        if i - sum(dic_info.values()) == len(srcl) - 1 and i < len(tagl) - 1:
            dic_info[i + 1] = dic_info.get(i + 1, 0) + len(tagl) - 1 - i
            return dic_info
    return dic_info


if __name__ == '__main__':
    l1 = [0, 1, 'x1', 'x2', 2, 3, 4, 5, 6, 7, 8, 9]
    l2 = [0, 1, 2, 3, 4, 5, 6, 7, 'x1', 'x2', 8, 9]

    ins_info = is_insert(l2, l1)
    del_info = is_delete(l2, l1)
    print('insert:', ins_info)
    print('delete:', del_info)
