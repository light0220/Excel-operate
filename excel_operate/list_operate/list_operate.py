# -*- encoding: utf-8 -*-
'''
@File    :   list_operate.py
@Author  :   北极星光 
@Contact :   light22@126.com
'''


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
                for n in range(l.count(i)):
                    if f'{i} ({n+1})' not in tmp_list:
                        tmp_list.append(f'{i} ({n+1})')
                        break
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
    insert_info = {}
    n = 0
    for i in range(len(srcl)):
        if tagl[i + n] != srcl[i]:
            if srcl[i] not in tagl[i + n:]:
                n -= 1
            else:
                for ins in range(1, len(tagl) - n - i):
                    if tagl[i + n + ins] == srcl[i]:
                        insert_info[srcl.index(tagl[i + n + ins])] = ins
                        n += ins
                        break
    if insert_info == {}:
        return None
    else:
        return insert_info


def is_appand(srcl: list, tagl: list):
    '''---
    ### 传入两个列表，判断目标列表是否为原列表后添加元素所得。如果判断为是则返回添加的元素个数；否则返回None。
    ---
    + srcl: 原列表
    + tagl: 目标列表
    '''
    append_info = None
    if tagl[-1] != srcl[-1] and srcl[-1] in tagl:
        append_info = len(tagl[tagl.index(srcl[-1])+1:])
    return append_info


def is_delete(srcl: list, tagl: list):
    '''---
    ### 传入两个列表，判断目标列表是否为原列表删除元素所得。如果判断为是则返回一个以删除位置索引为键，该索引位置删除的元素个数为值的一个字典；否则返回None。
    ---
    + srcl: 原列表
    + tagl: 目标列表
    '''
    delete_info = {}
    n = 0
    for i in range(len(tagl)):
        if srcl[i + n] != tagl[i]:
            if srcl[i + n] not in tagl[i:]:
                for ins in range(1, len(srcl) - n - i):
                    if srcl[i + n + ins] == tagl[i]:
                        delete_info[i + n] = ins
                        n += ins
                        break
                    delete_info[i + n] = 1
            else:
                for ins in range(1, len(tagl) - i):
                    if tagl[i + ins] == srcl[i + n]:
                        n -= 1
                        break
    if delete_info == {}:
        return None
    else:
        return delete_info


if __name__ == '__main__':
    l1 = [0, 1, 'x1', 'x2', 2, 3, 4, 5, 6, 7, 8, 9]
    l2 = [0, 1, 2, 3, 4, 5, 6, 7, 'x1', 'x2', 8, 9]
    # l = ['a', 'a', 'a', 'a', 'a', 'a', 'a', 'a', 'a']
    # l0 = duplicate_to_only(l,False)
    # print(l0)

    ins_info = is_insert(l1, l2)
    del_info = is_delete(l1, l2)
    print('insert:', ins_info)
    print('delete:', del_info)
