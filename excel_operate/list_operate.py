# -*- encoding: utf-8 -*-
'''
@File    :   list_operate.py
@Author  :   北极星光 
@Contact :   light22@126.com
'''


def duplicate_to_only(l: list, remove=False):
    '''===================================\n
    传入一个列表，去除列表中的重复项
    l: 目标列表
    remove: 默认参数False将第二个及之后重复的项目重命名为name(1),name(2)...的格式，传入True则直接删除列表中第二次及以后出现的所有重复项。
    '''
    tmp_list = [i for i in l]  # 复制一份列表以便后续操作
    if remove == False:
        for i in range(len(l)):
            if l.count(l[i]) > 1:
                if i != l.index(l[i]):
                    for n in range(l.count(l[i])):
                        if f'{l[i]} ({n+1})' not in tmp_list:
                            tmp_list[i] = f'{l[i]} ({n+1})'
                            break
        return tmp_list
    else:
        tmp_list.reverse()
        for i in l:
            if tmp_list.count(i) > 1:
                tmp_list.remove(i)
        tmp_list.reverse()
        return tmp_list


def is_insert(srcl: list, tagl: list):
    '''===================================\n
    传入两个列表，判断目标列表是否为原列表插入元素所得。如果判断为是则返回一个以插入位置索引为键，该索引位置插入的元素个数为值的一个字典；否则返回None。
    srcl: 原列表
    tagl: 目标列表
    '''
    insert_info = {}
    for i in range(len(tagl)):
        if tagl[i] not in srcl:
            for ins in range(1, len(tagl)-i):
                if tagl[i+ins] in srcl:
                    if srcl.index(tagl[i+ins]) not in insert_info:
                        insert_info[srcl.index(tagl[i+ins])] = 1
                    else:
                        insert_info[srcl.index(tagl[i+ins])] += 1
                    break
    if insert_info == {}:
        return None
    else:
        return insert_info


def is_appand(srcl: list, tagl: list):
    '''===================================\n
    传入两个列表，判断目标列表是否为原列表后添加元素所得。如果判断为是则返回添加的元素个数；否则返回None。
    srcl: 原列表
    tagl: 目标列表
    '''
    append_info = None
    if tagl[-1] != srcl[-1] and srcl[-1] in tagl:
        append_info = len(tagl[tagl.index(srcl[-1])+1:])
    return append_info


def is_delete(srcl: list, tagl: list):
    '''===================================\n
    传入两个列表，判断目标列表是否为原列表删除元素所得。如果判断为是则返回一个以删除位置索引为键，该索引位置删除的元素个数为值的一个字典；否则返回None。
    srcl: 原列表
    tagl: 目标列表
    '''
    delete_info = {}
    for i in srcl:
        if i not in tagl:
            if srcl.index(i) not in delete_info:
                delete_info[srcl.index(i)] = 1
            else:
                delete_info[srcl.index(i)] += 1
    if delete_info == {}:
        return None
    else:
        return delete_info


if __name__ == '__main__':
    srcl = [0, 1, 2, 3, 4,None,None, 5, 6, 7, 8, 9]
    tagl = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]

    l = duplicate_to_only(srcl)
    print(l)