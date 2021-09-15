from collections import Counter  # 引入Counter

a = [29, 36, 57, 12, 79, 43, 23, 56, 28, 11, 14, 15, 16, 37, 24, 35, 17, 24, 33, 15, 39, 46, 52, 13,15,15]


def findDuplicatedElements(mylist):
    b = dict(Counter(mylist))
    return [key for key, value in b.items() if value > 1]  # 只展示重复元素
    # print({key: value for key, value in b.items() if value > 1})  # 展现重复元素和重复次数

print(findDuplicatedElements(a))
