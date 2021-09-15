
key_list = ['家庭','代表','意见','21']
str = '家庭\n代表         \n意见'

'''
for key in key_list:
    if key in str:
        print(key)
'''

def checkAllKeysInAString(list, str):
    for key in list:
        if key not in str:
            return False
    return True;

print(checkAllKeysInAString(key_list, str))
