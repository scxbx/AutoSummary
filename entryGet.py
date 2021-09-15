import tkinter as tk

master = tk.Tk()

tk.Label(master, text="身份证号码：").pack()
# tk.Label(master, text="作者：").grid(row=1)

e1 = tk.Entry(master)
# e2 = tk.Entry(master)
e1.pack()
# e2.grid(row=1, column=1, padx=10, pady=5)
text1 = tk.Text(master, width=30, height=20)

def checkIDNumber(num_str):
    str_to_int = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5,

                  '6': 6, '7': 7, '8': 8, '9': 9, 'X': 10}

    check_dict = {0: '1', 1: '0', 2: 'X', 3: '9', 4: '8', 5: '7',

                  6: '6', 7: '5', 8: '4', 9: '3', 10: '2'}

    if len(num_str) != 18:
        return u"身份证号: %s 位数不为18" % num_str
    if ' ' in num_str.strip():
        return u"身份证号: %s 中间存在空格" % num_str

    check_num = 0

    for index, num in enumerate(num_str):

        if index == 17:

            right_code = check_dict.get(check_num % 11)

            if num != right_code:
                print(u"身份证号: %s 校验不通过, 正确尾号应该为：%s" % (num_str, right_code))
                return u"%s 校验不通过" % num_str

        check_num += str_to_int.get(num) * (2 ** (17 - index) % 11)
    return 'pass'

def check():
    # print("身份证：《%s》" % e1.get())
    str = checkIDNumber(e1.get())
    text1.insert("end", str + '\n')
    # print("作者：%s" % e2.get())
    # e1.delete(0, "end")
    # e2.delete(0, "end")

# original author: j_hao104
# site: https://my.oschina.net/jhao104/blog/756241


tk.Button(master, text="校验", width=10, command=check).pack()
# tk.Button(master, text="退出", width=10, command=master.quit).pack()
text1.pack()
master.mainloop()