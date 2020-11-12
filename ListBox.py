from tkinter import *

root = Tk()

# height 默认显示10条数据 listBox = Listbox(root, height=11)

# 为了在某个组件上安装垂直滚动条，需要做两件事：
# 1.设置该组件的yscrollcommand选项为Scrollbar组件的set()方法
# 2.设置Scrollbar组件的command选项为该组件的yview()方法

sb = Scrollbar(root)
sb.pack(side=RIGHT, fill=Y)

listBox = Listbox(root, yscrollcommand=sb.set)

for i in range(100):
    listBox.insert(END, i)  # ListBox添加数据

listBox.pack()
listBox.see(20)  # 调整列表框的位置，使得 index 参数指定的选项是可见的

sb.config(command=listBox.yview)


def delete():
    listBox.delete(ACTIVE)  # 删除选中的


delButton = Button(root, text="删除", command=delete)
delButton.pack(padx=10, pady=10)

root.mainloop()