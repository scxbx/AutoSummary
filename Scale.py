import tkinter as tk

root = tk.Tk()
root.geometry('300x240')

value = tk.StringVar()
v = tk.StringVar()
b1 = tk.Scale(root, length=200,
              orient=tk.HORIZONTAL, variable=value)
b1.pack()
b3 = tk.Entry(root, textvariable=v)


def set1():
    value.set(b3.get())


def get():
    v.set(value.get())


b2 = tk.Button(root, text='Set', command=set1)
b2.pack()
b4 = tk.Button(root, text='Get', command=get)
b4.pack()
b3.pack()
root.mainloop()