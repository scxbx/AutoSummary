#!/usr/bin/python
# -*- coding: UTF-8 -*-

from tkinter import *

top = Tk()

text = StringVar()
text.set('old')

def change():
    text.set

l = Label(top, bg='white', width=20, text='empty')
l.pack()

L1 = Label(top, text="网站名")
L1.pack(side=LEFT)
E1 = Entry(top, bd=5)
E1.pack(side=RIGHT)



top.mainloop()