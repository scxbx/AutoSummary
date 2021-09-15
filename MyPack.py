import tkinter as tk

root = tk.Tk()


topFrame = tk.Frame(root)
topFrame.pack(side=tk.TOP)


leftFrame = tk.Frame(topFrame)
leftFrame.pack(side=tk.LEFT)

rightFrame = tk.Frame(topFrame)
rightFrame.pack(side=tk.RIGHT)

redbutton = tk.Button(leftFrame, text="Red", fg="red")
redbutton.pack()
redbutton2 = tk.Button(leftFrame, text="Red", fg="red")
redbutton2.pack()

greenbutton = tk.Button(rightFrame, text="green", fg="green")
greenbutton.pack()
greenbutton2 = tk.Button(rightFrame, text="green", fg="green")
greenbutton2.pack()

bluebutton = tk.Button(leftFrame, text="Blue", fg="blue")
bluebutton.pack()
if __name__ == '__main__':
    root.mainloop()
