from tkinter import *

MyRootDialog=Tk()
#Set Tkinter Size And Location
MyRootDialog.geometry("200x150+100+100")

#Set Tkinter Title
MyRootDialog.title("Get Input Value in Label")

# Define Function for get Value
def Get_MyInputValue():
     getresult = MyEntryBox.get()
     myTKlabel['text'] = getresult

# Create Tkinter Entry Widget
MyEntryBox = Entry(MyRootDialog, width=20)
MyEntryBox.place(x=5, y=6)

myTKlabel = Label(MyRootDialog, borderwidth=1, relief="ridge", height=3, width=25)
myTKlabel.place(x=4, y=42)

#command will call the defined function
MyTkButton = Button(MyRootDialog, height=1, width=10, text="Get text", command= Get_MyInputValue)
MyTkButton.place(x=4, y=112)

mainloop()