#print(" ........................................................ ..............................................") 
#print("LIBRARIES...") 
import random 
import os 
import datetime 
import qrcode 
import pandas as pd 
import time 
import mysql.connector as sql 
from tkinter import* 


#print(" ........................................................ ..............................................") 
#print("DEFINED FUNCTIONS..") 
def extvals(): 
 print(f"The Name is : {nameval.get()}") 
def xtvals(): 
 print(f"The Class is : {clssval.get()}") 
def vbvals(): 
 print(f"The Second is : {secval.get()}") 
def hvals(): 
 print(f"The Verification Code is : {vercodeval.get()}")  
def getvals(): 
 print(f"The Entered Roll Number is: {rollnum.get()}")
 


#print(" ........................................................ ..............................................") 
#print(" INITIAL STRUCTUIZATION..") 
print(" ") 
print(" ") 
print("***............................................................ .......***") 
print(" ") 
print(" Welcome To VIT Project Report\ n \t\t\t [2020-2021]") 
print(" ") 
print(" ***Enter The Following Details To Generate ID_Car d**** ") 
print(" ") 
print("............................................................... ..........") 


def openNewWindow(): 
 newWindow = Toplevel(root1) 
 newWindow.geometry("850x425") 
 a=Label(newWindow, text =" \nVELLORE INSTITUTE OF TECHNOLOGY",bg="grey",font=('Bahnschrift SemiBold SemiConden',40,'bold'),fg="navyblue"). pack() 
 b=Label(newWindow, text =" " ,bg="grey" ).pack()  
 c=Label(newWindow, text =" " ,bg="grey" ).pack()  
 d=Label(newWindow, text =" " ,bg="grey" ).pack()  
 e=Label(newWindow, text =" ID - CARD GENERATOR " ,bg="grey" , fg="maroon" ,font=('ink free',30,'bold')).pack() 
 f=Label(newWindow, text =" " ,bg="grey" ).pack()  
 g=Label(newWindow, text =" " ,bg="grey" ).pack()  
 '''bt = Button(newWindow, text ="CLICK TO GENERATE ID CARD\n ---->", command = countDown ) 
 bt.pack(ipadx = 20)''' 
 newWindow.configure(bg="grey")

 bt = Button(newWindow, text ="CLOSE THE TAB TO GENERATE ID CARD" , fg="darkgreen", font=(5) ) 
 bt.pack(ipadx = 20) 
 newWindow.mainloop() 
  
root1 = Tk() 
root1.geometry("750x350") 
f = Label(root1, text ="" ,bg="grey") 
f.pack() 
label = Label(root1, text ="WELCOME TO VIT PROJECT REPORT\n[2021-2022]\n" ,font=(20),fg='red',bg="grey")  
label.pack() 
label2 =Label(root1, text= "NAME: \t\tManu Mishra\n\nREG_NO:\t\t21BCE1639\n\n SUBMITTED TO:\t Ms. KAVYA ALLURU MA'AM\n",font=(0) ,bg="grey") 
label2.pack() 
btn = Button(root1, text ="CLICK\n --->", command = openNewWindow) 
btn.pack(ipadx = 20 ) 
root1.configure(bg="grey") 
root1.mainloop() 

#print(" ........................................................ ..............................................") 
#print(" TIME AND DATE..") 

os.system("Title: ID CARD Generator by Grasp Coding") 
d_date = datetime.datetime.now() 
reg_format_date = d_date.strftime("DATE: %d-%m-%Y\nTIME: %I:%M:%S %p") 
print(" ") 
print(reg_format_date)

#print(" ........................................................ ..............................................") 
#print(" GRAPHICAL USER INTERFACE [GUI] . TKINTER")
#  
root = Tk() 
root.geometry("650x270") 
root .title("ID-Card Generator") 


extt= Label(root, text="                           VELLORE INSTITUTE OF TECHNOLOGY ",fg='navyblue',bg='grey').grid(row=0,column=0) 
abc = Label(root, text="                           Welcome To ID Generator ",fg='navyblue',bg='grey').grid(row=1,column=0) 
ac = Label(root, text="                            VERIFICATION FOR CLASS TEACHERS ",fg='MAROON',bg='grey').grid(row=2,column=0) 

#print("Labels for Entry statements for TEACHER VERIFICATION") 
name = Label(root, text="Enter Your Name :" ,fg='black', bg='grey').grid(row=3, column=0) 
clss = Label(root, text="Enter Your Class :" ,fg='black', bg='grey').grid(row=4, column=0) 
sec = Label(root, text="Enter Your Section :" ,fg='black', bg='grey').grid(row=5, column=0) 
vercode = Label(root, text="Enter Teacher Verif. Code :" ,fg='black', bg='grey').grid(row=6, column=0) 


#print("Values optimization and input statements for TEACHER VERIFICAT ION") 
nameval = StringVar() 
nameentry = Entry(root, textvariable = nameval)  
nameentry.grid(row=3 , column=1) 
clssval = StringVar() 
clssentry = Entry(root, textvariable = clssval)  
clssentry.grid(row=4 , column=1) 
secval = StringVar() 
secentry = Entry(root, textvariable = secval)  
secentry.grid(row=5 , column=1)


vercodeval = StringVar() 
vercodeentry = Entry(root, textvariable = vercodeval)  
vercodeentry.grid(row=6 , column=1) 
#print("Button commands and statements for TEACHER VERIFICATION") 
B1 = Button(text="SUBMIT",fg='white',bg='grey', command= lambda:[extvals(), xtvals(),vbvals(), hvals()]) 
B1.grid(row=8, column=1) 
print("") 
print("") 
#print("Labels for Entry statements for ROLL NUMBER ENTRY") 
roll = Label(root, text="Enter the Roll no: >",fg='red',bg='grey').grid(row=10, column=0) 
#print("Values optimization and input statements for ROLL NUMBER ENTRY ") 


rollnum = StringVar() 
rollentry = Entry(root, textvariable = rollnum)  
rollentry.grid(row=11 , column=1) 
#print("Button commands and statements for ROLL NUMBER ENTRY") 
B2 = Button(text="APPLY",fg='white',bg='grey', command=getvals) 
B2.grid(row=13, column=1) 
  
root.configure(bg="grey") 
root.mainloop() 
rooto = Tk() 
rooto.geometry("700x350") 
f = Label(rooto, text ="" ,bg="grey") 
f.pack()
 
label = Label(rooto, text ="\n\n\n" ,font=(20),fg='yellow',bg="grey" )  
label.pack() 
label2 =Label(rooto, text= " PLEASE WAIT YOUR ID CARD IS GENERATING\n\n", fg="blue" , font=('ink free',20,'bold') ,bg ="grey") 
label2.pack() 
btnn = Button(rooto, text ="\nClose The Tab Again to get IDard\n") 
btnn.pack(ipadx = 20 ) 
rooto.configure(bg="grey") 
rooto.mainloop() 

#print("CREATING TMER") 
def countDown(): 
  
 lbl1.config(bg='yellow', font=('times', 20, 'bold'))   
 for k in range(15, 0, -1): 
    lbl1["text"] = (k) 
 root1.update() 
 time.sleep(2) 
 lbl1.config(bg='maroon') 
 lbl1.config(fg='white') 
 lbl1["text"] = " Your ID Card is Generated Succesfully!!!\n\nCLOSE THE TAB TO GET THE ID CARD " 
  
  
root1 = Tk() 
root1.geometry("670x350") 
root1.title("Result") 
lbl1 = Label() 
lbl1.pack(fill=BOTH, expand=1) 
countDown() 
root1.mainloop()
 
#print(" PRINTING DATABASE AND REQUIRED INPUTS AND STATEMENTS") 
print("") 
print("") 
print("***..........................YOUR CLASS DATABASE............... ...............***") 
print("***............................................................ ...............***") 
print("") 
print("") 
#print(" CONNECING MSQL WITH PYTHON") 

mycon= sql.connect(host= "localhost",user="root", passwd="Hacker", database="idcard") 
if mycon.is_connected(): 
  
 info=pd.read_sql("select * from idcardd;",mycon) 
 print(info) 
 print() 
 print("***........................................................ ................***") 
 print("***........................................................ ................***") 
 rollnum= rollnum.get() 
  
 name="select * from idcardd where Roll='%s';" %(rollnum,)  
 df1=pd.read_sql(name,mycon) 
 name=list(df1['Name']) 
 name_stud=name[0] 
 classStud="select Class from idcardd where Roll='%s';" %(rollnum,)  
 df2=pd.read_sql(classStud,mycon) 
 classstud=list(df2['Class']) 
 class_stud=classstud[0]


 section="select Section from idcardd where Roll='%s';" %(rollnum,)  
 df3=pd.read_sql(section,mycon) 
 sec=list(df3['Section']) 
 sec_stud=sec[0] 

 '''house="select House from idcardd where Roll='%s';" %(rollnum,)  
 df4=pd.read_sql(house,mycon) 
 house1=list(df4['House']) 
 house_stud=house1[0]''' 

 address="select Address from idcardd where Roll='%s';" %(rollnum,)  
 df5=pd.read_sql(address,mycon) 
 add=list(df5['Address']) 
 add_stud=add[0] 

 contact="select Contact from idcardd where Roll='%s';" %(rollnum,)  
 df6=pd.read_sql(contact,mycon) 
 cont=list(df6['Contact']) 
 cont_stud=cont[0] 

 transport="select Transport from idcardd where Roll='%s';" %(rollnum,) 
 df7=pd.read_sql(transport,mycon) 
 trans=list(df7['Transport']) 
 trans_stud=trans[0] 

else: 
 print("error") 
#print("IMAGE STRUCTURE AND FORMATION") 
from PIL import Image, ImageDraw, ImageFont 
image = Image.new('RGB', (1500, 1000), (211, 211, 211)) 
draw = ImageDraw.Draw(image) 
font = ImageFont.truetype('arial.ttf', size=40) 
# Starting position of the message


(x, y) = (50, 50) 
#message = input('\nEnter Your School Name: ') 
message =" Vellore Institute of Technology" 
company = message 
color = 'rgb(0, 0, 128)' 
font = ImageFont.truetype('arial.ttf', size=80) 
draw.text((x, y), message, fill=color, font=font) 

# adding an unique id number. You can manually take it from user 
(x, y) = (1300, 75) 
idno = random.randint(10000000, 90000000) 
message = str('ID: ' + str(idno)) 
color = 'rgb(128, 0, 0)' # black color 
font = ImageFont.truetype('arial.ttf', size=30) 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 250) 


#message = input('Enter Your Full Name: ') 
message= name_stud 
name=message 
message = str('# Name:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
font = ImageFont.truetype('arial.ttf', size=45) 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 350) 

#message = input('Enter Your Class: ') 
message = class_stud 
message = str('# Class:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 450) 


#message = int(input('Enter Your Section: ')) 
message = sec_stud

message = str('# Section:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 550) 


#message = input('Enter Your Address: ') 
message = add_stud 
message = str('# Address:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 650) 


#message = input('Enter Your Contact number: ') 
message = cont_stud 
message = str('# Contact No.:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 750) 


#message = input('Enter Your Transport: ') 
message = trans_stud 
message = str('# Transport:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font)


'''#message = int(input('Enter Your House: ')) 
message = house_stud 
message = str('# House:- ' + str(message)) 
color = 'rgb(0, 0, 0)' # black color 
draw.text((x, y), message, fill=color, font=font) 
(x, y) = (50, 650)'''

# save the edited image 
image.save(str(name) + '.png') 
img = qrcode.make(str(company) + str(idno)) # this info. is added in QR code, also add other things

img.save(str(idno) + '.bmp') 
til = Image.open(name + '.png') 
im = Image.open(str(idno) + '.bmp') # 25x25 
til.paste(im, (1050, 200)) 
til.save(name + '.png') 


#print(" PRINTING FINAL STATEMENTS AND ENDITING OF PROJECT") 
print('\n\n\n>> CONGRATULATIONS \n>>   Your ID Card Successfully Created') 
print(('\n>> Student Name is: '+ name +'.')) 
print("............................................................... ..........") 
input('\n>> Press any key to Close program...') #print("THANK YOU // END OF REPORT
