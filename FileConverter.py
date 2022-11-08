'''
Created on Sep 26, 2022

Last Edited on Nov 7, 2022

@author: acrie
'''

import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter.filedialog import askdirectory as ad
from tkinter.messagebox import showinfo
import PIL
from PIL import Image
import docx
import json

global wW
wW = 950 #window Width
wWstr = str(wW)
root = Tk()
root.geometry(wWstr + "x750") #width 950 height 750
root.resizable(False, False)
root.title("Filosopher's Stone")

bgframe = Frame(root, bg = 'yellow')
bgframe.place(width = wW, height = 750)

my_details = {
    "set1" : "1",
    "ypos1" : "150",
    "allowsectdel1":"No",
    "fts1": "img",
    "dts1": "folderselect",
    "seldes1": "none"
    }

#set = whole set in a row
#allowsectdel has a 'yes' or 'no' for if a section will have the option to be deleted. All, but the first set will be removable
#fts = file type selected: image (img) or text (txt) types (switch button)(is for file browser and the end type option menu)
#dts = destination type selection: folder or application (which could itself be the destination or sends the converted files to their intended final location)
#seldes = selected destination ex. picture folder
#delorig = delete original files? 'Yes' deletes original file after created a converted copy. 'No' keeps original files after creating converted copy.
#convert = conversion button

with open('fileconverterdata.json', 'w') as f:
    json.dump(my_details, f)

global ypad
ypad = 7
global xpad
xpad = 15
global total_sets
framedict = {} #frame dictionary
yposdict = {} #y position
delbtndict = {} #section delete button
ftsbtndict = {} #fts = file type switch
ofdbtndict = {} #ofd = open file directory
filelistdict = {} #temporarily stores list of files for conversion's sake (not saved in json)
dftmenudict = {} #dft = destination file type
dftoptdict = {} #stores dft options, includes selected
dtsdict = {} #des = destination type switch
dcbtndict = {} #dc = destination chooser
seldesdict = {} #selected destination
convertdict = {} #convert buttons



img_type_list = [".jpg", ".jpeg",".png", ".gif", ".dds"]

txt_type_list = [".txt",".docx", ".doc"] #future support for excel and pdf files in the works


def updatejson():
    with open('fileconverterdata.json', 'w') as f:
        json.dump(obj, f)

def select_txtfiles(n):
    txtfiles = fd.askopenfilenames(
        title='Select text files',
        initialdir='/',
        filetypes=(
        ('All Supported Text Files', '*.txt .docx .doc'),
        ('Txt Files','*.txt'),
        ('Word Documents','*.docx .doc')
        )
    )
    
    addtolist(n, txtfiles)

def select_imgfiles(n):
    imgfiles = fd.askopenfilenames(
        title='Select image files',
        initialdir='/',
        filetypes=(
        ('All Supported Image Files', '*.jpg .jpeg .png .gif .dds'),
        ('JPEG/JPG Files','*.jpg .jpeg'),
        ('PNG Files','*.png'),
        ('GIF Files','*.gif'),
        ('DDS Files','*.dds')
        )
    )
    
    addtolist(n, imgfiles)


def addtolist(n, f): #f here equals files, not frame
    fl = "fl" + n
    filelistdict[fl].clear()
    for x in f:
        filelistdict[fl].append(x)

def select_folder(n):
    dest = "seldest" + n
    folderdir = ad(parent=root,initialdir="/",title='Please select a directory')
    seldesdict[dest] = folderdir
    obj[dest] = seldesdict[dest]
    updatejson()


    

def imgtxtswitch(n):
    f = "fts" + n
    o = "filebrowser" + n
    if obj[f] == "img":
        obj[f] = "txt"
        ofdbtndict[o].configure(command = lambda n=n: select_txtfiles(n))
        alteroptions(n, txt_type_list)
    else:
        obj[f] = "img"
        ofdbtndict[o].configure(command = lambda n=n: select_imgfiles(n))
        alteroptions(n, img_type_list)
    updatejson()
    

def alteroptions(n, l):
    fft = "finalfiletype" + n
    dftoptdict[fft].set('')
    dftmenudict[fft]['menu'].delete(0,'end')
    for opt in l: 
        dftmenudict[fft]['menu'].add_command(label=opt, command=tk._setit(dftoptdict[fft], opt))
    dftoptdict[fft].set(l[0]) # default value
    
def desttypeswitcher(n): #currently only set up to allow for folder destination, other destination types in the works
    dest = 'dts' + n
    if obj[dest] == "folderselect":
        obj[dest] = "folderselect"

def makexbutton(a, f):
    standinxpic = PhotoImage(file = r'C:\Users\acrie\OneDrive\Pictures\Camera Roll\xbuttonstandin.png')
    if obj[a] == "No":
        delbtndict[a] = tk.Label(f, image = standinxpic,height= 1, width=5)
    else:
        delbtndict[a] = tk.Button(f, text = "X", height= 1, width=5)
        
    delbtndict[a].grid(row = 0, column = 1, padx = xpad, pady=ypad)

def makeftswitch(fts, num, f):
    ftsbtndict[fts] = tk.Button(f, text='Image Files', bg = 'blue', command= lambda num=num:imgtxtswitch(num),height= 1, width=10)
    ftsbtndict[fts].grid(row = 0, column = 2, padx = xpad, pady=ypad)

def makefbbtn(fts, n, f):
    ofd = "filebrowser" + n
    fl = "fl" + n
    if obj[fts] == "img":
        ofdbtndict[ofd] = tk.Button(f, text='Select File(s)', bg= '#FFD700', command= lambda n=n: select_imgfiles(n),height= 1, width=10)
    else:
        ofdbtndict[ofd] = tk.Button(f, text='Select File(s)', bg= '#FFD700', command=lambda n=n: select_txtfiles(n),height= 1, width=10)
        
    ofdbtndict[ofd].grid(row = 0, column = 3, padx = xpad, pady=ypad)
    filelistdict[fl] = []    

def makefinalfiletypechooser(fts, n, f):
    fft = "finalfiletype" + n
    dftoptdict[fft] = tk.StringVar(root)
    if obj[fts] == "img":
        dftoptdict[fft].set(img_type_list[0])
        dftmenudict[fft]= tk.OptionMenu(f, dftoptdict[fft],*img_type_list) #drop down menu for intended destination file type
    else:
        dftoptdict[fft].set(txt_type_list[0])
        dftmenudict[fft]= tk.OptionMenu(f, dftoptdict[fft],*txt_type_list) #drop down menu for intended destination file type
    
    dftmenudict[fft].config(bg="white", activebackground='#00acdb', highlightthickness= 0 ,width=5)
    dftmenudict[fft]['menu'].config(bg="white", activebackground='#00acdb')
    dftmenudict[fft].grid(row = 0, column = 4, padx = xpad, pady=ypad)


    
def makedesttypeswitch(n, f):
    dest = 'dts' + n
    dtsdict[dest] = tk.Button(f, text='Local', bg = 'blue', command= lambda n=n: desttypeswitcher(n),height= 1, width=10) #destination type switch
    dtsdict[dest].grid(row = 0, column = 5, padx = xpad, pady=ypad)



def makedestchooser(n, f):
    dc = "dc" + n
    dcbtndict[dc] = tk.Button(f, text='Select Destination', bg= '#FFD700', command=lambda n=n:select_folder(n),height= 1, width=15) #destination
    dcbtndict[dc].grid(row = 0, column = 6, padx = xpad, pady=ypad)
    
def makeconverter(n, f):
    cb = "converter" + n
    convertdict[cb] = tk.Button(f, text='Convert', bg= '#FFD700', command=lambda n=n:convert(n),height= 1, width=10)
    convertdict[cb].grid(row = 0, column = 7, padx = xpad, pady=ypad)
    
    
def convert(n):
    dts = "dts" + n
    fts = "fts" + n
    fft = "finalfiletype" + n
    fl = "fl" + n
    dest = "seldest" + n
    d = seldesdict[dest].replace("/","\\") #changes / to \
    ext = dftoptdict[fft].get() #intended file type extension
    if obj[dts] == "folderselect":
        for file in filelistdict[fl]:
            s1 = file.replace("/","\\") #changes / to \
            s2 = s1.rfind("\\") #finds location of final \
            s3 = s1.rfind(".") #finds locations of .
            s4 = s1[s2:s3] #separates filename from directory and removes the extension and period
            
            final = d + s4 + ext #makes directory string for saving the new file
            
            if obj[fts] == "img": 
                im = Image.open(r''+s1+'')
                im.save(r''+final+'')
            
            #if obj[fts] == "txt": 
            #    im = Image.open(r''+s1+'')
                #im.save(r''+final+'')


def makeset(num):
    global total_sets
    f = "frame" + num
    a = "allowsectdel" + num
    fts = "fts" + num
    total_sets = int(num)
    yposition = "ypos" + num
    yposit = int(obj[yposition])
    yposdict[yposition] = yposit
    framedict[f] = Frame(root, bg = '#00acdb')
    framedict[f].place(x= 5, y= yposit, width = (wW-10), height = 40)
    makexbutton(a, framedict[f])
    makeftswitch(fts, num, framedict[f])
    makefbbtn(fts, num, framedict[f])
    makefinalfiletypechooser(fts, num, framedict[f])
    makedesttypeswitch(num, framedict[f])
    makedestchooser(num, framedict[f])
    makeconverter(num, framedict[f])


def addnewset():
    pass

with open('fileconverterdata.json') as f:
    obj = json.load(f)
    
    total_sets = 0
    for s in obj:
        if s.startswith("set"):
            total_sets = total_sets+1
            num = obj[s]
            makeset(num)


aboveframe = Frame(root, bg = 'red')
aboveframe.place(x= 5, y= 2, width = (wW-10), height = 140)

addbtn = tk.Button(aboveframe, text = "Add set", height = 2, width = 15)
addbtn.grid(row = 0, column = 1, padx = xpad, pady=ypad)


root.mainloop()
