# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Frame, Paragraph, KeepInFrame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import utils
import traceback
import csv
import os

fileName = ''
saveLoc = ''
sheetNo = 0

def readEx(sh):
    wb = load_workbook(sh)
    wbs = wb.worksheets[sheetNo]
    createPrintPDF(wbs)
    return(wbs)

def extractImg(wb):
    imgLoad = SheetImageLoader(wb)
    sImg = []
    for i in range(1,countRow(wb)):
        if imgLoad.image_in('A{0}'.format(i)):
            sImg.append(imgLoad.get('A{0}'.format(i)))
        else:
            sImg.append(None)
    return sImg

def countRow(ws):
    rc = 0
    for row in ws:
        if any(cell.value is not None for cell in row):
            rc+=1
    return rc
    

def createPrintPDF(wbk):
    imgx = 1
    imgy = 24.95
    titx = 4.8
    tity = 24.95+1.3
    subx = 4.8
    suby = 24.95
    qrcx = 16.2
    qrcy = 24.95
    
    imgs = extractImg(wbk)
    c = canvas.Canvas('{0}{1}labels.pdf'.format(saveLoc,'/'),pagesize=A4)
    for crow in range(2,countRow(wbk)+2):
        if ((crow-2)%6==0 or crow==countRow(wbk)+1):
            if (crow!=2):
                c.showPage()
                if (crow==countRow(wbk)+1):
                    c.save()
                imgx = 1
                imgy = 24.95
                titx = 4.8
                tity = 24.95+1.3
                subx = 4.8
                suby = 24.95
                qrcx = 16.2
                qrcy = 24.95
                if (crow >= countRow(wbk)+1):
                    break;
        #img
        #fImg = Frame(imgx*cm,imgy*cm,3.8*cm,3.8*cm,showBoundary=1)
        c.drawImage(utils.ImageReader(imgs[crow-2]),x=imgx*cm, y=imgy*cm,width=3.8*cm,height=3.8*cm)
        #fImg.addFromList(dimg, c)
        imgy-=4.8
        
        #title
        fTit = Frame(titx*cm, tity*cm,11.4*cm, 2.5*cm, showBoundary=0)
        titStyle = ParagraphStyle('title', fontName='Helvetica', fontSize=70, alignment=1, wordWrap=None, leading=75)
        if wbk.cell(row=crow,column=2).value == None:
            dtittemp = ''
        else:
            dtittemp = [Paragraph(str(wbk.cell(row=crow,column=2).value),titStyle)]
        dtit = KeepInFrame(11.4*cm, 2*cm, dtittemp, mode='shrink', vAlign='MIDDLE', hAlign='CENTER', fakeWidth=False)
        fTit.addFromList([dtit], c)
        tity-=4.8
        
        
        #subtitle
        fSub = Frame(subx*cm,suby*cm,11.4*cm,1.3*cm,showBoundary=0)
        subStyle = ParagraphStyle('subtitle', fontName='Helvetica', fontSize=36, alignment=1, wordWrap=None, leading=40)
        if wbk.cell(row=crow,column=3).value == None:
            dsubtemp = ''
        else:
            dsubtemp = [Paragraph(str(wbk.cell(row=crow,column=3).value),subStyle)]
        dsub = KeepInFrame(11.4*cm, 1.3*cm, dsubtemp, mode='shrink', vAlign='MIDDLE', hAlign='CENTER', fakeWidth=False)
        fSub.addFromList([dsub], c)
        suby-=4.8
        
        #qrcode
        fQrc = Frame(qrcx*cm,qrcy*cm,3.8*cm,3.8*cm,showBoundary=1)
        qrcy-=4.8

def loadGUI():
    window = tk.Tk()
    window.geometry("600x115")
    window.resizable(False,False)
    window.title("Kevin's Super Duper Ultimate Label Printer Series 5000")
    
    def fileUp(event = None):
        fn = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx")])
        global fileName 
        fileName = fn
        fteT.set(fileName)
        fte["text"] = fteT.get()
        
    def saveUp(event = None):
        sl = filedialog.askdirectory()
        global saveLoc
        saveLoc = sl
        steT.set(saveLoc)
        dle["text"] = steT.get()
    
    def preRunCheck(event = None):
        if (fteT.get() == '' or steT.get()==''):
            el["text"]="After a not-so-comprehensive internal investigation we've determined this incident to be user error"
        else:
            mS = readEx(fileName)
    
    ft = tk.Label(text="xlsx File:")
    ft.place(width=50,height=25,x=10,y=10)
    
    fteT = tk.StringVar(window, '')
    fte = tk.Label(window, text=fteT.get())
    fte.place(width=400,height=25,x=70,y=10)
    
    fbut = tk.Button(window, text="Open", command=fileUp)
    fbut.place(width=50,height=25,x=480,y=10)
    
    rbut = tk.Button(window, text="Run", command=preRunCheck)
    rbut.place(width=50,height=25,x=540,y=10)
    
    dl = tk.Label(text="Save to:")
    dl.place(width=50,height=25,x=10,y=45)
    
    steT = tk.StringVar(window, '')
    dle = tk.Label(window,text=steT.get())
    dle.place(width=400,height=25,x=70,y=45)
    
    dbut = tk.Button(window,text="Open", command=saveUp)
    dbut.place(width=50,height=25,x=480,y=45)
    
    el = tk.Label(window,text='',wraplength=440)
    el.place(width=440,height=25,x=10,y=80)
    
    window.mainloop()

if __name__ == '__main__':
    loadGUI()
