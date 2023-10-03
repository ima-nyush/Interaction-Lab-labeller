# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
import tkinter as tk
from tkinter import filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Frame, Paragraph, KeepInFrame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import utils

fileName = ''
saveLoc = ''

def readExcel(wb):
    wb = load_workbook(wb).active
    createPrintPDF(wb)

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

def createPrintPDF(wb):
    imageX = 1
    imageY = 24.95
    titleX = 4.8
    titleY = 24.95+1.3
    subtitleX = 4.8
    subtitleY = 24.95
    #qrCodeX = 16.2
    #qrCodeY = 24.95
    
    imgs = extractImg(wb)
    c = canvas.Canvas('{0}{1}labels.pdf'.format(saveLoc,'/'),pagesize=A4)
    
    for row in range(2,countRow(wb)+2):
        if ((row-2)%6==0 or row==countRow(wb)+1):
            if (row!=2):
                c.showPage()
                if (row==countRow(wb)+1):
                    c.save()
                    break
                imageX = 1
                imageY = 24.95
                titleX = 4.8
                titleY = 24.95+1.3
                subtitleX = 4.8
                subtitleY = 24.95
                #qrCodeX = 16.2
                #qrCodeY = 24.95
        #img
        if imgs[row-2]!=None:
            c.drawImage(utils.ImageReader(imgs[row-2]),x=imageX*cm, y=imageY*cm,width=3.8*cm,height=3.8*cm)
        imageY-=4.8
        
        #title
        titleFrame = Frame(titleX*cm, titleY*cm,11.4*cm, 2.5*cm, showBoundary=0)
        titleStyle = ParagraphStyle('title', fontName='Helvetica', fontSize=70, alignment=1, wordWrap=None, leading=75)
        if wb.cell(row=row,column=2).value == None:
            titleTxt = ''
        else:
            titleTxt = [Paragraph(str(wb.cell(row=row,column=2).value),titleStyle)]
        titleTxtC = KeepInFrame(11.4*cm, 2*cm, titleTxt, mode='shrink', vAlign='MIDDLE', hAlign='CENTER', fakeWidth=False)
        titleFrame.addFromList([titleTxtC], c)
        titleY-=4.8
        
        
        #subtitle
        subtitleFrame = Frame(subtitleX*cm,subtitleY*cm,11.4*cm,1.3*cm,showBoundary=0)
        subtitleStyle = ParagraphStyle('subtitle', fontName='Helvetica', fontSize=36, alignment=1, wordWrap=None, leading=40)
        if wb.cell(row=row,column=3).value == None:
            subtitleTxt = ''
        else:
            subtitleTxt = [Paragraph(str(wb.cell(row=row,column=3).value),subtitleStyle)]
        subtitleTxtC = KeepInFrame(11.4*cm, 1.3*cm, subtitleTxt, mode='shrink', vAlign='MIDDLE', hAlign='CENTER', fakeWidth=False)
        subtitleFrame.addFromList([subtitleTxtC], c)
        subtitleY-=4.8
        
        #qrcode
        #qrCodeFrame = Frame(qrCodeX*cm,qrCodeY*cm,3.8*cm,3.8*cm,showBoundary=1)
        #qrCodeY-=4.8

def loadGUI():
    window = tk.Tk()
    window.geometry("600x115")
    window.resizable(False,False)
    window.title("IXL Label PDF Generator")
    
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
            el["text"]="Missing sheet and/or save location"
        else:
            readExcel(fileName)
    
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
