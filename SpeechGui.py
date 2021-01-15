from tkinter import *
from tkinter import filedialog
import speech_recognition as sr
import os
import docx

def askFile():
    def conv(): 
        text="t"
        print(filePath)
        status.config(text="Status: Busy")
        convBtn['state'] = DISABLED
        gui.update()
        recognizer = sr.Recognizer()
        
        extension=filePath[ -3:]
        
        if(extension=="wav"):
        
            with sr.AudioFile(filePath) as source:
                recorded_audio = recognizer.record(source)
            
            #''' Recorgnizing the Audio '''
            try:
                text = recognizer.recognize_google(
                        recorded_audio, 
                        language="en-UK"
                    )
            
            except Exception as ex:    
                text="777"
        
            textName=textInput.get()
            save_path=os.path.join(os.environ["HOMEPATH"], "Desktop")
            filePathSave=os.path.join(save_path, textName+".docx")   
            
            my_doc = docx.Document()
            my_doc.add_paragraph(text)
            my_doc.save(filePathSave)
            #f= open(filePathSave,"w+")
            #f.write(text)
            #f.close()
            
            status.config(text="Status: Finished")
            
        else:
            status.config(text="Status: Error, select a .wav file") 
            convBtn['state'] = DISABLED
        
        
    gui.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file")
    filePath=gui.filename
    
    textOfFile="Loaded File: "+filePath
    labelFile.config(text=textOfFile)
    
    
    if(convBtn['state'] == DISABLED):
        convBtn['command']= conv
        convBtn['state'] = NORMAL
        
        status.config(text="Status: Waiting for input")
     
text=""
gui = Tk(className=' Audio To Text')

gui.geometry("800x400")
gui.configure(bg='#3eab60')

myLabel = Label(gui, text="Enter Text File Name:",font=("Courier", 20),bg='#3eab60')
myLabel.pack()

textInput=Entry(gui, width=40,font=("Courier", 15))
textInput.pack()

makespace = Label(gui, text=" ",bg='#3eab60')
makespace.pack()

textNameButton=Button(gui,text="Select .wav File",font=("Courier", 15),command=askFile)
textNameButton.pack()

labelFile = Label(gui, text="Loaded: None",font=("Courier", 10),bg='#3eab60')
labelFile.pack()

makespace = Label(gui, text=" ",bg='#3eab60')
makespace.pack()

convBtn=Button(gui,text="Convert to Text",font=("Courier", 15),state=DISABLED)  
convBtn.pack()

status = Label(gui, text="Status: Waiting for input",font=("Courier", 15),bg='#3eab60')
status.pack()

gui.mainloop()