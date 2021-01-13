#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# It is easier to process .docx files so for any .doc file, 
# we create a .docx for it and delete the .docx after we get the file contents.
import os
from docx2python import docx2python
from docx2python.iterators import iter_paragraphs
import pandas as pd
from glob import glob
import re
import win32com.client as win32
from win32com.client import constants

# method gets a doc filename and creates a docx then returns the doc
def change_to_docx(path,file):
    filename = path+file
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(filename)
    doc.Activate()
    new_file_abs = os.path.abspath(filename)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)
    file = file+"x"
    return file, True



path = r"C:\\PATH\\TO\\FILES\\"
for file in os.listdir(path):
    if !file.endswith(".doc") and !file.endswith(".docx"):
        continue
    
    newFileCreated = False
    
    if file.endswith(".doc"):
        file, newFileCreated = save_as_docx(path,file)
    
    filepath = path+file           
    content = docx2python(filepath)           
    lines = list(iter_paragraphs(content.document))
    
    # if a new docx file was created, you want to delete it after getting the contents out
    if newFileCreated == True:
        os.remove(file)
    
    '''
    Do the file processing here
    '''

