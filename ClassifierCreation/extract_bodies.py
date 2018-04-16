'''
This file extracts and clean the body text from the
corpus email files.
'''
__author__ = 'Kurtis Kuszmaul'

import email.parser 
import os, sys, stat
import shutil
from HTMLParser import HTMLParser

class MLStripper(HTMLParser):
    ''' Class for stripping HTML from text '''

    def __init__(self):
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    ''' Strips HTML tags from text '''

    s = MLStripper()
    s.feed(html)
    return s.get_data()

def chunkstring(string, length):
    ''' Chunks texts into strings of specified length '''

    return (string[0+i:length+i] for i in range(0, len(string), length))

def process_text(body, dstfile):
    ''' 
    Removes HTML tags, whitespace characters and divides bodies
    into appropriately-sized chunks.
    '''

    cleaned_text = strip_tags(body)
    cleaned_text = cleaned_text.replace('\n','')
    cleaned_text = cleaned_text.replace('\t','')
    cleaned_text = cleaned_text.replace('\r','')
    
    for chunk in chunkstring(cleaned_text, 1000):
        stripped = chunk.lstrip()
        stripped = chunk.rstrip()
        dstfile.write(stripped + '\n')

    return cleaned_text

#############################################################
# Code below primarily taken from 
# http://www.csmining.org/index.php/spam-email-datasets-.html
#############################################################

def ExtractSubPayload (filename):
    '''
    Extract the body from the .eml file.
    '''
    if not os.path.exists(filename):
        print("ERROR: input file does not exist:", filename)
        os.exit(1)
    fp = open(filename)
    msg = email.message_from_file(fp)
    payload = msg.get_payload()
    if type(payload) == type(list()) :
        payload = payload[0]
    if type(payload) != type('') :
        payload = str(payload)
    
    return payload

def ExtractBodyFromDir ( srcdir, dstdir ):
    '''
    Extract the body information from all .eml files in the srcdir and 
    save the file to the dstdir with the same name.
    '''
    if not os.path.exists(dstdir): # dest path doesnot exist
        os.makedirs(dstdir)  
    files = os.listdir(srcdir)
    for i,file in enumerate(files):
        srcpath = os.path.join(srcdir, file)
        dstpath = os.path.join(dstdir, file)
        src_info = os.stat(srcpath)
        if stat.S_ISDIR(src_info.st_mode): # for subfolders, recurse
            ExtractBodyFromDir(srcpath, dstpath)
        else:
            body = ExtractSubPayload (srcpath)
            dstfile = open(dstpath, 'w')
            body = process_text(body, dstfile)
            dstfile.close()

print('Input source directory: ') #ask for source and dest dirs
srcdir = raw_input()
if not os.path.exists(srcdir):
    print 'The source directory %s does not exist, exit...' % (srcdir)
    sys.exit()
# dstdir is the directory where the content .eml are stored
print 'Input destination directory: ' #ask for source and dest dirs
dstdir = raw_input()
if not os.path.exists(dstdir):
    print 'The destination directory is newly created.'
    os.makedirs(dstdir)

ExtractBodyFromDir ( srcdir, dstdir )
