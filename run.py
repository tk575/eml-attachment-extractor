#-----------------------------------------------------------------------
# Extracts attachments from a folder containing .eml files
# Author        : Jayadeep Karnati
# Created on    : 27 April 2015
# Last Modified : 27 April 2015
# Dependencies  : patool
#-----------------------------------------------------------------------

import os
import shutil
import base64
import zipfile
import patoolib

# Get list of all files
files = [f for f in os.listdir('.') if os.path.isfile(f)]
# Create output directory
if os.path.exists("output"):
    shutil.rmtree("output")
os.makedirs("output")

for eml_file in files:
    if eml_file.endswith(".eml"):
        with open(eml_file) as f:
            email = f.read()
        exts = ['.docx', '.zip', '.pdf', '.rar', '.tar.gz', '.pptx']
        ext=""
        for e in exts:
            if e in email:
                ext = e
        if ext is not "":
            # Extract the base64 encoding part of the eml file
            encoding = email.split(ext+'"')[-1]
            if encoding:
                # Remove all whitespaces
                encoding = "".join(encoding.strip().split())
                encoding = encoding.split("=", 1)[0]
                # Convert base64 to string
                if len(encoding) % 4 != 0: #check if multiple of 4
                   while len(encoding) % 4 != 0:
                       encoding = encoding + "="
                try:
                    decoded = base64.b64decode(encoding)
                except:
                    print(encoding)
                    for i in range(100):
                        print('\n')
                # Save it as docx
                path = os.path.splitext(eml_file)[0]
                if path:
                    path = os.path.join("output", path + ext)
                    try:
                        os.remove(path)
                    except OSError:
                        pass
                    with open(path, "wb") as f:
                        f.write(decoded)
        else:
            print("File not done: " + eml_file)

# Get all useful files inside zip
files = [f for f in os.listdir(os.path.join("output"))]
exts = ['.docx', '.pdf', '.pptx']
for zip_file in os.listdir("output"):
    if zip_file.endswith(".zip") or zip_file.endswith(".rar"):
        if os.path.exists("zip"):
            shutil.rmtree("zip")
        os.makedirs("zip")
        patoolib.extract_archive(os.path.join("output", zip_file),
                                 outdir="zip", verbosity=0)
        for root, dirs, files in os.walk("zip"):
            for f in files:
                for ext in exts:
                    if f.endswith(ext):
                        shutil.copy(
                            os.path.join(root, f),
                            os.path.abspath(os.path.join(root, "..", "output"))
                        )

if os.path.exists("zip"):
    shutil.rmtree("zip")