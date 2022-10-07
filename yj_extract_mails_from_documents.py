#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May 10 22:17:11 2022
by Yves Jeanrenaud
Change log
---
Version 1.0.1 typos and comments
Verson 1.0 runs on PDF files
Version 1.2 now also reads doc/x files
This script scanns files in doc, docx and PDf format (unprotected) for e-mail addresses. They may also be camouflaged e. g. mail_at_somehwere.tld
currently only @, (at), [at] and _at_ are supported.

with thanks, regards and cudos to:
https://www.nicholasnadeau.com/post/2021/9/pdf-hero-how-to-extract-emails-with-python/
https://www.geeksforgeeks.org/append-list-of-dictionary-and-series-to-a-existing-pandas-dataframe-in-python/

@author: yjeanrenaud
https://github.com/yjeanrenaud/YJ_extract_mails_from_documents/
"""

import logging
import fitz  # this is pymupdf
from docx import Document

from pathlib import Path
from typing import Optional
import re


def _get_logger(name: str, level: str = "INFO"):
    # get named logger and set level
    logger = logging.getLogger(name)
    logger.setLevel(level=level)

    # set output channel and formatting
    ch = logging.StreamHandler()
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    return logger

LOGGER = _get_logger(__name__)


def parse_pdf(path: Path) -> dict:
    #LOGGER.info(f"Parsing {path.relative_to(ROOT_DIR)}")

    # prepare data to be returned
    data = {
        #"person": path.parent.name,
        #"category": str(path.relative_to(ROOT_DIR).parts[0]),
        "file": path.name,
        "path": str(path.relative_to(ROOT_DIR)),
    }
    
    
    # get email separately due to possible errors
    try:
        data["email"] = extract_email_pdf(path)
    except TypeError as e:
        LOGGER.error(f"Failed to open {path.name}: {e}")
    except Exception as e:
        # don't want misc errors crashing the entire script
        # better to have a few blank emails
        LOGGER.error(f"Failed to parse {path.name}: {e}")

    return data

def parse_docx(path: Path) -> dict:
    #LOGGER.info(f"Parsing {path.relative_to(ROOT_DIR)}")

    # prepare data to be returned
    data = {
        #"person": path.parent.name,
        #"category": str(path.relative_to(ROOT_DIR).parts[0]),
        "file": path.name,
        "path": str(path.relative_to(ROOT_DIR)),
    }
    
    
    # get email separately due to possible errors
    try:
        data["email"] = extract_email_docx(path)
    except TypeError as e:
        LOGGER.error(f"Failed to open {path.name}: {e}")
    except Exception as e:
        # don't want misc errors crashing the entire script
        # better to have a few blank emails
        LOGGER.error(f"Failed to parse {path.name}: {e}")

    return data

def validate_email_string(text: str) -> Optional[str]:
    # humans don't always follow instructions
    # this can cause problems with text and fields
    # be super strict when getting emails
    #email_regex = re.compile(r"\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b") # just mail@somehwere.tld4
    email_regex = re.compile(r"\b[\w\.-]+(@|_at_|\[at\]|\(at\))[\w\.-]+\.\w{2,}\b") # @ or _at_  or [at] or (at) and also include longer tld e. g. .architect
    #print(text)
    results = re.findall(email_regex, text)
    if results.count==0:
        return results[0]
    else:
        strReturn = results.pop(0)
        for strI in results:
            strReturn += ";" + strI
        return strReturn


def get_email_from_pages(reader) -> Optional[str]:
    # some PDFs had the email elsewhere in the document
    # need to iterate through the pages
    
    text = ""
    for page in reader:
        text += page.get_text()
    #print(text)
    email = validate_email_string(text)
    if email:
        return email

def extract_email_pdf(path: Path) -> Optional[str]:
    # PDFs are binary
    # need to use the "read binary" (`rb`) flag
    with fitz.open(path) as doc:

        email = get_email_from_pages(doc)

        if email:
            LOGGER.info(f"Email found in {path.relative_to(ROOT_DIR)}: {email}")
            return email
        #else:
            #LOGGER.info(f"No mail found in {path.relative_to(ROOT_DIR)}.")

def extract_email_docx(path: Path) -> Optional[str]:
    document = Document(path)
    for para in document.paragraphs:
        #print(para.text)
        email =  validate_email_string(para.text)

    if email:
        LOGGER.info(f"Email found in {path.relative_to(ROOT_DIR)}: {email}")
        return email
    #else:
        #LOGGER.info(f"No mail found in {path.relative_to(ROOT_DIR)}.")



#ROOT_DIR = r"."
ROOT_DIR = r"..\PUT_YOUR_DOCUMENTS_HERE"
ROOT_DIR = Path(ROOT_DIR)

glob_pattern = "*.pdf"
pdf_paths = ROOT_DIR.rglob(glob_pattern)

import pandas as pd
df = pd.DataFrame()#({"header":[1,2]})

for path in pdf_paths:
    extracted_data  = parse_pdf (path)
    df = df.append(extracted_data, ignore_index=True, sort=False)
	
glob_pattern = "*.doc*"
docx_paths = ROOT_DIR.rglob(glob_pattern)

for path in docx_paths:
    extracted_data  = parse_docx (path)
    df = df.append(extracted_data, ignore_index=True, sort=False)

df.to_csv("emails_collected.csv",sep=';',encoding='ansi') # change ; to , if you are not using a stupid office suite from Redmond in a German or Swiss-German localisation that assumes semicola as separators
