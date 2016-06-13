# -*- coding: utf-8 -*-
"""
Created on Sat Feb 13 11:55:01 2016

@author: Jon
"""

import re

from bs4 import BeautifulSoup
from difflib import SequenceMatcher
import pandas as pd

import bs4
import tempfile


from urllib.request import urlopen
import urllib
import os
import shutil
import subprocess


def extract_address_phone(link):

    if link == None:
        return (None, None)

    # Get the HTML soup
    response = urlopen(link)
    soup = BeautifulSoup(response.read(), "html5")

    # sections[0] : Mailing/Head Office address
    # sections[3] : Telephone, Reporting Jurisdictions
    sections = soup.find_all('tr', {'valign':'TOP'})

    # Get the address and only keep strings
    address_results = sections[0].find_all('td', {'class':'rt'})
    address_results = address_results[0].contents
    temp = [s for s in address_results if type(s) == bs4.element.NavigableString]

    address = '\r'.join(temp).strip()

    phone_results = sections[3].find_all('td', {'class':'rt'})
    phone_results = phone_results[0].contents[0].replace(' ', '-')

    return (address, phone_results)

def get_profile_link(company_name):

    # Get the first letter
    # Extract first letter as lowercase
    # IF first letter is a digit, use 'nc'
    first = company_name[0].lower()
    if first.isdigit():
        first = 'nc'

    # Go through the dictionary, access URL
    profile_url = 'http://sedar.com/issuers/company_issuers_' + first + '_en.htm'

    # Access the URL
    response = urlopen(profile_url)
    soup = BeautifulSoup(response.read(), "html5")

    # Find all lines containing lists, hopefully this will contain a hyperlink
    lines = soup.find_all('li' , {'class':'rt'})

    # Go through each list item
    for line in lines:

        # Find link
        link = line.find_all('a')
        link_name = link[0].contents[0]

        # Remove parenthesis and contents
        link_name = re.sub(r'\([^)]*\)', '', link_name).strip()

        # Remove commas
        link_name = link_name.replace(',', '').strip()

        # Remove trailing whitespace
        link_name = link_name.strip()

        # Match
        m = SequenceMatcher(None, link_name, company_name)

        if m.ratio() > 0.95:
            hyperlink = 'http://sedar.com' + link[0]['href']
            return hyperlink

def get_name_company(path_to_sedi_html):

    db = {}

    f = open(path_to_sedi_html, "r")

    for line in f:

        # Get Issuer name
        if "Issuer:" in line and "Insider's Relationship" not in line:
            company = line

            # Remove HTML
            company = re.sub(r"<.*?>", "", company)

            # Remove "Issuer:
            company = re.sub(r"Issuer:", "", company)

            # Replace &amp; with &
            company = re.sub("&amp;", "&", company)

            # Remove parenthesis and contents
            company = re.sub(r"\(.*?\)", "", company).strip()

        if "Insider:" in line and "," in line:
            name = line

            # Remove HTML
            name = re.sub(r"<.*?>", "", name)

            # Remove "Insider:"
            name = re.sub("Insider:", "", name).strip()

            # Rearrange name
            name = name.split(",")

            temp = ""
            if len(name) == 2:
                temp = name[1].strip() + " " + name[0].strip()
            else:
                temp = name[1].strip() + " " + name[2].strip() + " "  + name[0].strip()

            name = temp

        if "Grant of options" in line:
            option = line
            option = re.sub(r"<.*?>", "", option).strip()

            if company in db.keys():
                db[company].add(name)
            else:
                db[company] = set([name])

    f.close()

    # Convert names to lists
    for company in db:
        db[company] = list(db[company])

    return db

def get_address(db_name):

    db_address = {}

    for company_name in db_name:

        # Get SEDAR URL
        company_url = get_profile_link(company_name)

        # Get address and phone number
        address, phone = extract_address_phone(company_url)

        if address != None:
            address = address.replace("\r", "\r\n")

        # Add to new dictionary
        try:
            db_address[company_name] = (address.strip(), phone.strip())
        except Exception:
            db_address[company_name] = ("","")

    return db_address

def get_date(path_to_sedi_html):

    # Get date
    f = open(path_to_sedi_html, "r")

    date = None
    for line in f:

        if date == None:
            result = re.findall(r"[\d]{4}\-[\d]{2}\-[\d]{2}", line)

            if len(result) > 0:
                f.close()
                return result[0]
    f.close()

    if date == None:
        return "YYYY-MM-DD"

def download_pdf(temp_dir):
    request = urllib.request.Request("https://www.sedi.ca/sedi/SVTWeeklySummaryACL?name=W1ALLPDFI&locale=en_CA")
    request.add_header("Referer", "https://www.sedi.ca/sedi/SVTReportsAccessController?menukey=15.03.00&locale=en_CA")

    # Path to the PDF file with filename.
    path_to_pdf = os.path.join(temp_dir, "SVTWeeklySummaryACL.pdf")

    with urllib.request.urlopen(request) as response:
        f = open(path_to_pdf, "wb")
        f.write(response.read())
        f.close()

    # Return the full path to the PDF.
    return path_to_pdf

def setup():
    # Create Temporary directory.
    temp_dir = tempfile.mkdtemp()

    # Create paths
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # Path to the pdftohtml binary.
    filepath_pdftohtml = os.path.join(script_dir, "pdftohtml", "pdftohtml.exe")

    # Copy pdftohtml.exe to Temporary directory.
    filepath_pdftohtml = shutil.copyfile(filepath_pdftohtml, os.path.join(temp_dir, "pdftohtml.exe"))

    # Download the SEDI PDF.
    filepath_pdf = download_pdf(temp_dir)

    # Execute pdftohtml.exe with the PDF file (/path/to/pdftohtml.exe /path/to/SVTWeeklySummary.pdf).
    subprocess.call([filepath_pdftohtml, filepath_pdf])

    # Get the path to the HTML file which was created by pdftohtml.exe.
    filepath_html = os.path.join(temp_dir, "SVTWeeklySummaryACLs.html")

    return (temp_dir, filepath_html)

# Setup procedure.
#
# Get the temporary directory used, and the path to the HTML file (converted PDF).
print("Downloading SVTWeeklySummaryACL.pdf.")
temp_dir, filepath_html = setup()

# Get the date of the summary
print("Getting the date the Summary file was generated on.")
date = get_date(filepath_html)

# Get issuer and insider
print("Getting list of Issuers and Insiders")
db_name = get_name_company(filepath_html)

# Get address
print("Getting addresses.")
db_address = get_address(db_name)

# Build the dataframe
print("Building the table.")
df = pd.DataFrame(columns=["Name", "Company", "Address", "Phone Number", "Title (if any)"])

# Go through each company name, lookup address, check if in ontario.
counter = 0
for company_name in db_address:

    # Look up address, phone
    address, phone = db_address[company_name]

    # check if in ontario,
    if address == None or "ON" in address or ", ON" in address or "Ontario" in address or "ONTARIO" in address:

        for name in db_name[company_name]:

            df.loc[counter] = [name, company_name, address, phone, ""]
            counter = counter + 1

df.sort_values(by=["Company"], ascending=True, inplace=True)
df.set_index(keys=["Name"], inplace=True)
df.fillna(value="", inplace=True)
#df.to_csv("New Stock Options (" + date + ").csv", encode="utf-8")

# Write to .XLS file.
filename = "New Stock Options (" + date + ").xls"
print("Writing to: " + filename)
writer = pd.ExcelWriter(filename)
df.to_excel(writer, "Sheet1")
writer.save()

# Teardown: delete temporary directory
shutil.rmtree(temp_dir)
