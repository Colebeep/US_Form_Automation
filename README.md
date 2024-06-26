# US_Form_Automation
Poggio Civitate Archaeological Project, made in 2023
Author: Cole Adam Reilly 

The Problem:
    Trench Supervisors need to fill out two sets of forms when they close a Locus (The stratigraphic unit for this site). One is hosted on Kobotoolbox and integrates with our database OpenContext.org (called the kobo form). The other is a governmental form required by the soprintendente of archaeology for our region (called the US form).
    The kobo form works much like a google forms survey, is relatively user friendly and made by us using our sites language/terminology.
    The US form is a standard for the region, is written in italian, and made to accommodate the varing language of every site, it can be rather confusing to educate our excavators on the nuances of the form.
    The workflow for filling US forms is significantly slower, requiring you to duplicate all your information in distinct word files, one for each locus. This resulted in multiple hours of extra labor for each trench, with some have upwards of 30 distinct loci in any given trench, resulting in 30 kobo forms, and 30 distinct US Forms.

## The Solution:
Kate Kriendler (head of Excavation) approached myself and Eric Kansa (co-creator of our database) to find a way to automate the population of the US form. 
Trench supervisors should fill out the kobo forms, and at the end of the season run a program to grabbed all the information and populate it into .docx files.
    
To do this several things needed to happen:
        1. The kobo form was expanded to account for information only asked for in the US form. This was done by Eric and Kate.
        2. A US Form word document needed to be populated with mailmerge fields to allow for population. This was done by me, with the helpf of a guide to the US Form written by Ann Glennie (catalogour) and Kate.
        3. A python file needed to be written to read the xlsx output of the kobo form and populate the template US Form. This was done by me.

## How to run the program?
currently the solution is written as a .py file and requires the correct files/folder system as well as some presinstalled libraries. I might create a .ipynb file to better the system.
run the file through cmd (terminal for mac) or through an IDE like VSCode open to the correct folder. 
        python compile_forms.py

The file needs a couple things to run correctly, two libraries that must be installed beforehand, and two files that need to be named correctly and put in the correct space. Those are outlined below:

 ### Python Requirements: 
  Download Python3 and the following libraries
        mailmerge
        pandas
    Should be installed using Pip, instructions to that can be found on the web.
    html.parser, os, and stringIO are all apart of the standard library and do not need to be installed.

All locus forms should be downloaded from the kobo website (need official Log-in): 
        https://kform.opencontext.org/#/forms
    navigate to the Locus Summary Entry form for this year
      click the data tab
         click the downloads tab
            select export type as xls, and 'XML values and headers'
            click export
            in exports section click download on generated export
            rename file to 'kobo_input_all.xlsx' and move the file to the same folder as compile_forms.py

The opening trench coordinates need to be compiled into a CSV called 'Trench_Coordinates.csv'.
        Each column header should be the shortened designation for the trench. So the trench Tesoro 102 would get shortened to T102, the trench Civitate B 25 would be CB25, and so on.
        Each row should be the coordinates in the local master grid (NOT WGS84 or ESPG3003) The trench should start with the NW corner and move clockwise around the trench. so a standard rectangular trench would be NW NE SE SW.
        save that file and place it in the same folder as compile_forms.py



## What is compile_forms.py?
This is the main python file that solves the above problem. Read in the kobo form xlsx. This is a excel file (workbook) that has multiple internal sheets used for organization. Each sheet gets read in as a pandas dataframe. Every locus is processed, will all appropriate fields being found through the excel sheet. Each section is formatted correctly for the form. Then through the use of mailmerge, a new form is made, populated, and saved in the appropriate spot.

## What is Mail Merge?
Mail Merge is a system that allows for populating copies of documents on mass (https://en.wikipedia.org/wiki/Mail_merge)

## What is Trench_Coordinates.csv?
The only piece of information needed for the US form that is not stored in the kobo form are the coordinates for each trench. These coordinates need to be in the master grid for Poggio Civitate, and NOT in a known CRS like WGS84 or ESPG3003 (The CRS for our GIS)
    The format of the csv should be as follows:
        Each column header should be the shortened designation for the trench. So the trench Tesoro 102 would get shortened to T102, the trench Civitate B 25 would be CB25, and so on.
       More information in the FURTHER_INFO.txt file. Each row should be the coordinates in the local master grid (NOT WGS84 or ESPG3003) The trench should start with the NW corner and move clockwise around the trench. so a standard rectangular trench would be NW NE SE SW.
        save that file and place it in the same folder as compile_forms.py
    in practice the csv could look something like this (example comes from 2023):

#What is kobo_input_all.xlsx?
    This is all the data. The specifics of which can be read in the the FURTHER_INFO.txt file.
