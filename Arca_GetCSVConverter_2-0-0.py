import glob
import datetime
import os
import pandas as pd
import numpy as np
import re
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import messagebox


# pyinstaller --onefile --noconsole --icon GetCSV.ico Arca_GetCSVConverter_2-0-0.py
#for MMW 18-6 spreadsheets

probCol = False
#infer desktop
desktopPath = os.path.expanduser("~/Desktop/")
filelist=['']
probRecords = []
probColls = []
#filename = r'arms_modsonly_May9.csv'

col_names =  ["IslandoraContentModel","BCRDHSimpleObjectPID",'imageLink','filename','directory','childKey','title', 'alternativeTitle', 'creator1', 'creator2','creator3']
col_names += ['corporateCreator1','corporateCreator2','contributor1','contributor2','corporateContributor1','publisher_original','publisher_location']
col_names += ['dateCreated','description','extent','topicalSubject1','topicalSubject2','topicalSubject3','topicalSubject4','topicalSubject5']
col_names += ['geographicSubject1','coordinates','personalSubject1','personalSubject2','corporateSubject1','corporateSubject2', 'dateIssued_start']
col_names += ['dateIssued_end','dateRange', 'frequency','genre','genreAuthority','type','internetMediaType','language1','language2','notes']
col_names += ['accessIdentifier','localIdentifier','ISBN','classification','URI']
col_names += ['source','rights','creativeCommons_URI','rightsStatement_URI','relatedItem_title','relatedItem_PID','recordCreationDate','recordOrigin']

pattern1 = r'^[A-Z][a-z]{2}-\d{2}$' #%b-%Y date (e.g. Jun-17) 
pattern2 = r'^\d{2}-\d{2}-[1-2]\d{3}$'

contentModels = {
  r"info:fedora/islandora:sp_large_image_cmodel": "Large Image",
  r"info:fedora/islandora:sp_basic_image": "Basic Image",
  r"info:fedora/islandora:bookCModel": "Book",
  r"info:fedora/islandora:newspaperIssueCModel":"Newspaper - issue",
  r"info:fedora/islandora:newspaperPageCModel":"Newspaper",
  r"info:fedora/islandora:sp_PDF":"PDF",
  r"info:fedora/islandora:sp-audioCModel":"Audio",
  r"info:fedora/islandora:sp_videoCModel":"Video",
  r"info:fedora/islandora:sp_compoundCModel":"Compound",
  r"info:fedora/ir:citationCModel":"Citation"
}

def browse_button():  
    # Allow user to select a directory and store it in global var
    # called folder_path1
    lbl1['text'] = ""
    csvname =  filedialog.askopenfilename(initialdir = desktopPath,title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
    if ".csv" not in csvname:
        lbl1['text'] = "**Please choose a file with a .csv extension!"
    else:
        filelist[0] = csvname
        lbl1['text'] = csvname

def splitMultiHdgs(hdgs):
    if pd.notna(hdgs):
        hdgs = hdgs.replace("\\,",";")
        hdgs = hdgs.split(",")
        newhdgs = []
        for hdg in hdgs:
            newhdg = hdg.replace(";", ",")
            newhdgs.append(newhdg)
        return newhdgs
    else:
        return None
    
def getMultiVals(item, string, df, pd):
    hdgs = df.filter(like=string).columns
    for hdg in hdgs:
        vals = df.at[item.Index,hdg]
        if pd.notna(vals):
            vals = splitMultiHdgs(vals)
            return vals
    return None


def convert_date(dt_str, letter_date):
    """
    Converts an invalid formatted date into a proper date for ARCA Mods
    Correct format:  Y-m-d
    Fixes:
    Incorrect format: m-d-Y
    Incorrect format (letter date): m-d e.g. Jun-17
    :param dt_str: the date string
    :param letter_date: whether the string is a letter date. Letter date is something like Jun-17
    :return: the correctly formatted date
    """
    if letter_date:
        rev_date = datetime.datetime.strptime(dt_str, '%b-%y').strftime('%Y-%m')  # convert date to yymm string format
        rev_date_pts = rev_date.split("-")
        year_num = int(rev_date_pts[0])
        if year_num > 1999:
            year_num = year_num - 100
        year_str = str(year_num)
        rev_date_pts[0] = year_str
        revised = "-".join(rev_date_pts)

    else:
        revised = datetime.datetime.strptime(dt_str, '%d-%m-%Y').strftime(
            '%Y-%m-%d')  # convert date to YY-mm string format

    return revised

def sortValues(lst):
    for item in lst:
        if pd.isna(item):
            lst.remove(item)
    lst = set(lst)
    lst = list(lst)
    return lst

def dropNullCols(df):
    nullcols = []
    for col in df.columns:
        notNull = df[col].notna().sum()
        if notNull < 1:
            nullcols.append(col)
    return nullcols
 

def convert():
    probCol = False
    df2  = pd.DataFrame(columns = col_names)
    df2.append(pd.Series(), ignore_index=True)
    f=filelist[0]
#     if not os.path.exists(savePath):    #if folder does not exist
#         os.makedirs(savePath) 
    try:
        df = pd.read_csv(f,dtype = "string", encoding = 'utf_7')
    except UnicodeDecodeError:
        df = pd.read_csv(f,dtype = "string", encoding = 'utf_8')
   
    nullcols = dropNullCols(df)
    df.drop(nullcols, axis=1, inplace=True)


    i = 1
    for item in df.itertuples():
        #PID
        df2.at[i, 'BCRDHSimpleObjectPID'] = item.PID
        if 'mods_subject_name_personal_namePart_ms' in df.columns:
            pNames = item.mods_subject_name_personal_namePart_ms
        
        #ContentModel
        cModel = item.RELS_EXT_hasModel_uri_s
        df2.at[i,"IslandoraContentModel"] =contentModels[cModel]
        
        #Local Identifier
        if 'mods_identifier_local_ms' in df.columns:
            localID = item.mods_identifier_local_ms
            if pd.notna(localID) and localID != "None":
                df2.at[i,'localIdentifier'] = localID
         
        #Access Identifer
        if 'mods_identifier_access_ms' in df.columns:
            accessID = item.mods_identifier_access_ms
            if pd.notna(accessID):
                df2.at[i,'accessIdentifier'] = accessID
        #Image Link
        # Link to Image
        
        PIDparts = item.PID.split(":")
        repo = PIDparts[0] #repository code
        num = PIDparts[1] #auto-generated accession number
        imageLink = "https://bcrdh.ca/islandora/object/" + repo + "%3A" + num 
        df2.at[i, 'imageLink'] = imageLink
 
        #Title
        if 'mods_titleInfo_title_ms' in df.columns:
            title = item.mods_titleInfo_title_ms
            if pd.notna(title):
                df2.at[i,'title'] = title.replace("\,",",")
 
        #Alternative Title
        if "mods_titleInfo_alternative_title_ms" in df.columns:
            altTitle = item.mods_titleInfo_alternative_title_ms
            if pd.notna(altTitle):
                df2.at[i, 'alternativeTitle'] = altTitle.replace("\,",",")
 
        #Date
        if "mods_originInfo_dateIssued_ms" in df.columns:
            
            dt = item.mods_originInfo_dateIssued_ms
            if pd.notna(dt):
                if (re.match(pattern1, dt)): #letter date, i.e. Jun-17
                    dt = convert_date(dt, True)
                elif (re.match(pattern2, dt)): #reverse date
                    dt = convert_date(dt, False)
                df2.at[i,'dateCreated'] = dt
     
        #Date Issued Start
        if 'mods_originInfo_encoding_w3cdtf_keyDate_yes_point_start_dateIssued_ms' in df.columns:
            startDt = item.mods_originInfo_encoding_w3cdtf_keyDate_yes_point_start_dateIssued_ms
            if pd.notna(startDt):
                df2.at[i,'dateIssued_start'] = startDt

        #Date Issued End
        if 'mods_originInfo_encoding_w3cdtf_keyDate_yes_point_end_dateIssued_ms' in df.columns:
            endDt = item.mods_originInfo_encoding_w3cdtf_keyDate_yes_point_end_dateIssued_ms
            if pd.notna(endDt):
                df2.at[i,'dateIssued_end'] = startDt
                
        #Publisher
        if 'mods_originInfo_publisher_ms' in df.columns:
            pub = item.mods_originInfo_publisher_ms
            if pd.notna(pub):
                df2.at[i, 'publisher_original'] = pub
     
        #Publisher Location
        if 'mods_originInfo_place_placeTerm_text_ms' in df.columns:
            place = item.mods_originInfo_place_placeTerm_text_ms
            if pd.notna(place):
                df2.at[i,'publisher_location'] = place
     
        #Frequency (serials only)
        if 'mods_originInfo_frequency_ms' in df.columns:
            freq = item.mods_originInfo_frequency_ms
            if pd.notna(freq):
                df2.at[i,'frequency'] = freq
     
        #Extent
        if "mods_physicalDescription_extent_ms" in df.columns:
            extent = item.mods_physicalDescription_extent_ms
            if pd.notna(extent):
                extent = extent.replace("\,",",")
                df2.at[i, 'extent'] = extent
        
        #Notes
        if 'mods_note_ms' in df.columns:
            notes = item.mods_note_ms
            if pd.notna(notes):
                notes = notes.replace("\,",",")
                df2.at[i, 'notes'] = notes  
        
        #Description/Abstract
        if "mods_abstract_ms" in df.columns:
            descr = item.mods_abstract_ms
            if pd.notna(descr):
            #if descr is not None:
                df2.at[i, 'description'] = descr.replace("\,",",")
        
        #Personal Creators & Contributors
        
        if 'mods_name_personal_namePart_ms' in df.columns:
            names = item.mods_name_personal_namePart_ms
            if pd.notna(names):
                names = splitMultiHdgs(names)
                roles = getMultiVals(item,"personal_role",df,pd)
                if len(roles)==0:
                    pass
                else:
                    creatorCount = 0
                    contribCount = 0
                    for x in range(len(names)):
                        if roles[x] == "creator":
                            creatorCount = creatorCount + 1
                            hdg = "creator" + str(creatorCount)
                            df2.at[i,hdg] = names[x].strip()
                     
                        else:
                            contribCount = contribCount + 1
                            hdg = "contributor" + str(contribCount)
                            df2.at[i,hdg] = names[x].strip()    
                        
            
                        
                
        #Corporate Creators and Contributors
        if 'mods_name_corporate_namePart_ms' in df.columns:
            corpNames = item.mods_name_corporate_namePart_ms
            if pd.notna(corpNames):
                creatorCount = 0
                contribCount = 0
                if 'mods_name_corporaterole_roleTerm_ms' in df.columns:
                    corpRoles = item.mods_name_corporate_role_roleTerm_ms
                    if pd.notna(corpRoles):
                        corpRoles = corpRoles.split(",")
                else:
                    corpRoles = np.nan
                corpNames = splitMultiHdgs(corpNames)
                count = 0
                for corpName in corpNames:
                    if (pd.isna(corpRoles) or corpRoles[count]=='creator'):
                        creatorCount = creatorCount + 1
                        hdg = "corporateCreator" + str(creatorCount)
                        df2.at[i,hdg] = corpName
                    else:
                        contribCount = contribCount + 1
                        hdg = "corporateContributor" + str(contribCount)
                        df2.at[i,hdg] = corpName
 
        #topical subjects
        if 'mods_subject_topic_ms' in df.columns:
            topics = item.mods_subject_topic_ms
            if pd.notna(topics):
                topics = splitMultiHdgs(topics)
                for x in range(len(topics)):
                    hdg = "topicalSubject" + str(x+1)
                    df2.at[i, hdg] = topics[x]
         
        #corporate subjects
        if 'mods_subject_name_corporate_namePart_ms' in df.columns:
            corpSubs = item.mods_subject_name_corporate_namePart_ms
            if pd.notna(corpSubs):
                corpSubs = splitMultiHdgs(corpSubs)
                corpSubs = list(set(corpSubs)) #remove duplicates
                for x in range(len(corpSubs)):
                    hdg = "corporateSubject" + str(x+1) 
                    df2.at[i, hdg] = corpSubs[x]

#        #personal subjects
        if 'mods_subject_name_personal_namePart_ms' in df.columns:
            pnames = item.mods_subject_name_personal_namePart_ms
            if pd.notna(pnames):
                pnames = splitMultiHdgs(pnames)
                for x in range(len(pnames)):
                    hdg = "personalSubject" + str(x+1)
                    if pd.notna(pnames[x]):
                        df2.at[i,hdg] = pnames[x].strip()
        
        #temporal subject (date range)
        if 'mods_subject_temporal_ms' in df.columns:
            tempSub = item.mods_subject_temporal_ms
            if pd.notna(tempSub):
                df2.at[i,'dateRange'] = tempSub
         
        #geographic subject
        if 'mods_subject_geographic_ms' in df.columns:
            geosub = item.mods_subject_geographic_ms
            if pd.notna(geosub):
                geosubs = splitMultiHdgs(geosub)
                for x in range(len(geosubs)):
                    hdg = "geographicSubject" + str(x+1) 
                    df2.at[i, hdg] = geosubs[x]
         
        #coordinates 
        if 'mods_subject_geographic_cartographics_ms' in df.columns:
            coords = item.mods_subject_geographic_cartographics_ms
            if pd.notna(coords):
                df2.at[i,"coordinates"] = coords
                
        #classification
        if 'mods_classification_authority_lcc_ms' in df.columns:
            lcClass = item.mods_classification_authority_lcc_ms
            if pd.notna(lcClass):
                df2.at[i,'classification'] = lcClass
        
        #isbn
        if 'mods_identifier_isbn_ms' in df.columns:
            isbn = item.mods_identifier_isbn_ms
            if pd.notna(isbn):
                df2.at[i,'ISBN'] = isbn
        
        #genre
        if 'mods_genre_authority_aat_ms' in df.columns:
            genre_aat = item.mods_genre_authority_aat_ms
            if pd.notna(genre_aat):
                df2.at[i, 'genre'] = genre_aat
                df2.at[i, 'genreAuthority'] = "aat"
        elif 'mods_genre_authority_marcgt_ms' in df.columns:
            if pd.notna(item.mods_genre_authority_marcgt_ms):
                df2.at[i, 'genre'] = item.mods_genre_authority_marcgt_ms
                df2.at[i, 'genreAuthority'] = "marcgt"
        elif 'mods_originInfo_genre_ms' in df.columns:
            if pd.notna(item.mods_originInfo_genre_ms):
                df2.at[i, 'genre'] = item.mods_originInfo_genre_ms
             
        #type
        if 'mods_typeOfResource_ms' in df.columns:
            _type = item.mods_typeOfResource_ms
            if pd.notna(_type):
                df2.at[i, 'type'] = _type
         
        #internet media type
        if 'mods_physicalDescription_internetMediaType_ms' in df.columns:
            mediaType = item.mods_physicalDescription_internetMediaType_ms
            if isinstance (mediaType, str):
                df2.at[i, 'internetMediaType'] = mediaType
         
        #Languages
        languages = None
        langs = getMultiVals(item,"languageTerm",df,pd)
        if langs is not None:
            for x in range(len(langs)):
                lang = langs[x]
                hdg = "language" + str(x+1)
                if pd.notna(lang):
                    df2.at[i,hdg] = lang
         
        #Source
        if 'mods_location_physicalLocation_ms' in df.columns:
            source = item.mods_location_physicalLocation_ms
            if pd.notna(source):
                df2.at[i, 'source'] = source
        
        #URI
        if 'mods_identifier_uri_ms' in df.columns: 
            uri = item.mods_identifier_uri_ms
            if pd.notna(uri):
                df2.at[i, 'URI'] = uri
            
        #Rights
        if 'mods_accessCondition_use_and_reproduction_ms' in df.columns:
            rights = item.mods_accessCondition_use_and_reproduction_ms
            if isinstance(rights, str):
                rights = splitMultiHdgs(item.mods_accessCondition_use_and_reproduction_ms)
            
                for stmt in rights:
                    if "Permission to publish" in stmt:
                        df2.at[i, "rights"] = stmt
                    elif "rightsstatements" in stmt:
                        df2.at[i,"rightsStatement_URI"] = stmt
                    else:
                        df2.at[i,"creativeCommons_URI"] = stmt
                 
        #Related Item
        if 'mods_relatedItem_host_titleInfo_title_ms' in df.columns:
            coll_title = item.mods_relatedItem_host_titleInfo_title_ms
            coll_PID = item.mods_relatedItem_host_identifier_PID_ms
            if pd.notna(coll_title):
                df2.at[i, "relatedItem_title"] = coll_title
            if pd.notna(coll_PID):
                df2.at[i, "relatedItem_PID"] = coll_PID
                 
        #Record Origin & Creation Date
        if 'mods_recordInfo_recordOrigin_ms' in df.columns:
            recOrigin = item.mods_recordInfo_recordOrigin_ms
            if pd.notna(recOrigin):
                df2.at[i,'recordOrigin'] = recOrigin
            recDate = item.mods_recordInfo_recordCreationDate_ms
            if pd.notna(recDate):
                recDate = recDate.split(",")
                df2.at[i, 'recordCreationDate'] = recDate[0]
        if (probCol):
            probColls.append(f)
        i = i + 1
    
    f = os.path.basename(f).replace(".csv","_Rev1.csv")
    dest = os.path.join(desktopPath,f)
    df2.to_csv(dest, encoding='utf-8', index=False)
    
    msg = f + " has been written to the your desktop."
    messagebox.showinfo(title = 'Conversion Finished', message = msg)

    
root = Tk()

root.eval('tk::PlaceWindow %s center' % root.winfo_toplevel())
root.title = "Arca GetCSV -> BCRDH CSV Converter"
root.configure(background='#DAE6F0')
folder_path = StringVar()

intro = ttk.Label(master=root,text="Arca_GetCSV_Converter Version 2.0.0", background='#DAE6F0',font="Arial 15 bold")
intro.grid(row=0,column=1,padx=(10,0),sticky='w')
info = ttk.Label(master=root,text="This app converts an Arca Get Metadata CSV to an MMW 18-6 BCRDH\nmetadata spreadsheet. The revised CSV, with 'Rev1' appended to it,\nwill be written to your desktop. Single-field personal-name (combined\ngiven name and family names) data are assumed as input.", background='#DAE6F0',font="Arial 9 bold")
info.grid(row=1,column=1,padx=(10,0),sticky='w')


srclbl = ttk.Label(master=root,text="\nChoose CSV to convert",background='#DAE6F0',font="Arial 10 italic")
srclbl.grid(row=2,column=1,padx=(10,0),sticky='w')

button1 = ttk.Button(text="Browse", command=browse_button)
button1.grid(row=3, column=1,  padx=(10,0),sticky='w')

lbl1 = ttk.Label(master=root,background='#DAE6F0',font="Arial 10")
lbl1.grid(row=4, column=1,padx=(10,5), sticky='w')

compButton = ttk.Button(text="Generate BCRDH CSV",command=convert)
compButton.grid(row=8, column=1,padx=(10,0),pady=(15,15),sticky='w')

mainloop()