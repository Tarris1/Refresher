import pandas as pd
from math import ceil
from tkinter import Tk
from tkinter import filedialog as fd
import docx
import os
from dotenv import load_dotenv



def find(search, df, name = None):
    ''' Function that combines all columns and searches for a match. Search from input, df = pandas dataframe '''
    #search = search.split("@")[1].lower()
    cols = list(df.columns)
    df["search"] = df[cols].apply(lambda x: "".join(str(x).lower()), axis = 1)
    result = df.loc[df["search"].str.contains(search, case=False)]
    result.drop(columns=["search"])
    result = result.sort_values(by=["Birth","Death"], ascending = True, na_position="first")
    return result

def find_mult(search, df):
    search = search.split("@")[1].lower()
    search_elements = search.split(" ")
    step = 0
    for i in search_elements:
        if step == 0:
            result = find(i,df)
            step+=1
        else: result = find(i,result)
    return result


def print_people(df):
    '''Prints search results of people'''
    def trans_na(df, col, what):
        if pd.isnull(df[col]): trans = "at an unknown " + what
        else: trans = str(df[col])
        return trans
    def trans_year(df, col):
        if pd.isnull(df[col]): trans = "(unknown)"
        else:
            age = str(abs(ceil(df[col])))
            if df[col]<0: trans = age +"BC"
            else: trans = age +"AD"
        return trans
    col = list(df.index)
    text = ""
    for i in col:
        row = df.loc[i]
        birth = trans_year(row, "Birth")
        death = trans_year(row,"Death")
        position = row["Position"]
        city = trans_na(row,"City/Region","city")
        country = trans_na(row, "Country", "country")
        name = row["Name"]
        if "," in name: name = name.split(",")[0]
        else: name = name
        to_print = name + ": "+position + ", born " + birth + " in " + city + ", " + country + ", and died " + death
        text+= to_print+"\n\n"
    return text

#Functions for updating/fixing data, reupload to Excel-file, adding data from another excel sheet, etc.
def fill_data(df):
    '''Fixes birth & death dates on people without a specified date based on floruit input'''
    cols_to_fix = ["Birth", "Death"]
    stopwords = ["century", "st", "th", "rd", "nd", "ad", "bc", "pre", "late", "h", "?", "or"] #Removal list
    dfc = df.copy()
    col = list(dfc.index) #index list
    for i in col:
        row = dfc.loc[i]
        if (pd.isnull(row["Birth"]) or pd.isnull(row["Death"])) and not pd.isnull(row["Floruit"]):
            floruit = str(row["Floruit"]).lower()
            if "ad" in floruit:
                period = "ad"
                addition_birth = 0
                addition_death = 99
            else:
                period = "bc"
                addition_birth = +99
                addition_death = 0
            querywords = floruit.split(' ') #Split all words in floruit
            resultwords = [word for word in querywords if word not in stopwords]
            result = ''.join(resultwords)
            for word in stopwords: result = result.replace(word,"")
            result_range = result.split("-")
            if not result_range[0] or "," in result_range[0]: continue #Skip problematic data
            result_range = [(int(float(word))-1)*100 for word in result_range if int(float(word))<100]
            if len(result_range) == 0: continue #Skip those without floruit data
            if len(result_range)==2:
                birth = result_range[0]+addition_birth
                death = result_range[1]+addition_death
            else:
                birth = result_range[0]+addition_birth
                death = result_range[0]+addition_death
            if period == "bc":
                birth = birth*-1
                death = death*-1
            if pd.isnull(row["Birth"]): dfc.loc[i,"Birth"] = birth
            if pd.isnull(row["Death"]): dfc.loc[i,"Death"] = death
    return dfc        

def help():
    print (
        '''
            Available functions: \n
                "search @<input>" - searches for any match in the database. Multiple criterias may be searched from with a simple space \n
                "exit/quit" - exits the application and allows you to save your data \n
                "read" - reads specified file \n
                "write" - saves the file \n
                "help" - lists all available functions \n
        '''
        )

def write(df, name="data", sheet="sheet1"):
    df.to_excel(name+".xlsx", sheet_name = sheet, index = False)

def read():
    '''Prints and returns the contents of files read'''
    Tk().withdraw()
    filename = fd.askopenfilename()
    data = ''
    if ".xlsx" in filename:
        sheets = pd.ExcelFile(filename).sheet_names
        data = []
        for sheet in sheets:
            df = pd.read_excel(filename, sheet_name = sheet)
            data.append(df)
        print(data)
    elif ".txt" in filename:
        with open(filename) as f:
            lines = f.readlines()
        data = ''.join(lines)
        print(data)
    elif ".docx" in filename:
        doc = docx.Document(filename)
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
        data = '\n'.join(fullText)
        print(data)
    else: print("no printable content found in " + filename)
    return data

def append(df, addon):
    '''Concats every imported list and returns the final dataframe'''
    for i in addon:
        i_filled = fill_data(i)
        dfc = pd.concat([df,i_filled],ignore_index=True).drop_duplicates().reset_index(drop=True)
    return dfc

def find_dupl(df,col):
    duplicates = df[col].duplicated(keep = False)
    duplicates = duplicates.loc[duplicates]
    for i in duplicates.index:
        if (pd.isna(df.loc[i,"Birth"]) or pd.isna(df.loc[i,"Death"])) and not pd.isna(df.loc[i,"Floruit"]):
            floruit = str(df.loc[i,"Floruit"])
        else:
            floruit = "("+str(df.loc[i,"Birth"])+")-("+str(df.loc[i,"Death"])+")"
        print("\n"+df.loc[i,"Name"] + " floruit " + floruit)
    print("\n"+str(duplicates.sum())+" duplicates found")

def main():
    load_dotenv()
    file = os.getenv("base_data")
    cols = "B:J"
    sheet = "People"
    rows = range(0,1)
    data = pd.read_excel(file, sheet_name = sheet, usecols = cols, skiprows=rows).drop_duplicates()
    data = fill_data(data)
    find_dupl(data, "Name")
    read_data = []
    loop = 0

    while loop < 1:
        inp = input("What do you want me to do? ").lower()
        if inp == "quit" or inp == "exit":
            to_save = input("Do you want to save the data? ").lower()
            if "yes" in to_save or "" in to_save or "y" in to_save: write(data)
            quit()
        elif "search" in inp:
            result = find_mult(inp, data)
            people = print_people(result)
            print(people)
            print(str(len(result.index)) + " results have been found from the search '" + inp.split("@")[1] + "'")
        elif inp == "help": help()
        elif inp == "write": write(data)
        elif inp == "read":
            for i in read():
                read_data.append(i)
            print(read_data)
        elif inp == "append":
            if len(read_data)>0: data = append(data, read_data)
            else: continue
            print(data)
        else: print("'"+inp+"' is not a valid function.")

if __name__ == "__main__":
    main()        




