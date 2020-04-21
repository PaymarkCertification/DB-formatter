import pandas as pd
import numpy as np

__author__ = 'Michael Yu'


def again():
    while True:
        djKhaled = input("Another one? Y/N: ")
        if djKhaled.lower() not in ('y', 'n'):
            print("Invalid input, enter Y or N")
        else:
            break

    if djKhaled.lower() == 'y':
        main()
    elif djKhaled.lower() == 'n':
        print("Exiting")
        exit()


def adaptavist_tc(filename):
    reader = pd.read_excel(filename+'.xlsx')
    print("Reading file to memory")
    ndf = pd.DataFrame(reader)
    print("Setting to Data Frame")
    df = pd.DataFrame(ndf[['Key', 'Name', 'Component', 'Labels', 'Coverage (Issues)']])
    df.insert(4, "Labels2", "", True)
    df.insert(5, "Labels3", "", True)
    df.insert(6, "Labels4", "", True)
    df.insert(7, "Labels5", "", True)
    df.insert(8, "Labels6", "", True)
    df.insert(9, "Labels7", "", True)
    df.insert(10, "Labels8", "", True)
    print("inserting labels as columns")
    df[['Labels', 'Labels2', 'Labels3', 'Labels4', 'Labels5', 'Labels6', 'Labels7', 'Labels8']] = df['Labels']\
        .str.split(',', expand=True)
    print("Splitting Label data")
    df.insert(12, "Link2", "", True)
    df.insert(13, "Link3", "", True)
    df.insert(14, "Link4", "", True)
    df.insert(15, "Link5", "", True)
    df.insert(16, "Link6", "", True)
    df.insert(17, "Link7", "", True)
    df.insert(18, "Link8", "", True)
    df.insert(19, "Link9", "", True)
    df.insert(20, "Link10", "", True)
    df.insert(21, "Link11", "", True)
    df.insert(22, "Link12", "", True)
    df.insert(23, "Link13", "", True)
    df.insert(24, "Link14", "", True)
    df.insert(25, "Link15", "", True)
    df.insert(26, "Link16", "", True)
    df.insert(27, "Link17", "", True)
    df.insert(28, "Link18", "", True)
    df.insert(29, "Link19", "", True)
    df.insert(30, "Link20", "", True)
    df.insert(31, "Link21", "", True)
    df.insert(32, "Link22", "", True)
    df.insert(33, "Link23", "", True)
    df.insert(34, "Link24", "", True)
    df.insert(35, "Link25", "", True)
    df.insert(36, "Link26", "", True)
    print("inserting links as columns")
    df[['Coverage (Issues)', 'Link2', 'Link3', 'Link4', 'Link5', 'Link6', 'Link7', 'Link8', 'Link9', 'Link10', 'Link11',
        'Link12', 'Link13', 'Link14', 'Link15', 'Link16', 'Link17', 'Link18', 'Link19', 'Link20', 'Link21', 'Link22',
        'Link23', 'Link24', 'Link25', 'Link26']] = df[
        'Coverage (Issues)'].str.split(',', expand=True)
    print("splitting Coverage data")

    df.rename({'Labels': 'Labels1', 'Coverage (Issues)': 'Link1'}, axis=1, inplace=True)
    print("Renaming Label, Coverage Columns")

    df['Key'].replace('', np.nan, inplace=True)

    print("removing null rows")
    df.dropna(subset=['Key'], inplace=True)

    writer = pd.ExcelWriter('adaptavist_clean.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet1', index=False)
    try:
        writer.save()
        print("saving file")
        again()

    except Exception as e:
        print("close the file: ", e)



def jiraRequirements(filename):

    #parsing JQL project = TAC AND issuetype = Story and labels != 'STP'
    reader = pd.read_csv(filename+'.csv')
    print("Reading file to memory")
    ndf = pd.DataFrame(reader)

    try:
        df = pd.DataFrame(ndf[['Summary', 'Issue key', 'Issue id', 'Issue Type', 'Project key', 'Creator', 'Components',
                               'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels',
                               'Description']])
        print("Setting to Data Frame")
    except KeyError:
        print("Missing key in index")
        ndf.insert(6, "Components", "", True)
        print("Adding header")
        df = pd.DataFrame(ndf[['Summary', 'Issue key', 'Issue id', 'Issue Type', 'Project key', 'Creator', 'Components',
                               'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels', 'Labels',
                               'Description']])
        print("Setting to Data Frame")
    writer = pd.ExcelWriter('Jira Rq clean.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet1', index=False)
    try:
        writer.save()
        print("saving file")
        again()
    except Exception as e:
        print("close the file: ", e)

def main():
    print("     A: PTS requirements")
    print("     B: Adaptavist Manual Test Cases")
    print("\nPlease select which schema you will like to use: A OR B")
    while True:
        choice = input("Option: ")
        if choice.lower() not in ('a', 'b', 'exit'):
            print("Invalid option. Select option A or B")
        else:
            break

    if choice.lower() == 'a':
        print("Loading: PTS requirements")
        print("Ensure file is in root folder!")
        filename = input("enter file name: ")
        try:
            jiraRequirements(filename)
        except OSError as e:
            print("No such file exists", e)
            print("exiting...")


    if choice.lower() == 'b':
        print("Loading: Adaptavist Manual Test Cases")
        print("Ensure file is in root folder!")
        filename = input("enter file name: ")
        try:
            adaptavist_tc(filename)
        except OSError as e:
            print("No such file exists", e)

    if choice.lower() == 'exit':
        print("exiting")
        exit()

main()
