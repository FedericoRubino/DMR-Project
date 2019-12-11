# -*- coding: utf-8 -*-
"""
Created on Wed Oct 23 08:11:26 2019

@author: frubino

This is the script to get data for the LATAM/Plamex DMR forms!
"""
# import xlrd


class LatamDMR:

    def __init__(self, filepath):
        import numpy as np
        import pandas as pd  
        
        raw_df = pd.read_excel(filepath)  # , sheet_name = 1)
        self.df = raw_df.replace(np.nan, "", regex=True, inplace=False)
        self.team = ""
        self.revision = ""
        self.dateReceived = ""
        self.effDate01 = ""
        self.effDate02 = ""
        self.permanentCorr = ""
        self.problemDetected = ""
        self.qualEngineer = ""
        self.rootCause = ""
        self.tempCorrAction = ""
        self.run_dmr()

    # returns the Quality Engineer in charge of the DMR
    def quality_engineer(self):
        engi = ""
        for row in self.df.values:
            for i, cell in enumerate(row):
                if type(cell) is str and ("Plamex QE" in cell or "PMX QE" in cell):
                    j = 1
                    while row[i + j] == "":
                        j += 1
    #                print(row[i + j])
                    engi = row[i + j]
        self.qualEngineer = engi
        return engi

    # Returns the revision number found in the DMR sheet
    def get_revision(self):
        revision = "N/A"
        for i, col in enumerate(self.df.values):
            for j, x in enumerate(col):
                if type(x) is str and "Revision" in x:
                    # print(self.df.values[i+1][j])
                    revision = self.df.values[i][j+1]
        self.revision = revision
        return revision
        
    # returns the date received
    def date_received(self):
        date = ""
        for i, row in enumerate(self.df.values):
            for j, x in enumerate(row):
                if type(x) is str and ("Date Rec" in x or "Issue Date" in x):
                    # print(self.df.values[i+1][j])
                    k = 1
                    while self.df.values[i][j + k] == "":
                        k += 1
                    date = self.df.values[i][j + k]
        self.dateReceived = date
        return date
    
    # returns the location the problem was detected at
    def problem_detected_at(self):
        detected = ""
        for i, row in enumerate(self.df.values):
            for j, cell in enumerate(row):
                if type(cell) is str and "Problem Detected At" in cell:
                    for k, cel in enumerate(self.df.values[i]):
                        if "X" in cel:
                            # print(self.df.values[i][k + 1])
                            detected = self.df.values[i][k + 1]
        self.problemDetected = detected
        return detected
                    
    # returns the team responsible for the DMR
    def generate_team(self):
        team = "N/A"
        
        for i, row in enumerate(self.df.values):
            for j, x in enumerate(row):
                if type(x) is str and ("Responsible" in x):
                    # print(self.df.values[i+1][j])
                    k = 2
#                    while self.df.values[i][j + k] == "":
#                        k += 1
                    team = self.df.values[i][j + k]
        
        # other form type
        if team == "N/A":
            for i, col in enumerate(self.df.values):
                for j, x in enumerate(col):
                    if type(x) is str and "Inspector" in x:
                        # print(self.df.values[i+1][j])
                        team = self.df.values[i+1][j]
            self.team = team
        return team

    # returns the different reports found on the DMR
    # such as the rootCause, the temporary Corrective action and
    # the Permanent Corrective Action
    def reports(self):
        root_cause = ""
        temp_corr_action = ""
        permanent_corr = ""
        col1 = self.df["Unnamed: 0"]
        for i,cell in enumerate(col1):
            if "Temporary Corrective" in cell:
                temp_corr_action = col1[i + 1]
            if "Permanent Corrective Action" in cell:
                permanent_corr = col1[i + 1]
            if "Root Cause" in cell:
                root_cause = col1[i + 1]
                
        # different forms have different layouts
        if root_cause == "" or temp_corr_action == "" or permanent_corr == "":
            col1 = self.df["Unnamed: 1"]
            for i, cell in enumerate(col1):
                if "Root Cause" in cell:
                    j = 1
                    while "Temporary Corrective" not in col1[i + j]:
                        # print(col1[i + j])
                        root_cause += col1[i + j] + "\n"
                        j += 1
                if "Temporary Corrective" in cell:
                    j = 1
                    while "Effective Date" not in col1[i + j]:
                        # print(col1[i + j])
                        temp_corr_action += col1[i + j] + "\n"
                        j += 1
                if "Permanent Corrective Action" in cell:
                    j = 1
                    while "Effective Date" not in col1[i + j]:
                        # print(col1[i + j])
                        permanent_corr += col1[i + j] + "\n"
                        j += 1 
        self.rootCause, self.tempCorrAction, self.permanentCorr = root_cause, temp_corr_action, permanent_corr
        return root_cause, temp_corr_action, permanent_corr

    # returns the Containmaint Action Effective date as date01
    # returns the Corrective action Effective date as date02 - not sure if needed
    def effective_dates(self):
        date01 = ""
        date02 = ""
        for i, row in enumerate(self.df.values):
            for j, x in enumerate(row):
                if type(x) is str and ("Effectivity Date" in x):
                    # print(self.df.values[i+1][j])
                    k = 1
                    while self.df.values[i][j + k] == "":
                        k += 1
                    if date01 == "":
                        date01 = str(self.df.values[i][j + k])
                    else:
                        date02 = str(self.df.values[i][j + k])
        
        # other form type needs this
        if date01 == "" and date02 == "":
            col1 = self.df["Unnamed: 1"]
            for i, cell in enumerate(col1):
                if "Effective Date" in cell and date01 == "":
                    date01 = cell.replace("Effective Date:", "")
                elif "Effective Date" in cell and date01 != "":
                    date02 = cell.replace("Effective Date:", "")
        
        self.effDate01, self.effDate02 = date01, date02
        return date01, date02

    # this will populate all the member variables
    def run_dmr(self):
        self.date_received()
      
        self.problem_detected_at()
    
        self.quality_engineer()
     
        self.reports()
     
        self.get_revision()
       
        self.generate_team()
        
        self.effective_dates()
        

# main is there for testing purposes
# will not use this code for the final piece
def main():
    filepath = "C:/Users/frubino/DMR Project/DMR_forms_for_code/Latam DMR forms/Latam master DMR links/0001-0300-15/0201-0300-15/0271-15 PN203431-02.xls"
    latam_dmr = LatamDMR(filepath)

    print("Date Rec:", latam_dmr.date_received())
    print()
    print("Problem Detected at:", latam_dmr.problem_detected_at())
    print()
    print("Quality Engineer:", latam_dmr.quality_engineer())
    print()
    print("Different reports:")
    for report in latam_dmr.reports():
        print()
        print("##:",report)
    
    
    print()
    print("Revision:", latam_dmr.get_revision())
    print()
    print("Team in charge:", latam_dmr.generate_team())
    print()
    print("effictivity dates:", latam_dmr.effective_dates())


if __name__ == "__main__":
    main()
