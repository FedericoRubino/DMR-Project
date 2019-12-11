# -*- coding: utf-8 -*-
"""
Created on Wed Oct 23 08:11:26 2019

@author: frubino

This is the script to get data from the APAC/PLC DMR forms!
"""

# import xlrd
import numpy as np
import pandas as pd


class ApacDMR:

    def __init__(self, filepath):
        df_raw = pd.read_excel(filepath, sheet_name=0)
        self.df = df_raw.replace(np.nan, "", regex=True, inplace=False)
        self.po_num = ""
        self.date_receiv = ""
        self.sqe = ""
        self.problem = ""
        self.closure_date = ""
        self.preventive_action = ""
        self.m_revision = ""
        self.team = ""
        self.root_cause_description = ""
        self.preventive_action = ""
        self.temporary_corrective_action = ""
        self.run_dmr()

    # return original Closure Date
    def original_closure(self):
        date = ""
        for row in self.df.values:
            for i, cell in enumerate(row):
                if type(cell) is str and "Date Corrective Action Closed" in cell:
                    j = 1
                    while (row[i + j] == "") and ((i+j) < len(row) - 1):
                        j += 1
    #                print(row)
                    date = row[i + j]
        self.closure_date = date
        return date
    
    # return SQE Approval
    def sqe_approval(self):
        sqe = ""
        for row in self.df.values:
            for i, cell in enumerate(row):
                if type(cell) is str and "SQE Approval" in cell:
                    j = 1
                    while (row[i + j] == "") and ((i+j) < len(row) - 1):
                        j += 1
                    # print(row)
                    sqe = row[i + j]
        self.sqe = sqe
        return sqe
    
    # returns the date received
    def date_received(self):
        try:
            date = ""
            for row in self.df.values:
                for j, x in enumerate(row):
                    if type(x) is str and "Received Date" in x:
                        i = 1
                        while row[j + i] == "":
                            i += 1
                        if "Inspect" not in row[j + i]:
                            date = row[j + i]
            self.date_receiv = date
            return date
        except Exception as e:
            return 0
    
    # should return the place the problem was detected at
    def problem_detected_at(self):
        problem = ""
        for x in self.df.values:
            if "Problem Detected at" in x:
                for i, cell in enumerate(x):
                    # print(cell)
                    if type(cell) is str and "Problem" in cell:
                        problem = x[i + 2]
                        # print(x[i + 2])
        # this doesn't quite work yet...
        # there is a check box that is checked
        self.problem = problem
        return problem
    
    # get the PO/Invoice Number
    def po_invoice_num(self):
        po_num = ""
        for x in self.df.values:
            if "P.O." in x:
                for i, cell in enumerate(x):
                    if type(cell) is str and "P.O." in cell:
                        po_num = x[i + 1]
                        # print(x[i + 1])
        self.po_num = po_num
        return po_num
    
    # finds and returns the values for pmx_qe, pmx_sqe, scz, apo_sqe
    # this will be used to fill the team cell
    # as well as Quality Engineer column
    def get_sqe(self):
        pmx_qe = ""
        pmx_sqe = ""
        scz = ""
        apo_sqe = ""   
        for x in self.df.values:
            if "PMX SQE" in x:
                for i, cell in enumerate(x):
                    if type(cell) is str and "PMX QE" in cell:
                        # print(x[i + 1])
                        pmx_qe = x[i + 1]
                    if type(cell) is str and "PMX SQE" in cell:
                        # print(x[i + 1])
                        pmx_sqe = x[i + 1]
                    if type(cell) is str and "SCZ SQE" in cell:
                        # print(x[i + 1])
                        scz = x[i + 1]
                    if type(cell) is str and "APO SQE" in cell:
                        # print(x[i + 1])
                        apo_sqe = x[i + 1]
        team = pmx_qe + " " + pmx_sqe + " " + scz + " " + apo_sqe
        self.team = team.strip()
        return team.strip()
                                   
    # finds and returns the revision number
    def revision(self):
        for x in self.df.values:
            if "Revision" in x:
                for i, cell in enumerate(x):
                    if type(cell) is str and "Revision" in cell:
                        # print(x[i + 1])
                        self.m_revision = x[i + 1]
                        return x[i + 1]

    # this function will return root cause description
    # and permanent corrective action action
    def cause_prevention_action(self):
        temp = self.df.iloc[:, :1]
        temp.replace(np.nan, "", regex=True, inplace=True)
        
        root_cause_description = ""
        preventive_action = ""
        temporary_corrective_action = ""
        
        for i, x in enumerate(temp.values):
            if "Root Cause of Non Conformance" in x[0]:
                # print(temp.values[i + 1][0])
                root_cause_description = temp.values[i + 1][0]
            if "Permanent Corrective Action" in x[0]:
                # print(temp.values[i + 1][0])
                preventive_action = temp.values[i + 1][0]
            if "Temporary Corrective Action" in x[0]:
                # print(temp.values[i + 1][0])
                temporary_corrective_action = temp.values[i + 1][0]
        self.root_cause_description, self.preventive_action,\
            self.temporary_corrective_action = root_cause_description, preventive_action, temporary_corrective_action
        return root_cause_description, preventive_action, temporary_corrective_action

    def run_dmr(self):
        self.cause_prevention_action()
        self.date_received()
        self.get_sqe()
        self.original_closure()
        self.po_invoice_num()
        self.problem_detected_at()
        self.revision()
        self.sqe_approval()


# main is there for testing purposes
# will not use this code for the final piece
def main():
    filename = "C:/Users/frubino/DMR Project/DMR_forms_for_code/APAC DMR forms/2015/2015-053/2015-053- pn202828-43 _DMR.xls"
    apac_dmr = ApacDMR(filename)
    
    print("Cause and Preventative actions:", apac_dmr.cause_prevention_action())
    print()
    print("Problem detected at (not quite functional):", apac_dmr.problem_detected_at())
    print()
    print("Revision number:", apac_dmr.revision())
    print()
    print("SQE's involved:", apac_dmr.get_sqe())
    print()
    print("The PO invoice number:", apac_dmr.po_invoice_num())
    print()
    print("The date received:", apac_dmr.date_received())
    print()
    print("The original closure date:", apac_dmr.original_closure())
    print()
    print("SQE's approval:", apac_dmr.sqe_approval())


if __name__ == "__main__":
    main()
