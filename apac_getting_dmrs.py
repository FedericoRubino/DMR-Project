# -*- coding: utf-8 -*-
"""
Created on Fri Nov  1 14:11:21 2019

@author: frubino

This file will handle reading all the different documents and organizing it somehow


"""

from latam_getting_dmrs import get_file_paths
from xlrd import XLRDError
from ApacDMR_information import ApacDMR


def run_apacs():

    filename = "C:/Users/frubino/DMR Project/DMR_forms_for_code/APAC DMR forms"
    list_of_dmrs = list()
    list_of_dmrs = get_file_paths(filename)
    
    # dictionary of dmr objects
    # the key is the dmr id
    dic_of_dmrs = {}
    
    for dmr in list_of_dmrs:
        #print()
        #print(dmr)
        try:
            dmrObj = ApacDMR(dmr)
            temp = dmr.split("/")
            doc_name = temp[len(temp) - 1]
#            print(doc_name)
            if ".xls" in doc_name:
                doc_name = doc_name.replace(".xls", "")
            elif ".xlsx" in doc_name:
                doc_name = doc_name.replace(".xlsx", "")
            dic_of_dmrs[doc_name] = dmrObj
        except FileNotFoundError:
            print("File not found")
        except XLRDError:
            print("encryption error")
        except Exception as e:
            pass
            
    return dic_of_dmrs


if __name__ == "__main__":
    run_apacs()
