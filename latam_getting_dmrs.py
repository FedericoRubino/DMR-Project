# -*- coding: utf-8 -*-
"""
Created on Fri Nov  1 14:11:21 2019

@author: frubino

This file will handle reading all the different documents and organizing it somehow


"""


def get_file_paths(path_name):
    import os, pathlib

    list_of_paths = list()

    for root, dir, files in os.walk(path_name):
        for file in files:
            if (file.endswith('.xls') or file.endswith('.xlsx')): #and "~$" not in file:
                if file not in list_of_paths:
                    p = pathlib.PureWindowsPath(root + "\\" + file)
                    list_of_paths.append(str(p.as_posix()))
    # testing purposes
    # for items in list_of_paths:
    #     print(items)
    return list_of_paths


def run_latams():
    from xlrd import XLRDError
    from LatamDMR_information import LatamDMR

    filename = "C:/Users/frubino/DMR Project/DMR_forms_for_code/Latam DMR forms"
    list_of_dmrs = list()
    list_of_dmrs = get_file_paths(filename)
    
    # dictionary of dmr objects
    # the key is the dmr id
    dic_of_dmrs = {}

    for dmr in list_of_dmrs:
        #print()
        #print(dmr)
        try:
            dmrObj = LatamDMR(dmr)
            temp = dmr.split("/")
            doc_name = temp[len(temp) - 1]
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
#            print(e)

    return dic_of_dmrs


if __name__ == "__main__":
    run_latams()
