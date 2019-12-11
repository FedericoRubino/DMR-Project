"""
DMR Legacy Data Project

Project involves manipulating data from Excel files

Structure:
- Template for Salesforce
- DMR layout --> connects the different lists
- APAC - CDC DMRs
- LATAM - Plamex DMRs

This turns the master templates of both the China and Mexico DMRs
into the same format as the End product Template

"""

# Import numpy for array usage
# import pandas for dataframe options
import numpy as np
import pandas as pd

# increase display options
pd.set_option('display.max_columns', 300)
pd.options.display.width = 150


def populate_dataframe(master_df):
    
    # At this point we have all of the master data in a form that is similar to 
    # the template
    # now we must populate the cells with the dmr forms:
    from apac_getting_dmrs import run_apacs
    from latam_getting_dmrs import run_latams
    
    # these two dictionaries have all of the neccessary information to fill out the master template
    d_apac = run_apacs()
    d_latam = run_latams()
    
    
    # this is a column for the source material dmrs
    master_df["Source material"] = ""
    
    
    # make sure all of the columns are of type object aka string
    master_df = master_df.replace(np.nan, '', regex=True)

    master_df = master_df.astype(object)
    
    for i, row in enumerate(master_df.values):
        
        # Getting the Material Disposition to be in the correct format
        disposition = master_df.at[i, 'Material Disposition'].lower()
        if "sort" in disposition:
            master_df.at[i, 'Material Disposition'] = "Sort"
        elif "rtv" in disposition:
            master_df.at[i, 'Material Disposition'] = "RTV"
        elif "scrap" in disposition:
            master_df.at[i, 'Material Disposition'] = "Scrap"
        else:
            master_df.at[i, 'Material Disposition'] = "Use As Is"
           

        # Getting the "DMR Status" column to be in the correct format
        status = master_df.at[i, "DMR Status"].lower()
        if "open" in status:
            master_df.at[i, "DMR Status"] = "Open"
        elif "closed" in status:
            master_df.at[i, "DMR Status"] = "Closed"
        elif "void" in status:
            master_df.at[i, "DMR Status"] = "Void"
        else:
            master_df.at[i, "DMR Status"] = "Monitoring Effectiveness"
        
        # getting the "Problem Type" column to be the right format
        p_type =  master_df.at[i, "Problem Type"].lower()
        if "line" in p_type:
            master_df.at[i, "Problem Detected at"] = "Line"
        elif "incoming" in p_type or "iqa" in p_type:
            master_df.at[i, "Problem Detected at"] = "IQA"
        elif "pilot" in p_type:
            master_df.at[i, "Problem Detected at"] = "Pilot"
        elif "field" in p_type:
            master_df.at[i, "Problem Detected at"] = "Field"
        elif "ooba" in p_type:
            master_df.at[i, "Problem Detected at"] = "OOBA"
        elif "par" in p_type:
            master_df.at[i, "Problem Detected at"] = "PAR"
        else:
            master_df.at[i, "Problem Detected at"] = "Supplier"

        # Attempt at filling out the Supplier Problem Description"  column
        if "cosmetic" in p_type:
            master_df.at[i, "Supplier Problem Description"] = "Cosmetic" 
        else:
            master_df.at[i, "Supplier Problem Description"] = "Other"
        
        
        # start of by getting this rows DMR id:
        dmr_id = str(master_df.loc[[i], "Problem Description Detail"]).split(":")[0]
        dmr_id = dmr_id.split(" ")[-1]
        
        # trying to fix some issues with .0's in the dmr slot
        if ".0" in dmr_id:
            dmr_id = dmr_id.replace(".0","")
        
        
#        dmr_id = temp[0]
        location = str(master_df.loc[[i], "Reporting Location"])
        
        if "LATAM" in location:
            #use d_latam
            for key, val in d_latam.items():
                if dmr_id in key:
                    master_df.at[i,'Receive Date'] = val.dateReceived
                    master_df.at[i,'Containment Action Effectivity Date'] = val.effDate01
                    master_df.at[i,'Corrective Action'] = val.permanentCorr
                    master_df.at[i,'Team (Enter names of team members)'] = val.qualEngineer
                    master_df.at[i,'Quality Engineer'] = val.qualEngineer
                    master_df.at[i,'Revision'] = val.revision
                    master_df.at[i,'Root Cause Description'] = val.rootCause
                    master_df.at[i,'Preventive Action'] = val.tempCorrAction
                
                    # this will record where the information is from
                    master_df.at[i,'Source material'] = key

                    break
          
        else:
            #use d_apac
            for key, val in d_apac.items():
                if dmr_id in key:
                    master_df.at[i,'Original Closure Date'] = val.closure_date
                    master_df.at[i,'Containment Action Effectivity Date'] = val.date_receiv
                    master_df.at[i,'Revision'] = val.m_revision
                    master_df.at[i,'PO/Invoice Number'] = val.po_num
                    master_df.at[i,'Corrective Action'] = val.preventive_action
                    master_df.at[i,'Root Cause Description'] = val.root_cause_description
                    master_df.at[i,'Quality Engineer'] = val.sqe
                    master_df.at[i,'Team (Enter names of team members)'] = val.team
                    
                    # this will record where the information is from
                    master_df.at[i,'Source material'] = key
                    break
    return master_df


# LATAM
# manip takes a dataframe from either plamex or latam
# this cleans and preps the data to then be concatenated with the main dataframe
def plamexManip(data, template):
    transformed_data = data.rename(columns = {'C/A    Status':'DMR Status', 'P/N':'Part Number', 'SQE':'Poly SQE'
        ,'Supplier':'Supplier Name', 'Type': 'Problem Type'
        , 'Disposition': 'Material Disposition', 'ISSUE DATE (month / day / year)': 'Original Creation Date'
        , 'PO #':'PO/Invoice Number', "DMR C/A Closed (month / day / year)":"Original Closure Date"}, inplace = False)

    # creates the column ['Problem Description Detail'] as a combination of DMR and Defect
    transformed_data['DMR '] = transformed_data['DMR '].astype(str)
    transformed_data['Problem Description Detail'] = transformed_data['DMR '] + ": " + transformed_data['Defect']

    transformed_data.drop(['Buyer', 'Qty', 'Qty Inspected', 'Qty Rejected', 'Fault',
                           'RGA / RTV / RMA number', 'Comments ', 'Defect', 'DMR ' ], axis = 1, inplace = True)

    transformed_data.drop(transformed_data.filter(regex='C/A').columns, axis=1, inplace = True)

    # remove empty col and rows that are empty
    transformed_data.dropna(thresh = 5, inplace = True)
    transformed_data.dropna(thresh = 5, axis = 1, inplace = True)

    # add the "Reporting Location" column, and populate it with "LATAM"
    # check out the size of the data and make a column that has that many rows full of "LATAM"
    transformed_data["Reporting Location"] = "LATAM"
    
#    # adding a root cause category, not sure where this info should be found
#    transformed_data['Root Cause Category'] = "Material"
#    
    # This should be kept blank
    transformed_data['Root Cause Category'] = ""



    manip = template.copy()
    manip.dropna(thresh = 5, inplace = True)

    # removes duplicate columns
    manip.drop(transformed_data.columns, axis = 1,inplace =True)

    # concatenate transformed data with mainp to have all the same columns as the template
    return pd.concat([transformed_data,manip], axis = 1, sort = True)


# APAC
# manip takes a dataframe from either plamex or apac
# this cleans and preps the data to then be concatenated with the main dataframe
def apacManip(data, template):
    transformed_data = data.rename(columns = {'Status':'DMR Status', 'P/N':'Part Number', 'Release date':'Receive Date'
        , 'Failure Catelog':'Problem Type', 'SQE':'Poly SQE'
        ,'Supplier':'Supplier Name'
        , 'Model':'Part Description'}, inplace = False)
    
    transformed_data['Original Creation Date'] = transformed_data['Receive Date']
  
    
#    transformed_data['Root Cause Category'] = "Material"  # transformed_data['Problem Type']
    transformed_data['Root Cause Category'] = ""  # transformed_data['Problem Type']




    # creates the column ['Problem Description Detail'] as a combination of DMR and Failure description
    transformed_data['DMR NO.'] = transformed_data['DMR NO.'].astype(str)
    transformed_data['Problem Description Detail'] = transformed_data['DMR NO.'] + ": " + transformed_data['Fail Description']

    transformed_data.drop(['Corrective Action', 'Cycle time', 'Fail\r\n Q\'ty',
                           'Fail \r\nRate', 'Feedback\r\nWeek', 'Feedback Date', 'Feedback Month',
                           'Final Close Date', 'Lot \r\nsize',
                           'Release\r\nMonth', 'Release\r\nWeek', 'Sample\r\n Q\'ty','Root Casue',
                           'DMR NO.','Fail Description'], axis = 1, inplace = True)

    # add the "Reporting Location" column, and populate it with "APAC"
    # check out the size of the data and make a column that has that many rows full of "APAC"
    transformed_data["Reporting Location"] = "APAC"

    # removes columns and rows with very little to no data
    transformed_data.dropna(thresh = 5, inplace = True)
    transformed_data.dropna(thresh = 5, axis = 1, inplace = True)

    manip = template.copy()
    manip.dropna(thresh = 5, inplace = True)

    # removes duplicate columns
    manip.drop(transformed_data.columns, axis = 1,inplace =True)

    return pd.concat([transformed_data,manip], axis = 1, sort = True)


# this function will add the 20 to a the year so that the format is the same everywhere
def reformat_dates(string):
    date = ""
    string = str(string)
    date += string
    num_of_slash = 0
    for i,letter in enumerate(string):
        if "/" == letter:
            num_of_slash +=1
        if num_of_slash == 2 and i+2 < len(string):
            if string[i+1] != "2" and string[i+2] != "0":
                date = string[:i+1] + "20" + string[-2:]
                
    if "2020" in date:
        date = date.replace("2020","20")
        
    # some weird formating happens, this fixes it    
    if len(date) < 10:
        temp = date.split("/")
        for i, x in enumerate(temp):
            if len(x) < 2:
                temp[i] = "0" + x
        date = "/".join(temp)
    
    # change the order to %Y%m%d
    if "nan" not in date:
        temp = date.split("/")
        if len(temp[0]) < 4 and len(temp) == 3:
            temp[0], temp[1], temp[2] = temp[2], temp[0], temp[1]
        date = "/".join(temp)
        
    return date


# After the ordering of the data frame the dates need to be reordered to the format of
# dd-mm-yyyy
def reverse_dates(date):
    if "nan" not in date:
        temp = date.split("/")
        if len(temp) == 3:
            temp = temp[::-1]
        date = "-".join(temp)
        
    return date


# Sets up the date column so that it is the right format and that the whole datafram 
# is sorted by the date in descending order
def date_setup(master_df):
    # do some data clean up, by making the dates into a nice format:
    master_df['Receive Date'] = master_df["Original Creation Date"]
    master_df['Receive Date'].isna().sum()
    master_df['Receive Date'].dropna(inplace=True)

    # this applies the function to every cell
    master_df["Receive Date"] = master_df["Receive Date"].astype(str)
    master_df['Receive Date'] = master_df['Receive Date'].apply(reformat_dates)
    
    # this orders the entire dataframe by the date from most recent to least
    master_df.sort_values(by = ['Receive Date'], inplace=True, ascending=False)
    
    # now to reorder the dates in the "Receive Date"
    master_df["Receive Date"] = master_df["Receive Date"].apply(reverse_dates)

    
    # now to take all of the columns that contain dates and replace "/" with "-"
    # 'Receive Date', 'Containment Action Effectivity Date', 'Corrective Action Effectivity Date',
    # 'Supplier Action Close Date', 'Original Creation Date', 'Original Closure Date'
    # and reformat them to dd-mm-yyyy
    def replace_f(date):
        if type(date) is str and "nan" not in date:
            if "/" in date:    
                temp = date.split("/")
            elif "." in date:
                temp = date.split(".")
            elif "," in date:
                temp = date.split(",")
            else:
                temp = date.split("-")
                if len(temp) == 3:
                    temp[0], temp[1], temp[2] = temp[1], temp[0], temp[2]
                date = "-".join(temp)
        return date
        
    master_df["Containment Action Effectivity Date"] = master_df["Receive Date"]
    master_df["Corrective Action Effectivity Date"] = master_df["Corrective Action Effectivity Date"].apply(replace_f)
    master_df["Supplier Action Close Date"] = master_df["Supplier Action Close Date"].apply(replace_f)
    master_df["Original Creation Date"] = master_df["Original Creation Date"].apply(lambda x: x.replace("/","-"))
    master_df["Original Creation Date"] = master_df["Original Creation Date"].apply(replace_f)
    master_df["Original Closure Date"] = master_df["Receive Date"]
    
    return master_df




# Main combines all the data together to get the finished product
def main():
    # the template that needs filling
    template = pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/Historical_DMR_Template.csv")

    plamex_data = [pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-Plamex-MASTER-2015.csv"),
                   pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-Plamex-MASTER-2016.csv"),
                   pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-Plamex-MASTER-2017.csv"),
                   pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-Plamex-MASTER-2018.csv"),
                   pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-Plamex-MASTER-2019.csv"),
                   pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-MASTER-2019_02.csv")]

    apac_data = [pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-APAC-2015.csv"),
                 pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-APAC-2016.csv"),
                 pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-APAC-2017.csv"),
                 pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-APAC-2018.csv"),
                 pd.read_csv("C:/Users/frubino/DMR Project/DMR_Scripts/DMR-APAC-2019.csv"),]

    # this will create a list of prepped dataframes that
    # all have the same columns to be concatenated

    list_of_prepped_data = list()

    for i, df in enumerate(plamex_data):
        print("Plamex:",i)
        list_of_prepped_data.append(plamexManip(df,template))

    for i, df in enumerate(apac_data):
        print("APAC",i)
        list_of_prepped_data.append(apacManip(df,template))

    # combine all into a mega dataframe!!!
    master_df = pd.concat(list_of_prepped_data, sort = True)
    print("Number of rows:", master_df.index.size)
    print("Number of Columns:", master_df.columns.size)
    
    # clean up the index
    master_df.reset_index(inplace=True)
    master_df.drop(['index'],axis=1, inplace=True)  # it creates a column called index, not needed

    master_df = populate_dataframe(master_df)

    master_df = date_setup(master_df)
    
    # Now that all of the data has been sorted by date in descending order all
    # we need to do is add the columns 'Defective Material Report: ID',
    # 'Defective Material Report: Defective Material Report Name'
    
    # this is supposed to get me only 1000
    master_df = master_df[0:1000]
    
    master_df.reset_index(inplace=True)
    master_df.drop(['index'],axis=1, inplace=True)  # it creates a column called index, not needed
    
    # I have to reverse template index
    
    # this is a column for the source material dmrs
    template["Source material"] = ""
    
    template.sort_values(by = ['Defective Material Report: Defective Material Report Name'], inplace=True, ascending=False)
    template.reset_index(inplace=True)
    template.drop(['index'],axis=1, inplace=True)  # it creates a column called index, not needed
    
    
    master_df['Defective Material Report: ID'] = template['Defective Material Report: ID']
    master_df['Defective Material Report: Defective Material Report Name'] = template['Defective Material Report: Defective Material Report Name']
    
    master_df["Original Creation Date"] = master_df["Original Closure Date"]
    master_df["Original Closure Date"] = ""
    
    # now to reorder the columns of master_df to match the columns of the template
    master_df = master_df[template.columns]
    
    master_df["Supplier Problem Description"] = ""
        
    print(master_df.info())
    print(master_df["Receive Date"])
    print(master_df["Containment Action Effectivity Date"])


    # finally write the new and finished dataframe back as a csv file:
    master_df.to_excel('legacy_dmrs_for_salesforce.xlsx')


if __name__ == '__main__':
    main()
