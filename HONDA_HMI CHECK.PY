import pandas as pd
import numpy as np
BRIGHT_GREEN   = "\033[92m"
RESET = "\033[0m"
BRIGHT_RED     = "\033[91m"
#-------------------Taking User Input--------------
job_no=int(input("Enter The Job No. To check.-- "))
job_link=fr"C:\Users\Girish\Downloads\{job_no}.xlsx"

#-------------------Required_Columns--------------
Req_col=["Job No","Inv No","Inv Date","Invoice Value","Freight Amount","Fre Curnc","Ins Curnc","Ins Amt","Sequence Number",
         "Material Description","Material Quantity","Unit Price","Cth No","Bcd Notn","Bcd Sno","Social Welfare Rta",
         "Igst Sno","Aidc Sno","Sqc Qty","Sqc Unit","Cntry Orig","End Use"]

#-------------------Summary of Job--------------
Summary_excel=pd.read_excel(job_link,usecols=["Job No","Inv No","Inv Date","Invoice Value","Currency",
                                              "Freight Amount","Fre Curnc","Ins Curnc","Ins Amt"]).drop_duplicates(subset="Inv No",keep="first")
print(f'{BRIGHT_GREEN}{Summary_excel}{RESET}')


#------------Calculating_Total_Freight-------------
freight_by_currency = Summary_excel.groupby("Fre Curnc",as_index=False)["Freight Amount"].sum()
print(f"{BRIGHT_RED}{freight_by_currency}{RESET}")

#----------------Clean_Columns-------------

def clean_columns(df,cols):
    for col in cols:
        df[col]=df[col].fillna("NA").astype(str).str.strip()
    return df

#--------Cth_Master_link=r"\\192.168.1.10\Status\2.E-Way_Bill_SEA\1 GIRISH-------HMI--------PAPERS\HSN MASTER applicable updated  23 11 2024 onwards..xlsx"
Cth_Master_link=r"\\192.168.1.10\Status\2.E-Way_Bill_SEA\1 GIRISH-------HMI--------PAPERS\HSN MASTER applicable updated  23 11 2024 onwards..xlsx"
Cth_Master_link=pd.read_excel(Cth_Master_link,sheet_name="HMI CTH",usecols=["PART NO","DESCRIPTION","HSN Code"]).drop_duplicates(subset="PART NO",keep="first")
job_data=pd.read_excel(job_link,sheet_name="Product Jobwise",usecols=["Material Description","Cth No","Inv No","Sequence Number"])
job_data["U_code"]=job_data["Material Description"].astype(str).str.split(" ").str[0]

#------------------------------Checking_CTH------------------------#
Cth_check=pd.merge(job_data,Cth_Master_link,left_on="U_code",right_on="PART NO",how="left")

Cth_check2=Cth_check[Cth_check["Cth No"]
                    !=Cth_check["HSN Code"]]

if not Cth_check2.empty:
    print(Cth_check2)
    Cth_check2.to_excel(fr"C:\Users\Girish\Desktop\{job_no} ERROS.xlsx")
    
else:
    print("------------------------CTH CHECKED\n----------------------CTH IS OK")
    pass
#----------------------cth checked------------------------
#----------------------BASIC NOT,IGST,SQC CHECK--------------------

columns_to_clean=["BN-1","BN-2","GST-2","UQC","Bcd Notn","Bcd Sno","Igst Sno","Sqc Unit"]

def apply_checks(df):
    conditions = [
        df["BN-1"] != df["Bcd Notn"],
        df["BN-2"] != df["Bcd Sno"],
        df["GST-2"] != df["Igst Sno"],
        df["UQC"] != df["Sqc Unit"]
    ]
    choices=["BN1 WRONG", "BN2 WRONG", "IGST WRONG", "UQC WRONG"]   
    
    
    df["Error"]=np.select(conditions,choices,default="OK")
    return df

auto_power_link=pd.read_excel(job_link,sheet_name='Product Jobwise',usecols=["Material Description","Cth No","Bcd Notn","Bcd Sno","Igst Sno","Sqc Unit"])

auto_basic_job=auto_power_link[auto_power_link["Material Description"].str.contains("Automobile",case=False,na=False)]
power_basic_job=auto_power_link[auto_power_link["Material Description"].str.contains("power product",case=False,na=False)]

#------------------MASTER BASIC OF POWER AND AUTOMOBILE LINK---------------------
Master_auto=pd.read_excel(r"\\192.168.1.10\Status\2.E-Way_Bill_SEA\1 GIRISH-------HMI--------PAPERS\HMI HS Code Master _14.02.2020-HSN CODE 22-09.xlsx",sheet_name="AUTO",skiprows=1,usecols=["HS Code for filing BOE","UQC","BN-1","BN-2","GST-2"])
Master_power=pd.read_excel(r"\\192.168.1.10\Status\2.E-Way_Bill_SEA\1 GIRISH-------HMI--------PAPERS\HMI HS Code Master _14.02.2020-HSN CODE 22-09.xlsx",sheet_name="POWER",skiprows=1,usecols=["HS Code for filing BOE","UQC","BN-1","BN-2","GST-2"])


Master_auto_merged=pd.merge(auto_basic_job,Master_auto,left_on="Cth No",right_on="HS Code for filing BOE",how="left")
Master_power_merged=pd.merge(power_basic_job,Master_power,left_on="Cth No",right_on="HS Code for filing BOE",how="left")

cols_to_clean = ["BN-1","BN-2","GST-2","UQC","Bcd Notn","Bcd Sno","Igst Sno","Sqc Unit"]

auto_merged=clean_columns(Master_auto_merged,cols_to_clean)
power_merged=clean_columns(Master_power_merged,cols_to_clean)

auto_merged=apply_checks(auto_merged)
power_merged=apply_checks(power_merged)

auto_merged["Type"] = "AUTOMOBILE"
power_merged["Type"] = "POWER"


final_df = pd.concat([auto_merged, power_merged], ignore_index=True)

errors_df = final_df[final_df["Error"] != "OK"]

if not errors_df.empty:
    print(errors_df)
    print("send to downloads with job no. ")
    errors_df.to_excel(fr"C:\Users\Girish\Downloads\{job_link}.xlsx")
else:
    print(BRIGHT_GREEN+"Basic notification igst uqc checked everyting is ok"+RESET)

#---------------------------checking country of origin-------------------------
IDATA=pd.read_excel(job_link,sheet_name="Product Jobwise",usecols=["Inv No","Cth No","Material Quantity","Social Welfare Notn","Social Welfare Sno",
                                                                       "Social Welfare Rta","Aidc Notn","Aidc Sno","Sqc Qty","Sqc Unit","Cont Of Orig"])

#---------------------Checking 73 cth 48, and 76 IF exists then give date-------------------------

CTH_DATA=IDATA.drop_duplicates(subset=["Cont Of Orig"],keep="first")
print("---------------Country of Origin Data---------------\n",CTH_DATA[["Cont Of Orig","Inv No"]])


def check_imp_conditons(df):
    conditions=[(
        (df["Cth No"].astype(str).str.startswith("94")))
        &
        ((df["Aidc Sno"]!='16D')
        |
        (df["Social Welfare Rta"]!=0)),
# condtion 2
        ((df["Cth No"].astype(str).str.startswith("64")))
        &
        ((df["Aidc Sno"]!="14A")
        |
        (df["Social Welfare Rta"]!=0)),
#condiiton 3
        (df["Sqc Unit"]=="NOS")
        &
        (df["Sqc Qty"]!=df["Material Quantity"])
        ]
    choices=["Error in Cth startwith 94",
             "Error in CtH startswith 64",
             "NOS mismatch"]
    df["Error"]=np.select(conditions,choices,default="ok")
    return df



chechcth94=check_imp_conditons(IDATA)

chechcth94=chechcth94[chechcth94["Error"]!="ok"]

if not chechcth94.empty:
    print("-------------------------------Checking 94 64 cth and sqc")
    print(chechcth94)
else:
    print("CTH 94 and 64 with UQC checked ok pass")

    pass
#--------------------------NOW SIMS, NFMIMS, AND PIMS CHECKING-----------------------------

DATAC=IDATA.drop_duplicates(subset=["Cth No"],keep="first")
DATAC=DATAC[["Cth No"]]

Filtered_data_CONDITON=[
    (DATAC["Cth No"].astype(str).str.startswith("73")),
    (DATAC["Cth No"].astype(str).str.startswith("76")),
    (DATAC["Cth No"].astype(str).str.startswith("48"))
                        ]
choices=["SIMS REQ.","NFMIMS REQ.","PIMS REQ."]

DATAC["Errors"]=np.select(Filtered_data_CONDITON,choices,default="OK")

filtered_c=DATAC[DATAC["Errors"]!="OK"]

if not filtered_c.empty:
    print(filtered_c)
else:
    print("NO SIMS, NFMIMS, PIMS REQUIRED")


#------------------------------  DES CONTAINS #N/A ---------------------------
datal=pd.read_excel(job_link,usecols=["Inv No","Material Description"])

datal1=datal[datal["Material Description"].str.contains("#N/A",case=False,na=False)]

if not datal1.empty:
    print("--------------------------------N/a in Descripton")
    print(datal1)
else:
    pass
#---------------------------- CHECKING IF INV NO. CONTAINES NO MORE THAN 1 SAME END USE DES.
Inv_enduse_check=pd.read_excel(job_link,usecols=["Material Description","Inv No"])
Inv_enduse_check["End use des"]=Inv_enduse_check["Material Description"].astype(str).str.extract(r'(\([^()]*\))$')
inInv_enduse_check=Inv_enduse_check.drop_duplicates(subset=["Inv No","End use des"],keep="first",inplace=True)
print(Inv_enduse_check[["Inv No","End use des"]])








