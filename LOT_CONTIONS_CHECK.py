import pandas as pd
import glob
import numpy as np

old_link = input("Please_Share_Hawb_No.- ")

new_link = glob.glob(fr"\\192.168.1.10\Status\BINDU MAM\1.SAM LOT DATA\*{old_link}*.xlsb")

if not new_link:
    print("File not found")
    exit()
new_link = new_link[0]
Data1 = pd.read_excel(
    new_link,
    sheet_name="XLS_MIRROR_DATA",
    skiprows=10,
    usecols=["Inv No","Product Desc","CTH"],
    nrows=1000,
    engine="pyxlsb"
)

cond1 = (
    Data1["Product Desc"].str.contains("resistor", case=False, na=False) &
    (Data1["CTH"] == 85334030)
)
cond2 = Data1["Product Desc"].str.contains(
    "JDM-SMT RESISTOR|IC-POLYMER|IC-MICROCONTROLLER|JDM-FILTER|JDM-DUPLEXER|JDM-RF FILTER",
    case=False,
    na=False
)
cond3 = (
    Data1["Product Desc"].str.contains("SM-X|SM-T", case=False, na=False) &
    Data1["Product Desc"].str.contains("MOBILE", case=False, na=False)
)






Data1["Reason"] = np.select(
    [cond1, cond2, cond3],
    [
        "REV REQ. Resistor Mismatch with CTH",
        "Add letter in esanchit of given Items",
        "SM-T and Mobile phone in Same Desription"
    ],
    default=""
)








filterd_Data = Data1[Data1["Reason"] != ""].drop_duplicates(
    subset="Product Desc", keep="first"
)

if filterd_Data.empty:
    print(f"Hawb No. {old_link} Data Is OK")
else:
    print("Error List send to Desktop with Hawb no.")
    filterd_Data.to_excel(
        fr"C:\Users\Girish\Desktop\{old_link}.xlsx",
        index=False
    )





