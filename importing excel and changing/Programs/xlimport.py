import pandas as pd
import openpyxl
import os
import numpy as np

# creating paths like libname statment in SAS to easily use in the programs like input folder and ouptput folder

# created a setup folder and created a file called Paths_prog.py to create paths for input, output, programs and setup folder

# since the Paths_prog.py is in setup folder, we need to import the paths from that file to use in this program# 
# but the problem is paths_prog.py is in setup folder and this program is in Programs folder so we can't run like python xlimport from Pragrams directory because pythopn cant acces program from a parent directory so we have to go back to the previous directory which is the directory where programs and setup is present and type python -m Programs.xlimport to run the program from the parent directory


from setup.Paths_prog import inputfdl, outputfdl, prog_fdl, setupfdl

#print(inputfdl)

form=pd.read_excel(inputfdl / "sample_study_ALS.xlsx", sheet_name="Forms")
form_fin=form[["OID", "FormName"]]
#print(form_fin)

field=pd.read_excel(inputfdl / "sample_study_ALS.xlsx", sheet_name="Fields")
field_fin=field[["FormOID","FieldOID", "DataDictionaryName", "UnitDictionaryName","ControlType", "IndentLevel", "PreText", "FixedUnit", "IsRequired"]]
#print(field_fin)



mer1=pd.merge(field_fin, form_fin, left_on="FormOID",right_on="OID", how="left" )

#print(mer1)

dictxl=pd.read_excel(inputfdl / "sample_study_ALS.xlsx", sheet_name="DataDictionaryEntries")
dictxl_fin=dictxl[["DataDictionaryName", "UserDataString","Ordinal"]]
#print(dictxl_fin)
dictxl_fin["Ordinal"] = "col-" + dictxl_fin["Ordinal"].astype(str)
codval=dictxl_fin.pivot(
    index='DataDictionaryName',
    columns='Ordinal',
    values='UserDataString'
).reset_index()
#print(codval)
# codval.to_excel(
#     r"D:\vs progs\import excel\test_res_exp\check_this.xlsx",
#     index=False
# )

mer2=pd.merge(mer1,codval,on='DataDictionaryName',how='left')

ct_lst= [col for col in mer2.columns if col.startswith("col-")]

mer2["DictionaryValues"] = mer2[ct_lst].values.tolist()

#mer_chk=
mer2.to_excel(
    outputfdl / "check_this.xlsx",
    index=False
)


mer2.to_json( outputfdl / "output.json",orient="records", indent=4 )
