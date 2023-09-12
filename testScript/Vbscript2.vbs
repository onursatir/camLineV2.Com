dim myObj
dim unitstatus
dim test
dim test2
dim test3
dim unit(0) 
unit(0)= "100309300107102301600330"
dim statuts
dim parameter

Set myObj = CreateObject("Continental.camLineV2ComWrapper")

test = myObj.Init ("TEST", err, settingsAsJSON)
MsgBox test


'test2 = myObj.EQP_CheckUnit("",unit,stauts)
'MsgBox test2
'MsqBox status


test3 = myObj.EQP_GetParameter("18414301018243323041","MRA2_SN","p_pdi_hsm_cmac",parameter,status)
MsgBox test3
MsgBox parameter