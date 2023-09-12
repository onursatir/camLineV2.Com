dim myObj
dim err
dim setting
dim settingsAsJSON
dim test
dim test1
dim test2
dim test3
dim test4
dim unitstatus
dim unitstatus2(0)

Set myObj = CreateObject("camLineV2.Com.camLineV2ComWrapper")


test4 = myObj.test(unitstatus2)


test = myObj.Init ("TEST", err, settingsAsJSON)
MsgBox test


test2 = myObj.EQP_CheckUnit("100309300107102301600330", "ICAS_SEMI1", unitstatus2)

MsgBox test2
MsgBox unitstatus2

test3 = myObj.EQP_CheckStationStatus( err, settingsAsJSON ) 
MsgBox test3