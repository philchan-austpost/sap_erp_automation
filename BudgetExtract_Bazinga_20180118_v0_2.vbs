If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 214,32,false
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00002"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00002"
session.findById("wnd[0]/usr/lbl[5,4]").setFocus
session.findById("wnd[0]/usr/lbl[5,4]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[9,7]").setFocus
session.findById("wnd[0]/usr/lbl[9,7]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[13,11]").setFocus
session.findById("wnd[0]/usr/lbl[13,11]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[20,13]").setFocus
session.findById("wnd[0]/usr/lbl[20,13]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/lbl[1,3]").caretPosition = 4
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtPAR_05").setFocus
session.findById("wnd[0]/usr/ctxtPAR_05").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "J-00015-O-4-2202"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "J-00013-O-2-1001"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0-00587-C-4-2107"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "J-00014-O-2-1004"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "J-00015-O-4-2205"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "J-00015-O-4-1112"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "J-00015-O-4-1406"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "0-00632-C-4-1105"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[47]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Users\chank\Desktop"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "RawExtract_Bazinga_20180118_v0_1.dat"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 36
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[2]/usr/btnSPOP-VAROPTION1").press
