' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 13.0.0
' 3:51:42 下午  六月 11, 2018
' ----------------------------------------------
Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.SetActiveProject("Project3")
Set oDesign = oProject.SetActiveDesign("HFSSDesign1")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:xx", "PropType:=", "VariableProp", "UserDef:=",  _
  true, "Value:=", "[0, 1, 3, 3, 4, 5]"), Array("NAME:yy", "PropType:=", "VariableProp", "UserDef:=",  _
  true, "Value:=", "[0, 1, 2, 3, 4, 5]"), Array("NAME:x0", "PropType:=", "VariableProp", "UserDef:=",  _
  true, "Value:=", "xx[0]"), _
  Array("NAME:y0", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "yy[0]"), _
  Array("NAME:x1", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "xx[1]"), _
  Array("NAME:y1", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "yy[1]"), _
  Array("NAME:x2", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "xx[2]"), _
  Array("NAME:y2", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "yy[2]"), _
  Array("NAME:x3", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "xx[3]"), _
  Array("NAME:y3", "PropType:=", "VariableProp", "UserDef:=",true, "Value:=", "yy[3]") _
  )))
