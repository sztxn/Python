' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 13.0.0
' 4:15:11 下午  六月 11, 2018
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
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateEquationCurve Array("NAME:EquationBasedCurveParameters", "XtFunction:=",  _
  "x0+(x1-x0)*_t", "YtFunction:=", "y0+(y1-y0)*_t", "ZtFunction:=", "0", "tStart:=",  _
  "0", "tEnd:=", "1", "NumOfPointsOnCurve:=", "10", "Version:=", 1), Array("NAME:Attributes", "Name:=",  _
  "part1", "Flags:=", "", "Color:=", "(255 0 0)", "Transparency:=", 0, "PartCoordinateSystem:=",  _
  "", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateEquationCurve Array("NAME:EquationBasedCurveParameters", "XtFunction:=",  _
  "x1+(x2-x1)*_t", "YtFunction:=", "y1+(y2-y1)*_t", "ZtFunction:=", "0", "tStart:=",  _
  "0", "tEnd:=", "1", "NumOfPointsOnCurve:=", "10", "Version:=", 1), Array("NAME:Attributes", "Name:=",  _
  "part2", "Flags:=", "", "Color:=", "(255 0 0)", "Transparency:=", 0, "PartCoordinateSystem:=",  _
  "", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "" & Chr(34) & "", "SolveInside:=",  _
  false)
