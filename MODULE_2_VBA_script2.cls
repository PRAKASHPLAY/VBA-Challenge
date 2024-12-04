VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub reset()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    ws.Activate


ws.Range("I:ZZ").Delete

Worksheets("Q1").Select

Next ws

End Sub
