VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRegression 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "UserFormRegression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButtonRegression_Enter()
    
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Dim xs() As String
    Dim i As Integer
    
    Worksheets(1).Range(Worksheets(1).Cells(1, Int(TextBoxY.Value)), Worksheets(1).Cells(1, Int(TextBoxY.Value)).End(xlDown)).Copy
    newSheet.Cells(1, Int(TextBoxY.Value) + 10).PasteSpecial
    
    xs = Split(TextBoxXs.Value, ",")
    
    For i = 0 To UBound(xs)
        Worksheets(1).Range(Worksheets(1).Cells(1, Int(xs(i))), Worksheets(1).Cells(1, Int(xs(i))).End(xlDown)).Copy
        newSheet.Cells(1, 12 + i).PasteSpecial
    Next
    
    If UBound(xs) <> 0 Then
        Application.Run "ATPVBAEN.XLAM!Regress", newSheet.Range(newSheet.Cells(1, 11), newSheet.Cells(1, 11).End(xlDown)), _
            newSheet.Range(newSheet.Cells(1, 12), newSheet.Cells(1, 12).End(xlDown).End(xlToRight)), False, True, , newSheet.Range("A1") _
            , False, False, False, False, , False
        ActiveWindow.SmallScroll Down:=-3
    Else
        Application.Run "ATPVBAEN.XLAM!Regress", newSheet.Range(newSheet.Cells(1, 11), newSheet.Cells(1, 11).End(xlDown)), _
            newSheet.Range(newSheet.Cells(1, 12), newSheet.Cells(1, 12).End(xlDown)), False, True, , newSheet.Range("A1") _
            , False, False, False, False, , False
        ActiveWindow.SmallScroll Down:=-3
    End If
    
    newSheet.Range("A1").Select
    
End Sub


