VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo ErrHandler
    
    If Not Intersect(Target, Range("DataRange")) Is Nothing Then
        Application.EnableEvents = False
        Application.OnTime Now, "fillTable"
        Debug.Print "DataRange has been changed"
    End If

ErrHandler:
    Application.EnableEvents = True
End Sub


