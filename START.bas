Attribute VB_Name = "START"

Public get_FileName_MakeBOM
Public get_WorkBookNamePath
Sub GetBookName(Name As String)
    get_FileName_MakeBOM = Name
End Sub

Sub Start()

    
    
    'tempWorkBookName = ActiveWorkbook.Name
    
    get_FileName_MakeBOM = ActiveWorkbook.Name
    get_WorkBookNamePath = ActiveWorkbook.Path
    
    UserForm1.show vbModeless
    
    If Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B22") = 0 Then
     Call subModule.initSheet
    End If
    
    functionModule.updateUserFormValue 0

    
    
End Sub


