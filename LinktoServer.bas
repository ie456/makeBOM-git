Attribute VB_Name = "LinktoServer"
Sub LinktoServer()

    Dim testPath As String
    
    
    
    
   
    
    If Application.UserName = "Brian Lin (ªL¤h¶v)" Then
        UserForm3.show
        Exit Sub
    Else
        On Error GoTo Exception1
        testPath = "\\10.242.184.38\Department\EE_Arthur\Public_access\04_Technical document\05_Tools\Excel\Data\testLink.txt"
        
        mf = FreeFile
        
        Open testPath For Input As #mf
        
        
        Do Until EOF(1)
        
            Line Input #1, textline
             If textline = "BU9_Arthur" Then GoTo EndPoint
    
        Loop
    End If
    
    
    
    
EndPoint:
    Call Start.Start
    Exit Sub
    
    
    
    
Exception1:
    MsgBox "Error: Can not Link to Server" & vbCrLf & "Please check you can link to 'EE_Arthur' folder"
    Exit Sub





End Sub

