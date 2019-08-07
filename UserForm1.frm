VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5244
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7416
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public step








'Insert to Lv3

Private Sub CommandButton20_Click()
    
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.Selected(i) Then
        
            'If ListBox3.ListCount = "" Then
            '    indexRow = 0
            'Else
                indexRow = ListBox3.ListCount
            'End If
            
            ListBox3.AddItem
            ListBox3.List(indexRow, 0) = ListBox1.List(i, 0)
            ListBox3.List(indexRow, 1) = ListBox1.List(i, 1)
        End If
    
    Next
    
    counter = 0
    
    For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i - counter) Then
        ListBox1.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
    
    'For i = ListBox1.ListCount - 1 To 0 Step -1
'
'        If ListBox1.Selected(i) = True Then
'            ListBox1.RemoveItem
'
'            Exit For
'        End If
    
'    Next
End Sub

'Lv3 to NeedAssign

Private Sub CommandButton21_Click()
    For i = 0 To ListBox3.ListCount - 1
        
        If ListBox3.Selected(i) Then
        
            ListBox1.AddItem
            ListBox1.List(ListBox1.ListCount - 1, 0) = ListBox3.List(i, 0)
            ListBox1.List(ListBox1.ListCount - 1, 1) = ListBox3.List(i, 1)
        End If
    
    Next
    

    counter = 0
    For i = 0 To ListBox3.ListCount - 1
    If ListBox3.Selected(i - counter) Then
        ListBox3.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
    
End Sub

Private Sub CommandButton22_Click()
    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.Selected(i) Then
        
            'If ListBox3.ListCount = "" Then
            '    indexRow = 0
            'Else
                indexRow = ListBox4.ListCount
            'End If
            
            ListBox4.AddItem
            ListBox4.List(indexRow, 0) = ListBox1.List(i, 0)
            ListBox4.List(indexRow, 1) = ListBox1.List(i, 1)
        End If
    
    Next
    
     counter = 0
    For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i - counter) Then
        ListBox1.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
    
    
End Sub

Private Sub CommandButton23_Click()
    For i = 0 To ListBox4.ListCount - 1
        
        If ListBox4.Selected(i) Then
        
            ListBox1.AddItem
            ListBox1.List(ListBox1.ListCount - 1, 0) = ListBox4.List(i, 0)
            ListBox1.List(ListBox1.ListCount - 1, 1) = ListBox4.List(i, 1)
        End If
    
    Next
    
     counter = 0
    For i = 0 To ListBox4.ListCount - 1
    If ListBox4.Selected(i - counter) Then
        ListBox4.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
    
End Sub

Private Sub CommandButton24_Click()
        For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.Selected(i) Then
        
            'If ListBox3.ListCount = "" Then
            '    indexRow = 0
            'Else
                indexRow = ListBox5.ListCount
            'End If
            
            ListBox5.AddItem
            ListBox5.List(indexRow, 0) = ListBox1.List(i, 0)
            ListBox5.List(indexRow, 1) = ListBox1.List(i, 1)
        End If
    
    Next
    
    counter = 0
    
    For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i - counter) Then
        ListBox1.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
End Sub

Private Sub CommandButton25_Click()
    For i = 0 To ListBox5.ListCount - 1
        
        If ListBox5.Selected(i) Then
        
            ListBox1.AddItem
            ListBox1.List(ListBox1.ListCount - 1, 0) = ListBox5.List(i, 0)
            ListBox1.List(ListBox1.ListCount - 1, 1) = ListBox5.List(i, 1)
        End If
    
    Next
    
    counter = 0
    
    For i = 0 To ListBox5.ListCount - 1
    If ListBox5.Selected(i - counter) Then
        ListBox5.RemoveItem (i - counter)
        counter = counter + 1
    End If
    Next i
End Sub

Private Sub CommandButton34_Click()

    Call saveData("MAIN", "B", "B34")
    Call saveData("MAIN", "D", "B35")
    Call saveData("MAIN", "E", "B36")
    Call saveData("MAIN", "G", "B37")
    Call saveData("MAIN", "N", "B38")
    Call saveData("MAIN", "O", "B39")
    Call saveData("MAIN", "8", "B40")
   functionModule.updateUserFormValue 5
End Sub


Private Sub CommandButton36_Click()
    Call subModule.saveSetData_col_para
End Sub

Private Sub CommandButton38_Click()  'create BOM
    Call mainModule.createBOM
End Sub

Private Sub CommandButton39_Click()
    Call NextPage
End Sub

Private Sub CommandButton40_Click()
       ' On Error GoTo exit1
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Select a Folder for saving file"
        .ButtonName = "Select"
        .InitialFileName = get_WorkBookNamePath
        .AllowMultiSelect = False
        
        If .show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
            
        End If
    End With
    
    If sFolder <> "" Then
         'temp = InputBox("Please key-in file name.", "SET FILE NAME", "DEFAULT")
         'If temp = "DEFAULT" Then
             FileName = "BOM_" & Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B24").Value & "_" & Format(Now, "yyyymmdd_hhmm")
         'Else
         '    fileName = temp
         'End If
         
         Path = sFolder & "\" & FileName
         
         If Dir(Path, vbDirectory) = vbNullString Then
             VBA.FileSystem.MkDir (Path)
         End If
         
        tempfileName = functionModule.printOut_dir(FileName, Path)
        Call Module1.printOut_lv(tempfileName, Path)
    Else
    
    End If
    
    
    
    

    
    
End Sub

Private Sub CommandButton41_Click()
    tempfilter = "Component file (*.xls;*.xlsx;*.csv), *.xls;*.xlsx;*.csv"
    TextBox30.Text = functionModule.getPath_Gen("MAIN", "E29", "E30", 23, tempfilter)
End Sub

Private Sub CommandButton49_Click()
    Call PreviousPage
End Sub

Private Sub CommandButton50_Click()
 Call Not_myCode.btnGenerateComponentCSV_Click
End Sub

Private Sub CommandButton8_Click()
  TextBox7.Text = functionModule.getPath("MAIN", "B31", "B32", 2)
End Sub










''''''''''''''''''''''''''
'''''''Form Init '''''''''
''''''''''''''''''''''''''


Private Sub UserForm_Initialize()

    step = Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B22").Value
    

    

    
    
    temp = functionModule.updateUserFormValue(0)
    
    If TextBox1.Text = "Please Enter lv3 PartNumber" Then
      TextBox1.ForeColor = &H808080
    End If
    If TextBox2.Text = "Please Enter lv4 PartNumber" Then
      TextBox2.ForeColor = &H808080
    End If
    If TextBox3.Text = "Please Enter lv5 PartNumber" Then
      TextBox3.ForeColor = &H808080
    End If
    If TextBox4.Text = "Please Enter PCB PartNumber" Then
      TextBox4.ForeColor = &H808080
    End If
    
    CommandButton1.SetFocus
    
End Sub


''''''''''''''''''''''''''
''''''Page Conection '''''
''''''''''''''''''''''''''


'page1
Private Sub CommandButton1_Click()
    
    
    'If Replace(Me.TextBox5.Value, " ", "") = "" Then
    '    MsgBox "Please select component file"
    '    Exit Sub
    'End If
    
    'If Replace(Me.TextBox7.Value, " ", "") = "" Then
    '    MsgBox "Please select BOM file"
    '    Exit Sub
    'End If
    UserForm1.CommandButton1.Enabled = False
    
    If Replace(Me.TextBox5.Value, " ", "") = "" Then
        MsgBox "Please select component file"
        Exit Sub
    ElseIf Dir(Me.TextBox5.Value, vbDirectory) = vbNullString Then
        MsgBox "Path Not Exist: " & vbCrLf & Me.TextBox5.Value
        Exit Sub
        
    End If
    
    
    If Replace(Me.TextBox7.Value, " ", "") = "" Then
        MsgBox "Please select BOM file"
        Exit Sub
    ElseIf Dir(Me.TextBox7.Value, vbDirectory) = vbNullString Then
        MsgBox "Path Not Exist: " & vbCrLf & Me.TextBox7.Value
        Exit Sub
        
    End If
    
    
    
    
    
    Call subModule.initSheet
    Call mainModule.load
    Call mainModule.load_Gene("Component", UserForm1.TextBox5.Value, UserForm1.ComboBox2.Text)
    Call mainModule.get_brdLocation_csv("Component")
    'If Replace(Me.TextBox1.Value, " ", "") = "" Or Me.TextBox1.Value = "Please Enter lv3 PartNumber" Then
    '    MsgBox "Please fill out lv3 PartNumber"
    '    Exit Sub
    'End If
    
    'If Replace(Me.TextBox4.Value, " ", "") = "" Or Me.TextBox4.Value = "Please Enter PCB PartNumber" Then
     '   MsgBox "Please Enter PCB PartNumber"
     '   Exit Sub
    'End If
    'Please Enter PCB PartNumber
    
    
    UserForm1.CommandButton1.Enabled = True
    Workbooks(get_FileName_MakeBOM).Sheets("BOM").Select
    
    
    Call NextPage
End Sub


'page2
Private Sub CommandButton2_Click()


 
    
    
     If Replace(Me.TextBox1.Value, " ", "") = "" Or Me.TextBox1.Value = "Please Enter lv3 PartNumber" Then
        MsgBox "Please fill out lv3 PartNumber"
        Exit Sub
     End If
     
     If Me.TextBox2.Visible Then
        If Replace(Me.TextBox2.Value, " ", "") = "" Or Me.TextBox2.Value = "Please Enter lv4 PartNumber" Then
           MsgBox "Please fill out lv4 PartNumber"
           Exit Sub
        End If
     End If
     
     If Me.TextBox3.Visible Then
        If Replace(Me.TextBox3.Value, " ", "") = "" Or Me.TextBox3.Value = "Please Enter lv5 PartNumber" Then
           MsgBox "Please fill out lv5 PartNumber"
           Exit Sub
        End If
     End If
     
     If Replace(Me.TextBox4.Value, " ", "") = "" Or Me.TextBox4.Value = "Please Enter PCB PartNumber" Then
        MsgBox "Please Enter PCB PartNumber"
        Exit Sub
     End If
     
     
   
    Call NextPage

End Sub
Private Sub CommandButton3_Click()
    Call PreviousPage
End Sub

'page2-1
Private Sub CommandButton4_Click()
    Call NextPage
End Sub
Private Sub CommandButton5_Click()
    Call PreviousPage
End Sub

'page2-2
Private Sub CommandButton28_Click()
    
    
    Dim data_PN, data_QTY, data_Loc
    
    

    Set data_PN = CreateObject("system.collections.arraylist")
    Set data_QTY = CreateObject("system.collections.arraylist")
    Set data_Loc = CreateObject("system.collections.arraylist")
    
    Call mainModule.classifi(data_PN, data_QTY, data_Loc)
    Call printInListBOX(data_PN, data_Loc, data_QTY)
    
    Call NextPage
    
    
    
End Sub

Private Sub CommandButton29_Click()
    Call PreviousPage
End Sub

'page4
Private Sub CommandButton15_Click()
    Call PreviousPage
End Sub
Private Sub CommandButton16_Click()

    Call NextPage
   
End Sub

'page5
Private Sub CommandButton17_Click()
    Call PreviousPage
End Sub




Sub NextPage()

    Dim iNextPage As Long
    With Me.MultiPage1
        iNextPage = .Value + 1
        If iNextPage < .Pages.Count Then
            .Pages(iNextPage).Visible = True
            .Pages(iNextPage).Enabled = True
            .Value = iNextPage
            .Pages(iNextPage - 1).Enabled = False
            
            
            Select Case iNextPage
            
            Case 1
            
                 'functionModule.updateUserFormValue (3)
            Case 3
                functionModule.updateUserFormValue 4
                
            Case 4
                    Call formVisible
                    Call subModule.saveData("MAIN", UserForm1.TextBox1.Value, "B24")
                    Call subModule.saveData("MAIN", UserForm1.TextBox2.Value, "B25")
                    Call subModule.saveData("MAIN", UserForm1.TextBox3.Value, "B26")
                    Call subModule.saveData("MAIN", UserForm1.TextBox4.Value, "B27")
                    
                    ListBox1.MultiSelect = 2
                    ListBox3.MultiSelect = 2
                    ListBox4.MultiSelect = 2
                    ListBox5.MultiSelect = 2
                    
                    
            End Select
            
            
            
            
        If step < iNextPage Then
            step = iNextPage
        End If
            
        End If
    End With

End Sub

Sub PreviousPage()

    Dim iPrePage As Long
    With Me.MultiPage1
        iPrePage = .Value - 1
        If iPrePage >= 0 Then
            .Pages(iPrePage).Enabled = True
            .Value = iPrePage
            .Pages(iPrePage + 1).Enabled = False
            
        End If
        
        
            Select Case iPrePage
            
            Case 3
                 CommandButton2.SetFocus
            Case 4
                UserForm1.CommandButton40.Visible = False
            Case Else
                
            
            End Select
        
        
        
    End With
    
End Sub







''''''''''''''''''''''''''
''''''Example Setting ''''
''''''''''''''''''''''''''


Private Sub TextBox1_Enter()
        With TextBox1
        If .Text = "Please Enter lv3 PartNumber" Then
            .ForeColor = &H80000008 '<~~ Black Color
            .Text = ""
        End If
    End With
End Sub
Private Sub TextBox1_AfterUpdate()
    With TextBox1
        If .Text = "" Then
            .ForeColor = &H808080
            .Text = "Please Enter lv3 PartNumber"
        End If
    End With
End Sub

Private Sub TextBox2_Enter()
        With TextBox2
        If .Text = "Please Enter lv4 PartNumber" Then
            .ForeColor = &H80000008 '<~~ Black Color
            .Text = ""
        End If
    End With
End Sub
Private Sub TextBox2_AfterUpdate()
    With TextBox2
        If .Text = "" Then
            .ForeColor = &H808080
            .Text = "Please Enter lv4 PartNumber"
        End If
    End With
End Sub

Private Sub TextBox3_Enter()
        With TextBox3
        If .Text = "Please Enter lv5 PartNumber" Then
            .ForeColor = &H80000008 '<~~ Black Color
            .Text = ""
        End If
    End With
End Sub
Private Sub TextBox3_AfterUpdate()
    With TextBox3
        If .Text = "" Then
            .ForeColor = &H808080
            .Text = "Please Enter lv5 PartNumber"
        End If
    End With
End Sub
Private Sub TextBox4_Enter()
        With TextBox4
        If .Text = "Please Enter PCB PartNumber" Then
            .ForeColor = &H80000008 '<~~ Black Color
            .Text = ""
        End If
    End With
End Sub
Private Sub TextBox4_AfterUpdate()
    With TextBox4
        If .Text = "" Then
            .ForeColor = &H808080
            .Text = "Please Enter PCB PartNumber"
        End If
    End With
End Sub


Private Sub formVisible()
    'check enable
    
        bomlv = Workbooks(get_FileName_MakeBOM).Worksheets("Main").Range("J22").Value
    
        
        
        'Lv5
        
        ListBox3.Visible = True
        Frame3.Visible = True
        CommandButton20.Visible = True
        CommandButton21.Visible = True
        
        
        
        Select Case bomlv
        
        Case 3
           
            ListBox4.Visible = 0
            Frame4.Visible = 0
            CommandButton22.Visible = 0
            CommandButton23.Visible = 0
            ListBox5.Visible = 0
            Frame5.Visible = 0
            CommandButton24.Visible = 0
            CommandButton25.Visible = 0
        Case 4
            ListBox4.Visible = True
            Frame4.Visible = True
            CommandButton22.Visible = True
            CommandButton23.Visible = True
            ListBox5.Visible = 0
            Frame5.Visible = 0
            CommandButton24.Visible = 0
            CommandButton25.Visible = 0
        
        Case 5
             ListBox4.Visible = True
             Frame4.Visible = True
             CommandButton22.Visible = True
             CommandButton23.Visible = True
             ListBox5.Visible = True
             Frame5.Visible = True
             CommandButton24.Visible = True
             CommandButton25.Visible = True
        End Select
        
        
        
      
End Sub



''''''''''''''''''''''''''
'''''' TOP Side Item  ''''
''''''''''''''''''''''''''

Private Sub CommandButton6_Click()
    tempfilter = "Component file (*.xls;*.xlsx;*.csv), *.xls;*.xlsx;*.csv"
    TextBox5.Text = functionModule.getPath_Gen("MAIN", "B29", "B30", 22, tempfilter)
    'TextBox7.Text = functionModule.getPath("MAIN", "B31", "B32", 2)
End Sub

















''''''''''''''''''''''''''
'''''' Form Location  ''''
''''''''''''''''''''''''''


Private Sub UserForm_Activate()
     
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 180
    Me.Left = Application.Left + Application.Width - Me.Width - 50
    
End Sub


''''''''''''''''''''''''''
'''''' Save temp Data ''''
''''''''''''''''''''''''''

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)



    If CloseMode = 0 Then
    
        Answer = MsgBox("Do you want to retain the Data???", vbQuestion + vbYesNo, "")
 
            With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            End With
 
 
        If Answer = vbNo Then
            'Code for No button Press
            
            Call subModule.saveData("MAIN", "0", "B22")
            
            'Page1
            Call subModule.saveData("MAIN", "Please Enter lv3 PartNumber", "B24")
            Call subModule.saveData("MAIN", "Please Enter lv4 PartNumber", "B25")
            Call subModule.saveData("MAIN", "Please Enter lv5 PartNumber", "B26")
            Call subModule.saveData("MAIN", "Please Enter PCB PartNumber", "B27")
            
            'Page2
            Call subModule.saveData("MAIN", "", "B29")
            Call subModule.saveData("MAIN", "", "B30")
        
           
        Else
            'Code for Yes button Press
            
             Call subModule.saveData("MAIN", step, "B22")
            
            'Page1
            Call subModule.saveData("MAIN", UserForm1.TextBox1.Value, "B24")
            
            If UserForm1.TextBox2.Value = "" Then
                Call subModule.saveData("MAIN", "Please Enter lv4 PartNumber", "B25")
            Else
                Call subModule.saveData("MAIN", UserForm1.TextBox2.Value, "B25")
            End If
            
            If UserForm1.TextBox3.Value = "" Then
                Call subModule.saveData("MAIN", "Please Enter lv5 PartNumber", "B26")
            Else
                Call subModule.saveData("MAIN", UserForm1.TextBox3.Value, "B26")
            End If
            
            
            
            Call subModule.saveData("MAIN", UserForm1.TextBox4.Value, "B27")
            
            'Page2
            Call subModule.saveData("MAIN", UserForm1.TextBox5.Value, "B29")
            'Call subModule.saveData("MAIN", UserForm1.TextBox6.Value, "B30")
            
            
        End If
       
            With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
            End With
        
        
    End If
End Sub

Private Sub printInListBOX(ByRef data_PN As Variant, ByRef data_Loc As Variant, ByRef data_QTY As Variant)

    Dim tempAry

    indexRow = 0
    
    UserForm1.ListBox1.Clear
     UserForm1.ListBox5.Clear
     UserForm1.ListBox3.Clear
     UserForm1.ListBox4.Clear
    
    
    For i = 0 To data_PN.Count - 1
    
        
        If data_Loc(i) <> "" Then
                tempAry = Split(data_Loc(i), ",")
        Else
                ReDim tempAry(0 To data_QTY(i) - 1)
                
        End If
        
        
        
        For j = 0 To UBound(tempAry)
            ListBox1.AddItem
            ListBox1.List(indexRow, 0) = data_PN(i)
            ListBox1.List(indexRow, 1) = tempAry(j)
            indexRow = indexRow + 1
        Next
    
    
    
    Next


End Sub
