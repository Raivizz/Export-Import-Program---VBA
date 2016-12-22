VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectionDialogue_Intro 
   Caption         =   "Data Import/Export Wizard"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "SelectionDialogue_Intro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectionDialogue_Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExportDataButtonSpecialExcel_Click()
'Disables alerts and screen updating temporarily
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Checks if the current opened worksheet is empty
If WorksheetFunction.CountA(Cells) = 0 Then
MsgBox SD_NoData
GoTo ExportFileEnd
End If

If UserType = "Standard" Then
If IsNumeric(UserNameInput.Value) = True Then
MsgBox SD_NameNum
GoTo ExportFileEnd
End If
If IsNumeric(UserSurnameInput.Value) = True Then
MsgBox SD_SurnNum
GoTo ExportFileEnd
End If
If IsNumeric(CompanyNameInput.Value) = True Then
MsgBox SD_CompNum
GoTo ExportFileEnd
End If
If UserNameInput.Value = "" Then
UserNameInput.Value = "Unnamed"
End If
If UserSurnameInput.Value = "" Then
UserSurnameInput.Value = "McNoSurnameFace"
End If
If CompanyNameInput.Value = "" Then
CompanyNameInput.Value = "Undefined Industries"
End If

End If

    Sheets.Add After:=ActiveSheet
    Sheets(Sheets.Count).Name = "ExportSheet"
    Sheets("ExportSheet").Select
    
    Columns("A:G").ColumnWidth = 20
    
    Range("A1").Value = "User name"
    Range("A2").Value = UserNameInput.Value
        If UserType = "Admin" Then
        Range("A2").Value = UserType
        End If
    Range("B1").Value = "User surname"
    Range("B2").Value = UserSurnameInput.Value
        If UserType = "Admin" Then
        Range("B2").Value = UserType
        End If
    Range("C1").Value = "Company name"
    Range("C2").Value = CompanyNameInput.Value
        If UserType = "Admin" Then
        Range("C2").Value = UserType
        End If
    
    Range("A1:C1").Font.Bold = True
    Range("A4:G4").Font.Bold = True
    
    Range("A1:C2").Borders.LineStyle = xlContinuous
    Range("A1:C2").Borders.Weight = xlMedium
    
    Range("A1:C2").HorizontalAlignment = xlCenter
    Range("A1:C2").VerticalAlignment = xlCenter
    
    
    i = 1
    Do While Sheets("MainSheet").Cells(i, 1) <> ""
    RangeCounter = RangeCounter + 1
        For j = 1 To 8
        Sheets("ExportSheet").Cells(i + 3, j) = Sheets("MainSheet").Cells(i, j)
        Next
        i = i + 1
    Loop
    
    Range("A4:G" & RangeCounter + 3).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("A4:G4").Borders.Weight = xlMedium
    
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogSaveAs)

With fd
    .Title = SD_FDTitle
End With


Dim filename As String
If fd.Show = -1 Then
  filename = fd.SelectedItems(1)
End If
  
If filename = "" Then
GoTo ExportFileEnd
End If

Sheets("MainSheet").Visible = False
Sheets("ExportSheet").Copy
Application.ActiveWorkbook.SaveAs filename, FileFormat:=51
Application.ActiveWorkbook.Close False

MsgBox SD_ExportSuccessful

ExportFileEnd:
Sheets("MainSheet").Visible = True
Sheets("ExportSheet").Delete
'Enables application alerts and screen updating
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()
If UserType_Admin.Value = True Then
UserType_Admin.Value = True
UserType = "Admin"
SelectionDialogue_Intro.Width = 180
Else
UserType_Standard.Value = True
SelectionDialogue_Intro.Width = 410
UserType = "Standard"
End If

LanguageSelectBox.AddItem "English"
LanguageSelectBox.AddItem "Gangster"
LanguageSelectBox.Value = "English"

End Sub
Private Sub UserForm_Terminate()
Unload Me
End Sub
Private Sub ExportDataButton_Click()

'Checks if the current opened worksheet is empty
If WorksheetFunction.CountA(Cells) = 0 Then
MsgBox SD_NoData
GoTo ExportFileEnd
End If

If UserType = "Standard" Then
If IsNumeric(UserNameInput.Value) = True Then
MsgBox SD_NameNum
GoTo ExportFileEnd
End If
If IsNumeric(UserSurnameInput.Value) = True Then
MsgBox SD_SurnNum
GoTo ExportFileEnd
End If
If IsNumeric(CompanyNameInput.Value) = True Then
MsgBox SD_CompNum
GoTo ExportFileEnd
End If
If UserNameInput.Value = "" Then
UserNameInput.Value = "Unnamed"
End If
If UserSurnameInput.Value = "" Then
UserSurnameInput.Value = "McNoSurnameFace"
End If
If CompanyNameInput.Value = "" Then
CompanyNameInput.Value = "Undefined Industries"
End If

End If

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogSaveAs)

With fd
    .Title = SD_FDTitle
End With


Dim filename As String
If fd.Show = -1 Then
  filename = fd.SelectedItems(1)
  Open filename For Output As #1
  End If
  
If filename = "" Then
GoTo ExportFileEnd
End If

If UserType = "Standard" Then
Print #1, "Stock list generated by " & SelectionDialogue_Intro.UserNameInput.Value; " " & SelectionDialogue_Intro.UserSurnameInput.Value; " of company " & SelectionDialogue_Intro.CompanyNameInput.Value
Print #1, ""
Print #1, "Available builds in the shop are:"
End If

If UserType = "Admin" Then
Print #1, "Stock list generated by an administrator"
Print #1, ""
Print #1, "Available builds in the shop are:"
End If

    i = 1
    Do While Cells(i, 1) <> ""
        StringInput = ""
        For j = 1 To 8
            If Cells(i, j) = "" Then
                StringInput = StringInput
            Else
                StringInput = StringInput & Cells(i, j) & " | "
            End If
            
        Next
        Print #1, StringInput
        i = i + 1
    Loop

MsgBox SD_ExportSuccessful
ExportFileEnd:
                Close #1
                End Sub
Private Sub ImportDataButton_Click()
SelectionDialogue_Intro.Hide
DataImportForm.Show
DoNothing:
End Sub
Private Sub CancelButton_Click()
SelectionDialogue_Intro.Hide
If MsgBox(SD_QuitPrompt, vbYesNo, SD_QuitPromptTitle) = vbYes Then
Unload SelectionDialogue_Intro
Else
SelectionDialogue_Intro.Show
End If
End Sub
Private Sub UserType_Admin_Click()
UserType = "Admin"
SelectionDialogue_Intro.Width = 180
End Sub
Private Sub UserType_Standard_Click()
UserType = "Standard"
SelectionDialogue_Intro.Width = 410
End Sub
Private Sub LanguageSelectBox_Change()
If LanguageSelectBox.Value = "English" Then
Application.Run "Language_Variables.Language_English"
LanguageSelection = "English"
End If
If LanguageSelectBox.Value = "Gangster" Then
Application.Run "Language_Variables.Language_Gizoogled"
LanguageSelection = "Gizoogled"
End If
End Sub
