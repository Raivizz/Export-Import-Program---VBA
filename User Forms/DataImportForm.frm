VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataImportForm 
   Caption         =   "Data Import"
   ClientHeight    =   1110
   ClientLeft      =   9540
   ClientTop       =   435
   ClientWidth     =   12960
   OleObjectBlob   =   "DataImportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataImportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GeneratedLineCount As Integer
Dim i As Integer
Dim LabelCount As Integer
Private Sub UserForm_Initialize()
If LanguageSelection = "English" Then
    Application.Run "Language_Variables.Language_English"
End If
If LanguageSelection = "Gizoogled" Then
    Application.Run "Language_Variables.Language_Gizoogled"
End If
FilenameLabel.Caption = DIF_NoFile

End Sub
Private Sub UserForm_Terminate()
Unload Me
End Sub
Private Sub OpenFile_Click()

OpenFileProcess:

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogOpen)

'Makes the file dialog only allow selecting text files, forbids multi-select.
With fd
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "Text Files", "*.txt"
End With

If fd.Show = -1 Then
    Database_Filename = fd.SelectedItems(1)
    FilenameLabel.Caption = Database_Filename
    Open Database_Filename For Input As #1
End If

If Database_Filename = "" Then
GoTo FileOpenCancel:
End If

Top_DynamicVar = 1
i = 1

'This section used to remove each existing imported label once a file is opened.
'Instead, the section was moved to another sub.
'Instead of copy-pasting the same code, the sub containing the code is called to handle it.
DeleteLabels_Click

Do Until EOF(1)
    Line Input #1, InputLine
    'Left offset and overflow protection are reset at the beginning of every line input.
    Left_DynamicVar = 0
    Label_OverFlowProtection = 0

'This section is used to skip empty lines in the text file.
If Len(InputLine) = 0 Then
Debug.Print "Empty line found, skipping..."
GeneratedLineCount = GeneratedLineCount - 1
GoTo SkipEmptyLineAndCommentary
End If

'This section is used to skip commentary lines in the text file.
'Commentary is marked the same way as in VBA.
'Used in the import example text file as guidance.
If Left(InputLine, 1) = "'" Then
Debug.Print "Commentary line found, skipping..."
GeneratedLineCount = GeneratedLineCount - 1
GoTo SkipEmptyLineAndCommentary
End If

DatabaseFile_Split = Split(InputLine, " | ")

'This is where all the magic happens.
'The string gets split up into words, which then get inserted into each label in a single line.
'Each split word gets assigned it's own label. Label count is dynamic.
   For WordCounter = LBound(DatabaseFile_Split) To UBound(DatabaseFile_Split)
      CurrentLabel_Uppercase = DatabaseFile_Split(WordCounter)

Set AddLabel = DataImportForm.Controls.Add("Forms.TextBox.1", "ImportLabel_" & i, True)

    With AddLabel
        .Value = UCase(CurrentLabel_Uppercase)
        .Left = 10 + Left_DynamicVar
        .Width = 90
        .Top = 50 + Top_DynamicVar
        .TextAlign = fmTextAlignCenter
        .SpecialEffect = fmSpecialEffectSunken
        .Height = 15
        
    End With

    Label_OverFlowProtection = Label_OverFlowProtection + 1
    Left_DynamicVar = Left_DynamicVar + 90
    If Label_OverFlowProtection > 7 Then
        MsgBox DIF_TMLError
        GoTo PrematureCancel
    Else
    i = i + 1
    End If
Next
Top_DynamicVar = Top_DynamicVar + 15
'The inconsistent placement of both Left and Top variables is done on purpose.
'The Left variable changes after a single label is made to put them into a single, nice-looking line.
'The Top variable changes after all 7 labels in a line have been generated.

SkipEmptyLineAndCommentary:
GeneratedLineCount = GeneratedLineCount + 1
Loop

'Outputs an error if the text file does not adhere to the required formatting
'e.g. column count is not 7, separators are missing, etc.
If i = 0 Or Label_OverFlowProtection < 7 Then
DeleteLabels_Click
MsgBox DIF_Invalid
Close #1
GoTo OpenFileProcess
End If

'Creates a button to acknowledge the data and import it into the active Excel sheet.
'Commented out due to extreme difficulty of implementing a working clickable dynamic button.

    'Set AddButton = DataImportForm.Controls.Add("Forms.CommandButton.1", "ImportTableButton", True)
    'With AddButton
        '.Caption = "Accept and import data into Excel"
        '.Width = 140
        '.Left = DataImportForm.Width / 2 - AddButton.Width / 2
        '.Top = 55 + Top_DynamicVar
    'End With
    
With Me
'This dynamically scales the scrollbar height to encompass all generated labels.
    .ScrollHeight = 55 + Top_DynamicVar
    .Height = 150
End With
If GeneratedLineCount <= 10 Then Me.Height = Me.Height + (GeneratedLineCount * 7)
If GeneratedLineCount > 10 Then Me.Height = Me.Height + (GeneratedLineCount * 5)
DataImportForm.ScrollBars = fmScrollBarsVertical
If Me.Height >= 600 Then Me.Height = 600
'Makes the table import button visible if the text file is valid.
If GeneratedLineCount >= 1 Then
With ImportTableButton
    .Enabled = True
    .Visible = True
    .Left = FilenameLabel.Width * 1.4
End With
With FilterButton
    .Enabled = True
    .Visible = True
End With
If UserType = "Admin" Then
    With SaveFile
        .Enabled = True
        .Visible = True
    End With
    End If
End If

i = i - 1
LabelCount = i


PrematureCancel:
                Debug.Print ("Amount of generated labels: " & LabelCount)
                Debug.Print ("Amount of generated lines: " & GeneratedLineCount)
                If UserType = "Standard" Then
                    Debug.Print ("User name is: " & SelectionDialogue_Intro.UserNameInput.Value)
                    Debug.Print ("User surname is: " & SelectionDialogue_Intro.UserSurnameInput.Value)
                    Debug.Print ("Company name is: " & SelectionDialogue_Intro.CompanyNameInput.Value)
                End If
FileOpenCancel:
                Close #1
                End Sub
Private Sub RemoveFiltering_Click()
For SetVisible = 1 To LabelCount
Controls("ImportLabel_" & SetVisible).Visible = 1
Next
End Sub
Private Sub ReturnToSelection_Click()
DataImportForm.Hide
If MsgBox(DIF_Return, vbYesNo, DIF_ReturnPrompt) = vbYes Then
SelectionDialogue_Intro.Show
Unload DataImportForm
Else
DataImportForm.Show
End If
End Sub
Private Sub DeleteLabels_Click()
For Each Control In Me.Controls
    If InStr(Control.Name, "ImportLabel_") = 1 Then
        Me.Controls.Remove Control.Name
    End If
Next Control
With ImportTableButton
    .Enabled = False
    .Visible = False
End With
With FilterButton
    .Enabled = False
    .Visible = False
End With
With RemoveFiltering
    .Enabled = False
    .Visible = False
End With
With SaveFile
    .Enabled = False
    .Visible = False
End With
DataImportForm.Height = 85
DataImportForm.ScrollBars = fmScrollBarsNone
End Sub
Private Sub FilterButton_Click()

Filter_Input = InputBox(DIF_Filter, DIF_FilterInput)
While i > 0
If InStr(Controls("ImportLabel_" & i), Filter_Input) = 0 Then
Controls("ImportLabel_" & i).Visible = 0
i = i - 1
Else
Controls("ImportLabel_" & i).Visible = 1
i = i - 1
End If
Wend
i = LabelCount
RemoveFiltering.Visible = 1
RemoveFiltering.Enabled = 1
End Sub
Private Sub ImportTableButton_Click()
ImportTableButton.SpecialEffect = fmSpecialEffectSunken
RowRowFightThePowa = 0
CurrentRow_Offset = 0
CurrentRow = 2
LabelCounter = 1

Range("A1", "G1").Clear
Range("A1").Value = Manufacturer_Label.Caption
Range("B1").Value = Model_Label.Caption
Range("C1").Value = Motherboard_Label.Caption
Range("D1").Value = CPU_Label.Caption
Range("E1").Value = GPU_Label.Caption
Range("F1").Value = RAM_Label.Caption
Range("G1").Value = OSHDD_Label.Caption
Do Until GeneratedLineCount = 0
    Do Until RowRowFightThePowa = 7
    Range("A" & CurrentRow).Activate
    ActiveCell.Offset(CurrentRow_Offset, RowRowFightThePowa) = Controls("ImportLabel_" & LabelCounter).Value
    RowRowFightThePowa = RowRowFightThePowa + 1
    LabelCounter = LabelCounter + 1
    Loop
GeneratedLineCount = GeneratedLineCount - 1
CurrentRow = CurrentRow + 1
CurrentRow_Offset = 0
RowRowFightThePowa = 0
Loop

MsgBox DIF_ImportOK
ImportTableButton.SpecialEffect = fmSpecialEffectRaised

End Sub
Private Sub SaveFile_Click()
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogSaveAs)

With fd
    .Title = DIF_FDTitle
End With


Dim filename As String
If fd.Show = -1 Then
  filename = fd.SelectedItems(1)
  Open filename For Output As #1
  End If
  
If filename = "" Then
GoTo EndSaving
End If

LabelCounter = 1
Do Until GeneratedLineCount = 0
    Print #1, Controls("ImportLabel_" & LabelCounter).Value
    LabelCounter = LabelCounter + 1
GeneratedLineCount = GeneratedLineCount - 1
Loop

End If

MsgBox DIF_ExportOK

EndSaving:
Close #1
End Sub
