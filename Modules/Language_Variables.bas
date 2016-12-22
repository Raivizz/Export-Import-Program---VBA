Attribute VB_Name = "Language_Variables"
Global UserName As String
Global UserSurname As String
Global CompanyName As String

Global UserType As String

Global LanguageSelection As String
'Selection Dialogue
Global SD_ExportSuccessful As String
Global SD_NoData As String
Global SD_QuitPrompt As String
Global SD_QuitPromptTitle As String
Global SD_FDTitle As String
Global SD_NameNum As String
Global SD_SurnNum As String
Global SD_CompNum As String
'Data Import Form
Global DIF_NoFile As String
Global DIF_TMLError As String
Global DIF_Invalid As String
Global DIF_Return As String
Global DIF_ReturnPrompt As String
Global DIF_Search As String
Global DIF_SearchInput As String
Global DIF_ImportOK As String
Global DIF_FDTitle As String
Global DIF_ExportOK As String

Private Sub Language_English()
'Selection Form Strings - User Form
    SelectionDialogue_Intro.GreetingsLabel.Caption = "Welcome to the Data Import/Export Wizard."
    SelectionDialogue_Intro.RequestLabel.Caption = "Please select an action"
    SelectionDialogue_Intro.Caption = "Data Import/Export Wizard"
    SelectionDialogue_Intro.ImportDataButton.Caption = "Import new data to the current database"
    SelectionDialogue_Intro.ExportDataButton.Caption = "Export existing data to a text file"
    SelectionDialogue_Intro.ExportDataButtonSpecialExcel.Caption = "Export existing data to an Excel sheet"
    SelectionDialogue_Intro.CancelButton.Caption = "Cancel"
    SelectionDialogue_Intro.UserTypeLabel.Caption = "User type"
    SelectionDialogue_Intro.UserType_Standard.Caption = "Standard User"
    SelectionDialogue_Intro.UserType_Admin.Caption = "Administrator"
    SelectionDialogue_Intro.LanguageSelectLabel.Caption = "Language"
    SelectionDialogue_Intro.UserName.Caption = "User name"
    SelectionDialogue_Intro.UserSurname.Caption = "User surname"
    SelectionDialogue_Intro.CompanyName.Caption = "Company name"
'Selection Form Strings - Variables
    SD_ExportSuccessful = "Data has been succesfully exported"
    SD_NoData = "Database contains no data"
    SD_QuitPrompt = "Quit the application?"
    SD_QuitPromptTitle = "Cancel Prompt"
    SD_FDTitle = "Export data"
    SD_NameNum = "Name cannot be numeric."
    SD_SurnNum = "Surname cannot be numeric."
    SD_CompNum = "Company name cannot be numeric."

'Data Import Form - User Form
    DataImportForm.Caption = "Data Import"
    DataImportForm.OpenFile.ControlTipText = "Opens a file open dialog. Only .txt files are supported."
    DataImportForm.SaveFile.ControlTipText = "Saves the changes made to the file."
    DataImportForm.FilenameLabel.ControlTipText = "Currently opened file's directory and filename."
    DataImportForm.FilenameLabel.Caption = "No file opened"
    DataImportForm.ImportTableButton.ControlTipText = "Imports the table into the currently open database."
    DataImportForm.RemoveSearching.ControlTipText = "Resets searching and shows all generated labels."
    DataImportForm.SearchButton.ControlTipText = "Searches through the imported data."
    DataImportForm.DeleteLabels.ControlTipText = "Deletes the table."
    DataImportForm.ReturnToSelection.ControlTipText = "Returns to the Import/Export Wizard."
'Data Import Form - Variables
    DIF_NoFile = "No file opened"
    DIF_TMLError = "Too many labels in a single row. Please check the file formatting and try again."
    DIF_Invalid = "Invalid text file. Please check the file formatting and try again."
    DIF_Return = "Return to the selection dialogue?"
    DIF_ReturnPrompt = "Return Prompt"
    DIF_Search = "Enter a word to search. Only cells that contain the word will be shown."
    DIF_SearchInput = "Search Input"
    DIF_ImportOK = "Table has been succesfully imported into the database."
    DIF_FDTitle = "Export data"
    DIF_ExportOK = "Data has been succesfully exported"
    

End Sub
Private Sub Language_Gizoogled()
'Selection Fiznorm Sizzizzle - Usa Fizzorm
    SelectionDialogue_Intro.GreetingsLabel.Caption = "Welcome ta tha Dizzle Impizzle/Expizzle Wizzle."
    SelectionDialogue_Intro.RequestLabel.Caption = "Pleaze sizzle an actizzle."
    SelectionDialogue_Intro.Caption = "Data Impizzle/Expizzle Wizard"
    SelectionDialogue_Intro.ImportDataButton.Caption = "Impizzle new data ta tha current database"
    SelectionDialogue_Intro.ExportDataButton.Caption = "Export blingin' data ta a text file"
    SelectionDialogue_Intro.ExportDataButtonSpecialExcel.Caption = "Export blingin' data ta an Excizzle shizzeet"
    SelectionDialogue_Intro.CancelButton.Caption = "Cizzle"
    SelectionDialogue_Intro.UserTypeLabel.Caption = "Playa typizze"
    SelectionDialogue_Intro.UserType_Standard.Caption = "Standard Usa"
    SelectionDialogue_Intro.UserType_Admin.Caption = "Administrizzle"
    SelectionDialogue_Intro.LanguageSelectLabel.Caption = "Langizzle"
    SelectionDialogue_Intro.UserName.Caption = "Usa nizzame"
    SelectionDialogue_Intro.UserSurname.Caption = "Usa sizzle"
    SelectionDialogue_Intro.CompanyName.Caption = "Compizzle nizzle"
'Selection Fiznorm Str'n - Variables
    SD_ExportSuccessful = "Data has bizzay succesfully exported"
    SD_NoData = "Databaze contains no data"
    SD_QuitPrompt = "Qizzay tha application yeah yeah baby?"
    SD_QuitPromptTitle = "Cancel Prizzle"
    SD_FDTitle = "Export diznata"
    SD_NameNum = "Nizzle cannot be numerizzle."
    SD_SurnNum = "Surname ciznizzle be numeric. It dont stop till the wheels fall off."
    SD_CompNum = "Company name cannizzle be numerizzle gangsta style."

'Data Import Fizzorm - User Form
    DataImportForm.Caption = "Dizzy Impizzle"
    DataImportForm.OpenFile.ControlTipText = "Opens a file open dialog sho nuff. Only .txt files be supported."
    DataImportForm.SaveFile.ControlTipText = "Saves tha changes mizzy ta the F-to-tha-izzile. Drop it like its hot."
    DataImportForm.FilenameLabel.ControlTipText = "Currizzle openizzle file dirizzle n filename."
    DataImportForm.FilenameLabel.Caption = "No fiznizzle opizzle"
    DataImportForm.ImportTableButton.ControlTipText = "Imports tha table into tha currently open database."
    DataImportForm.RemoveSearching.ControlTipText = "Resets chillin' n shows all generatizzle lizzles."
    DataImportForm.SearchButton.ControlTipText = "Searches tha importizzle data."
    DataImportForm.DeleteLabels.ControlTipText = "Deletizzles tha tizzle. Death row 187 4 life."
    DataImportForm.ReturnToSelection.ControlTipText = "Returns ta tha Import/Export Wizizzle."
'D-A-to-tha-izzata Import Fiznorm - Variables
    DIF_NoFile = "No fizzy opened"
    DIF_TMLError = "Tizzy mizzle labels 'n a sizzy row. Pleaze check tha file ridin' n trizzle again."
    DIF_Invalid = "Invalid tizzay fiznile. Pliznease chizneck tha file formatt'n n try agizzle."
    DIF_Return = "Return ta tha selection dialogue?"
    DIF_ReturnPrompt = "Return Prompt"
    DIF_Search = "Enta a wizzord ta poser. Onlizzle cells thizzat contain tha word wizzill be shown."
    DIF_SearchInput = "Search Input"
    DIF_ImportOK = "T-A-to-tha-izzable has bizzle succesfully imported into tha database."
    DIF_FDTitle = "Export D-to-tha-izzata"
    DIF_ExportOK = "Data has bizzy succesfizzle expizzle"

End Sub
