Attribute VB_Name = "modMain"
' This is a macro-enabled global template containing the XML and VBA for the DISARM Word Plug-In.

' Copyright (C) 2023 DISARM Foundation

' This program is free software: you can redistribute it and/or modify it under the terms of the GNU 'General Public License
' as published by the Free Software Foundation, either version 3 of the License, or 'any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License at https://www.gnu.org/licenses/ for more details.

' To report bugs or suggestions for improvement please send email to info@disarm.foundation.
'___________________________________________________________________________________________________________________________

Option Explicit
Dim arrTechniques() As String
Dim arrCountermeasures() As String
Dim counter As Integer
Global oApp As Excel.Application
Global oWB_FrameworkMaster As Excel.Workbook
Const cPathXlstart As String = "\AppData\Roaming\Microsoft\Excel\XLSTART\"
Const cPathWordStartup As String = "\AppData\Roaming\Microsoft\Word\STARTUP\"
Const cSourceFrameworkMaster As String = "DISARM_FRAMEWORKS_MASTER.xlsx"
Const cSourceTaggingWorkbook As String = "DISARM_TAGGING_WORKBOOK.xlsx"
Const cSourceTaggingWorkbookxls As String = "DISARM_TAGGING_WORKBOOK.xls"
Global oWB_TaggingWorkbook As Excel.Workbook
Global oWS_techniques As Excel.Worksheet
Global oWS_metatechniques As Excel.Worksheet
Global oWS_countermeasures As Excel.Worksheet
Global oWS_SummaryRedUnformatted As Excel.Worksheet
Global oWS_SummaryBlueUnformatted As Excel.Worksheet
Global oWS_SummaryRedGraphic As Excel.Worksheet
Global oWS_Profile As Excel.Worksheet
Global oWS_UserProfile As Excel.Worksheet
Global oWS_DISARM_Red_with_IDs As Excel.Worksheet
Global oWS_DISARM_Red_no_IDs As Excel.Worksheet
Global oWS_RowCount_SRU As Double
Global oWS_RowCount_SBU As Double
Dim bStartApp As Boolean
Dim iWindowState As Integer
Dim tblSummaryRed As Table
Dim tblSummaryBlue As Table
Global lngSetRedTagColor As Double
Global lngSetRedGraphicColor As Double
Global bSetColorInlineTag As Boolean
Global lngSetBlueTagColor As Double
Global intTableRowNumber As Integer
Global intBlueTableRowNumber As Integer

'Following variables for Search Techniques
Global gstrSearchTerm As String
Global gstrPhaseName As String
Global gstrTacticName As String
Global bLookInDescriptions As Boolean
Global bTechniquesFound As Boolean

'Following variables for Search Countermeasures
Global gstrSearchTermCountermeasures As String
Global gstrMetatechniqueName As String
Global bLookInDescriptionsCountermeasures As Boolean
Global bCountermeasuresFound As Boolean

Private oConn As Object '(Late binding.  Use "As New ADODB.Connection" for early binding") _
                          'and add reference to Microsoft ActiveX Data Objects library.
Private oRS As Object '(or Use "As New ADODB.Recordset")
Private lngNumRecs As Long

Dim X As New EventClassModule

Sub Register_Event_Handler()

'
' This procedure connects the declared object in the class module EventClassModule (App in this example)
' with the Application object. After the procedure is run, the App object in the class module points to
' the Microsoft Word Application object, and the event procedures in the class module will run
' when the events occur. We use the DocumentBeforeClose to clean up Excel before the last DISARM-enabled
' document is closed.
'
' See https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "Register_Event_Handler"
  
Set X.App = Application

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Public Sub WriteTag(ByVal strText As String, ByVal strTagType As String)

'
' Write the tag out to the Word document
'

Dim oRng As Range

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "WriteTag"
  
Selection.InsertAfter (strText)
Set oRng = Selection.Range

If strTagType = "Red" Then
    lngSetRedTagColor = modMain.GetProfile("Red Color")
    If lngSetRedTagColor <> 0 Then
        oRng.Font.Shading.BackgroundPatternColor = lngSetRedTagColor
    End If
ElseIf strTagType = "Blue" Then
    lngSetBlueTagColor = modMain.GetProfile("Blue Color")
    If lngSetBlueTagColor <> 0 Then
        oRng.Font.Shading.BackgroundPatternColor = lngSetBlueTagColor
    End If
End If

Selection.Collapse (wdCollapseEnd)

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CreateTaggingSheets()

'
' We create four worksheets for each DISARM document. SummaryRedUnformatted records all red tags (techniques).
' SummaryBlueUnformatted records all blue tags (countermeasures). SummaryRedGraphic is a graphic of the red
' framework with all tagged techniques highlighted. The profile is used to record the color preferences for these.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateTaggingSheets"
  
modMain.CreateSummaryRedUnformatted
modMain.CreateSummaryRedGraphic
modMain.CreateSummaryBlueUnformatted
modMain.CreateProfile

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub WriteProfile(strColumn As String, varValue As Variant)

'
' Stores color preferences for tags and graphic
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "WriteProfile"

'
' The profile worksheet is used to store the formats the user desires for the tagging. This routine writes a specific frmatting value
' to the profile.
'

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If profile does not yet exist then create it. If oWS_Profile has been set check that it points to the active document. Do this by checking the
' datetimestamp only since the user might change the name of the document at any time and the DISARM Name will always contain the name of the
' document when the user first started tagging. Instrrev returns the position of an occurrence of one string in another, from the end of the string.
' If not create a profile worksheet for the active document.
'

If oWS_Profile Is Nothing Then
    modMain.CreateProfile
ElseIf Mid(oWS_Profile.Name, InStrRev(oWS_Profile.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateProfile
End If

'
' Update value in profile according to parameters
'

If strColumn = "Red Color" Then
    oWS_Profile.Cells(2, 1).Value = varValue
ElseIf strColumn = "Red Graphic Color" Then
    oWS_Profile.Cells(2, 2).Value = varValue
ElseIf strColumn = "Blue Color" Then
    oWS_Profile.Cells(2, 3).Value = varValue
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Function GetProfile(strColumn As String) As Variant

'
' Look up color preference
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "GetProfile"

'
' Retrieve desired formatting value from profile
'

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If profile does not yet exist then create it: if oWS_Profile has been set then make sure that it points to the profile for the active document.
' If it does not, then create a profile for the active document. May need to apply this logic to other worksheet objects.....
'

If oWS_Profile Is Nothing Then
    modMain.CreateProfile
ElseIf Mid(oWS_Profile.Name, InStrRev(oWS_Profile.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateProfile
End If

If strColumn = "Red Color" Then
    GetProfile = oWS_Profile.Cells(2, 1).Value
ElseIf strColumn = "Red Graphic Color" Then
    GetProfile = oWS_Profile.Cells(2, 2).Value
ElseIf strColumn = "Blue Color" Then
    GetProfile = oWS_Profile.Cells(2, 3).Value
End If

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub InitializeExcelAndOpenWorkbooks()

'
' Start Excel if not already running and open the workbooks for the DISARM Framework and for DISARM Tagging. The DISARM Framework should be taken from
' https://github.com/DISARMFoundation/DISARMframeworks/blob/main/DISARM_MASTER_DATA/DISARM_FRAMEWORKS_MASTER.xlsx.
'
' The Excel workbooks are installed in the XLSTART directory i.e.
' CStr(Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Excel\XLSTART")
' which is easier than having the user specify an installation directory. Consistent with this the .dotm global Word template is
' installed in the STARTUP directory for Word i.e.
' CStr(Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Word\STARTUP")
'

Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InitializeExcelAndOpenWorkbooks"

strPath = Environ("USERPROFILE") & cPathXlstart

If oApp Is Nothing Then
    On Error Resume Next
    Set oApp = GetObject(, "Excel.Application")
    If Err Then
        If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
        bStartApp = True
        Set oApp = New Excel.Application
    Else
        If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
    End If
End If

With oApp
    .Visible = False
    .ScreenUpdating = False
    '.Calculation = xlCalculationManual
    iWindowState = .WindowState
    '.WindowState = xlMinimized
    Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    Set oWB_TaggingWorkbook = .Workbooks.Open(strPath & cSourceTaggingWorkbook)
    'Debug.Print "New Workbook object oWB_TaggingWorkbook (InitializeExcelAndOpenWorkbooks)"
    Set oWS_techniques = oWB_FrameworkMaster.Worksheets("techniques")
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CreateSummaryBlueUnformatted()

'
' Add a worksheet to the Tagging Workbook to record blue tags for use in summary table
'

Dim strPrefix As String
Dim strMsg As String
Dim ReturnMsgBox As Integer
Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateSummaryBlueUnformatted"

strPrefix = DISARM_Name()

'
' SBU = Summary Blue Unformatted - this is a temporary sheet used to record all the blue tags inserted by the user in an unformatted table
'

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateSummaryBlueUnformatted)"
End If
    
If WorksheetExists(strPrefix & "_SBU", oWB_TaggingWorkbook) Then
    Set oWS_SummaryBlueUnformatted = oWB_TaggingWorkbook.Worksheets(strPrefix & "_SBU")
Else
    'On Error Resume Next
    Set oWS_SummaryBlueUnformatted = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_SummaryBlueUnformatted.Name = strPrefix & "_SBU"
    oWS_SummaryBlueUnformatted.Cells(1, 1).Value = "MetatechniqueID"
    oWS_SummaryBlueUnformatted.Cells(1, 2).Value = "MetatechniqueName"
    oWS_SummaryBlueUnformatted.Cells(1, 3).Value = "CountermeasureID"
    oWS_SummaryBlueUnformatted.Cells(1, 4).Value = "CountermeasureName"
    oWS_SummaryBlueUnformatted.Cells(1, 5).Value = "Sentence"
    oWS_SummaryBlueUnformatted.Cells(1, 6).Value = "SentenceIndex"
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CreateSummaryRedUnformatted()

'
' Add a worksheet to the Tagging Workbook to record red tags for use in summary table
'

Dim strPrefix As String
Dim strMsg As String
Dim ReturnMsgBox As Integer
Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateSummaryRedUnformatted"

strPrefix = DISARM_Name()

'
' SRU = Summary Red Unformatted - this is a temporary sheet used to record all the red tags inserted by the user in an unformatted table
'

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateSummaryRedUnformatted)"
End If
    
If WorksheetExists(strPrefix & "_SRU", oWB_TaggingWorkbook) Then
    Set oWS_SummaryRedUnformatted = oWB_TaggingWorkbook.Worksheets(strPrefix & "_SRU")
Else
    'On Error Resume Next
    Set oWS_SummaryRedUnformatted = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_SummaryRedUnformatted.Name = strPrefix & "_SRU"
    oWS_SummaryRedUnformatted.Cells(1, 1).Value = "TacticID"
    oWS_SummaryRedUnformatted.Cells(1, 2).Value = "TacticName"
    oWS_SummaryRedUnformatted.Cells(1, 3).Value = "TechniqueID"
    oWS_SummaryRedUnformatted.Cells(1, 4).Value = "TechniqueName"
    oWS_SummaryRedUnformatted.Cells(1, 5).Value = "Sentence"
    oWS_SummaryRedUnformatted.Cells(1, 6).Value = "SentenceIndex"
    oWS_SummaryRedUnformatted.Cells(1, 7).Value = "Status"
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub CreateProfile()

'
' Add a worksheet to the Tagging Workbook to record desired format of red and blue tags
'

Dim strPrefix As String
Dim strMsg As String
Dim ReturnMsgBox As Integer
Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateProfile"

strPrefix = DISARM_Name()

'
' P = Profile - this is a temporary sheet used to record the formats for the red and blue tags
'

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateProfile)"
End If
    
'
' If profile does not yet exist then create it.
' Set default colors to yellow.
'

If WorksheetExists(strPrefix & "_P", oWB_TaggingWorkbook) Then
    Set oWS_Profile = oWB_TaggingWorkbook.Worksheets(strPrefix & "_P")
Else
    'On Error Resume Next
    Set oWS_Profile = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_Profile.Name = strPrefix & "_P"
    oWS_Profile.Cells(1, 1).Value = "Highlight Color Red Inline Tag"
    oWS_Profile.Cells(1, 2).Value = "Highlight Color Red Graphic Cell"
    oWS_Profile.Cells(1, 3).Value = "Highlight Color Blue Inline Tag"
    oWS_Profile.Cells(2, 1).Value = 65535 ' yellow
    oWS_Profile.Cells(2, 2).Value = 65535 ' yellow
    oWS_Profile.Cells(2, 3).Value = 65535 ' yellow
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CreateSummaryRedGraphic()
'
' Add a worksheet to the Tagging Workbook to highlight the red tags for use in summary graphic
'

Dim strPrefix As String
Dim ReturnMsgBox As Integer
Dim oThirdRow As Excel.Range
Dim oLastRow As Excel.Range
Dim oLastColumn As Excel.Range
Dim intRowCount As Integer
Dim intColumnCount As Integer
Dim strMsg As String
Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateSummaryRedGraphic"

strPrefix = DISARM_Name()

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateSummaryRedGraphic)"
End If

If WorksheetExists(strPrefix & "_SRG", oWB_TaggingWorkbook) Then
    Set oWS_SummaryRedGraphic = oWB_TaggingWorkbook.Sheets(strPrefix & "_SRG")
Else
    '
    ' May 2024. If the tagging workbook is hidden this Copy will fail with error 1004
    ' "Copy Method of Worksheet Class Failed"
    ' See https://stackoverflow.com/questions/9327613/excel-vba-copy-method-of-worksheet-fails/9329827#9329827
    ' We can make the workbook visible then do the copy but it doesn't look good
    ' So make sure the tagging workbook is saved in the XLSTART directory in a unhidden state
    '
    'oWB_TaggingWorkbook.Windows(1).Visible = True
    oWB_TaggingWorkbook.Sheets("DISARM Red with IDs").Copy After:=oWB_TaggingWorkbook.Sheets(1)
    'oWB_TaggingWorkbook.Windows(1).Visible = False
    oWB_TaggingWorkbook.Sheets("DISARM Red with IDs (2)").Name = strPrefix & "_SRG"
    Set oWS_SummaryRedGraphic = oWB_TaggingWorkbook.Sheets(strPrefix & "_SRG")
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Function DISARM_Name() As String

'
' Returns the DISARM Name for the document. If no DISARM Name exists then create one.
'
' Use Word document variable to uniquely identify the document so that if it is saved or renamed it always has the same identifier until this
' variable is reset which the user can do by invoking the procedure to "Clear Summaries" in which case the document receives a new unique identifier.
' Think of this unique identifier as a kind of cookie which persists over multiple Word sessions even if the user saves the document under different names
' at different times.
'
' https://stackoverflow.com/questions/58156956/word-vba-how-to-reuse-a-temporary-unsaved-word-document
' https://stackoverflow.com/questions/27923926/file-name-without-extension-name-vba
'

Dim strPrefix As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "DISARM_Name"

On Error Resume Next
strPrefix = Application.ActiveDocument.Variables("DISARM_Name")
If Err.Number <> 0 Then
    'On Error GoTo 0
    If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0
    Dim fso As New Scripting.FileSystemObject
    If Len(fso.GetBaseName(Application.ActiveDocument.Name)) > 12 Then
        strPrefix = Left(fso.GetBaseName(Application.ActiveDocument.Name), 12)
    Else
        strPrefix = fso.GetBaseName(Application.ActiveDocument.Name)
    End If
    strPrefix = strPrefix & Format(Date, "mmddyyyy") & Format(Time, "hhmmss")
    Application.ActiveDocument.Variables("DISARM_Name") = strPrefix
End If
DISARM_Name = strPrefix

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
Sub Reset_DISARM_Name()

'
' Reset the DISARM Name. If the user chooses to clear the summaries then first reset the DISARM Name and then recreate the summary worksheets. Perhaps
' do not need to explicitly recreate the summaries as this will be done automatically via the strPrefix logic.
'

Dim strPrefix As String
Dim fso As New Scripting.FileSystemObject

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "Reset_DISARM_Name"

If Len(fso.GetBaseName(Application.ActiveDocument.Name)) > 12 Then
    strPrefix = Left(fso.GetBaseName(Application.ActiveDocument.Name), 12)
Else
    strPrefix = fso.GetBaseName(Application.ActiveDocument.Name)
End If
strPrefix = strPrefix & Format(Date, "mmddyyyy") & Format(Time, "hhmmss")
Application.ActiveDocument.Variables("DISARM_Name") = strPrefix

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub CheckReadyToTag()

'
' Checks whether we are ready to tag i.e. all the worksheets in the tagging workbook have been created
' for this document and worksheet objects have been initialized.
'

Dim strPrefix As String
Dim strMsg As String
Dim ReturnMsgBox As Integer
Dim strPath As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CheckReadyToTag"
    
'
' Use Word document variable to uniquely identify the document so that if it is saved or renamed it always has the same identifier until this
' variable is reset which the user can do by invoking the procedure to "Clear Summaries" in which case the document receives a new unique identifier.
' Think of this unique identifier as a kind of cookie which persists over multiple Word sessions even if the user saves the document under different names
' at different times.
'
' https://stackoverflow.com/questions/58156956/word-vba-how-to-reuse-a-temporary-unsaved-word-document
' https://stackoverflow.com/questions/27923926/file-name-without-extension-name-vba
'

strPrefix = DISARM_Name()

'
' Check if Excel running. If not start it and open Workbooks.
'

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' SRU = Summary Red Unformatted - this is a temporary sheet used to record all the red tags inserted by the user in an unformatted table
'

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CheckReadyToTag)"
End If
    
'
' First check summary red tagging worksheet
'

If WorksheetExists(strPrefix & "_SRU", oWB_TaggingWorkbook) Then
    Set oWS_SummaryRedUnformatted = oWB_TaggingWorkbook.Worksheets(strPrefix & "_SRU")
Else
    'On Error Resume Next
    Set oWS_SummaryRedUnformatted = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_SummaryRedUnformatted.Name = strPrefix & "_SRU"
    oWS_SummaryRedUnformatted.Cells(1, 1).Value = "TacticID"
    oWS_SummaryRedUnformatted.Cells(1, 2).Value = "TacticName"
    oWS_SummaryRedUnformatted.Cells(1, 3).Value = "TechniqueID"
    oWS_SummaryRedUnformatted.Cells(1, 4).Value = "TechniqueName"
    oWS_SummaryRedUnformatted.Cells(1, 5).Value = "Sentence"
    oWS_SummaryRedUnformatted.Cells(1, 6).Value = "SentenceIndex"
End If

'
' Now check summary red graphic worksheet
'

If WorksheetExists(strPrefix & "_SRG", oWB_TaggingWorkbook) Then
    Set oWS_SummaryRedGraphic = oWB_TaggingWorkbook.Sheets(strPrefix & "_SRG")
Else
    Set oWS_SummaryRedGraphic = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_SummaryRedGraphic.Name = strPrefix & "_SRG"
    oWB_TaggingWorkbook.Sheets("DISARM Red with IDs").UsedRange.Copy
    oWS_SummaryRedGraphic.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    oWS_SummaryRedGraphic.Range("A1").PasteSpecial xlPasteFormats
End If

'
' Now check summary blue worksheet
'

If WorksheetExists(strPrefix & "_SBU", oWB_TaggingWorkbook) Then
    Set oWS_SummaryBlueUnformatted = oWB_TaggingWorkbook.Worksheets(strPrefix & "_SBU")
Else
    'On Error Resume Next
    Set oWS_SummaryBlueUnformatted = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_SummaryBlueUnformatted.Name = strPrefix & "_SBU"
    oWS_SummaryBlueUnformatted.Cells(1, 1).Value = "MetatechniqueID"
    oWS_SummaryBlueUnformatted.Cells(1, 2).Value = "MetatechniqueName"
    oWS_SummaryBlueUnformatted.Cells(1, 3).Value = "CountermeasureID"
    oWS_SummaryBlueUnformatted.Cells(1, 4).Value = "CountermeasureName"
    oWS_SummaryBlueUnformatted.Cells(1, 5).Value = "Sentence"
    oWS_SummaryBlueUnformatted.Cells(1, 6).Value = "SentenceIndex"
End If

'
' Lastly check profile. If profile does not yet exist then create it. Set the default value for "Apply to Red Inline Tag" to True so that by default
' red tags are highlighted
'

If WorksheetExists(strPrefix & "_P", oWB_TaggingWorkbook) Then
    Set oWS_Profile = oWB_TaggingWorkbook.Worksheets(strPrefix & "_P")
Else
    'On Error Resume Next
    Set oWS_Profile = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_Profile.Name = strPrefix & "_P"
    oWS_Profile.Cells(1, 1).Value = "Highlight Color Red Inline Tag"
    oWS_Profile.Cells(1, 2).Value = "Highlight Color Red Graphic Cell"
    oWS_Profile.Cells(1, 3).Value = "Highlight Color Blue Inline Tag"
    oWS_Profile.Cells(2, 1).Value = 65535
    oWS_Profile.Cells(2, 2).Value = 65535
    oWS_Profile.Cells(2, 3).Value = 65535
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Function FillListTags(oListOrComboBox As Object)

'
' Populate the listbox oListOrComboBox with all techniques that have been tagged and are "Active"
'

Dim lngRows As Long
Dim lngIndex As Long
Dim i As Integer
Dim dblCount As Double
Dim varValue As Variant
Dim arrTableData() As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "FillListTags"

With oListOrComboBox
    .Clear
    'Create the list entries.
    lngRows = oWS_SummaryRedUnformatted.UsedRange.Rows.Count
    varValue = "Active"
    dblCount = oApp.WorksheetFunction.CountIf(oWS_SummaryRedUnformatted.UsedRange.Columns(7).Cells, varValue)
    If dblCount > 0 Then ' If there are active techniques
        ReDim arrTableData(1 To dblCount, 1 To 4)
        i = 0
        For lngIndex = 2 To lngRows
          'Add a row to the array
          If oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 7) = "Active" Then
            i = i + 1
            arrTableData(i, 1) = oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 3)
            arrTableData(i, 2) = oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 4)
            arrTableData(i, 3) = oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 6)
            ' remove carriage return, line feed, and carriage return line feed
            arrTableData(i, 4) = Replace(Replace(Replace(oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 5), vbCrLf, ""), vbCr, ""), vbLf, "")
          End If
        Next lngIndex
    'Populate listbox with array
    .List = arrTableData
    End If
End With
    
FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
  
Public Function MarkTagAsDeleted(SentenceIndex As String, TechniqueID As String)

'
' Update the tag in the tagging worksheet to mark it as deleted
'

Dim lngRows As Long
Dim lngIndex As Long

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "MarkTagAsDeleted"

lngRows = oWS_SummaryRedUnformatted.UsedRange.Rows.Count
For lngIndex = 2 To lngRows
  If oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 6) = SentenceIndex Then
    If oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 3) = TechniqueID Then
        oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 7).Value = "Deleted"
    End If
  End If
Next lngIndex

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub CreateSummaryRedTable()

'
' create a table within the Word document to house the list of attacker TTPs identified during tagging.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateSummaryRedTable"

If Selection.IsEndOfRowMark = True Then Selection.MoveDown

Selection.InsertBreak Type:=wdSectionBreakNextPage

With Selection.PageSetup
    If .Orientation = wdOrientPortrait Then
        .Orientation = wdOrientLandscape
    End If
    .TopMargin = InchesToPoints(1#) 'amended from 0.5 for ALF/debunk demo
    .BottomMargin = InchesToPoints(0.5)
    .LeftMargin = InchesToPoints(1) 'amended from 0.5 for ALF/debunk demo
    .RightMargin = InchesToPoints(0.5)
End With

'
' Create table. Fit table width within the window
'

Set tblSummaryRed = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:= _
    3, DefaultTableBehavior:=wdWord8TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
tblSummaryRed.AllowAutoFit = False
tblSummaryRed.AutoFormat Format:=0

With tblSummaryRed

'    If .Style <> "Table Grid" Then
'        .Style = "Table Grid"
'    End If
'
'    This code was bombing out with error 5834 "A megadott nevu elem nem létezik"
'    ("Item with specified name does not exist") in the Hungarian version of Word
'    so use the numeric value for Style "Table Grid" so that it works for any language
'    See https://www.msofficeforums.com/word-vba/46753-install-language-macro-language.html
'
    .Style = -155
    
    .AllowAutoFit = False
    .Columns.PreferredWidthType = wdPreferredWidthPoints
    .Columns(1).PreferredWidth = InchesToPoints(2)
    .Columns(2).PreferredWidth = InchesToPoints(0.9)
    .Columns(3).PreferredWidth = InchesToPoints(7#) ' was 5.9 - amended for ALF/debunk demo - need to rewrite code to fit to page width
    
    With .Rows(1)
    
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(1).Range.Text = "Technique Title"
        .Cells(2).Range.Text = "ID"
        .Cells(3).Range.Text = "Use"
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorBlueGray
        
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 14
            .Bold = True
            .ColorIndex = wdWhite
        End With
        
    End With
    
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
'Sub CreateNavigatorFile()
'
''
'' Create a layer file in JSON format for the DISARM Navigator from the techniques tagged by the user for this document
''
'
'Dim JSON As Object
'Dim JsonVBA As String
'Dim Part1 As String
'Dim Part2 As String
'Dim Part3 As String
'Dim fso As New FileSystemObject
'Dim fso2 As New FileSystemObject
'Dim fso3 As New FileSystemObject
'Dim DocumentName As String
'Dim strFolderPath As String
'Dim tsout As TextStream
'Dim JSONDirectory As String
'Dim sFolder As String
'Dim oFD As FileDialog
'Dim intMsgReturn As Integer
'
'If gcHandleProcErrors Then On Error GoTo PROC_ERR
'PushCallStack "CreateNavigatorFile"
'
''
'' Look up the worksheet "User Profile" to determine the location where the JSON layer file should be saved
''
'
'JSONDirectory = ReturnUserProfile("JSON_Directory")
'If JSONDirectory = "" Or Not fso3.FolderExists(JSONDirectory) Then 'No location specified or location does not exist
'    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
'    oFD.Title = "Choose a Location to Save JSON files for the DISARM Navigator"
'    oFD.ButtonName = "Choose"
'    oFD.InitialFileName = Environ("USERPROFILE") & "\"     'sets the folder e.g. C:\Users\steph\
'    oFD.InitialView = msoFileDialogViewLargeIcons
'
'    With oFD
'        If .Show = -1 Then 'if OK is pressed
'            strFolderPath = .SelectedItems(1)
'            Call SetUserProfile("JSON_Directory", .SelectedItems(1))
'        Else
'            GoTo PROC_EXIT ' If user cancels then we do not know where to save the JSON file so exit
'        End If
'    End With
'Else
'    strFolderPath = JSONDirectory ' location specified and exists
'End If
'
'DocumentName = fso2.GetBaseName(Application.ActiveDocument.Name) ' Use the same name for the .json file as the Word document
'
'' Pre-amble contains name of the layer plus some parameters for the ATT&CK Navigator
'Part1 = "{""name"":""" & DocumentName & """,""versions"":{""attack"":""1"",""navigator"":""4.8.2"",""layer"":""4.4""}," & _
'"""domain"":""DISARM"",""description"":"""",""filters"":{""platforms"":[""Windows"",""Linux"",""Mac""]}" & _
'"""sorting"":0,""layout"":{""layout"":""flat"",""aggregateFunction"":""average"",""showID"":true,""showName"":true," & _
'"""showAggregateScores"":false,""countUnscored"":false},""hideDisabled"":false,""techniques"":["
'
'' The next part contains the JSON for all the active techniques tagged by the user
'Part2 = ReturnJSONforTaggedTechniques
'
'If Part2 = "" Then GoTo PROC_EXIT
'
'' The postscript contains some formatting parameters for the ATT&CK Navigator
'Part3 = "],""gradient"":{""colors"":[""#ff6666ff"",""#ffe766ff"",""#8ec843ff""],""minValue"":0,""maxValue"":100},""legendItems"":[],""metadata"":[]," & _
'"""links"":[],""showTacticRowBackground"":false,""tacticRowBackground"":""#dddddd"",""selectTechniquesAcrossTactics"":true,""selectSubtechniquesWithParent"":false}"
'
'' Now combine all three parts
'JsonVBA = Part1 & Part2 & Part3
'
'' Create a JSON Object from the JsonVBA string, convert this to pretty print JSON, then write to file
'Set JSON = JsonConverter.ParseJson(JsonVBA)
'Set tsout = fso.CreateTextFile(strFolderPath & "\" & DocumentName & ".json")
'Call tsout.WriteLine(JsonConverter.ConvertToJson(JSON, Whitespace:=2))
'
'' Inform user that file created successfully
'intMsgReturn = MsgBox("Layer file " & strFolderPath & "\" & DocumentName & ".json created successfully", _
'               vbOKCancel + vbInformation, "DISARM: Create Navigator file")
'
'PROC_EXIT:
'  PopCallStack
'  Exit Sub
'
'PROC_ERR:
'  GlobalErrHandler
'  Resume PROC_EXIT
'End Sub
Function ReturnUserProfile(Parameter As String) As String

'
' Look up and return the value of the given parameter in the worksheet "User Profile"
'

Dim strPath As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnUserProfile"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateSummaryRedUnformatted)"
End If
    
If WorksheetExists("User Profile", oWB_TaggingWorkbook) Then
    Set oWS_UserProfile = oWB_TaggingWorkbook.Worksheets("User Profile")
    ReturnUserProfile = oApp.WorksheetFunction.VLookup(Parameter, oWS_UserProfile.UsedRange, 2, False)
Else
    'On Error Resume Next
    Set oWS_UserProfile = oWB_TaggingWorkbook.Sheets.Add(Before:=oWB_TaggingWorkbook.Sheets(1))
    oWS_UserProfile.Name = "User Profile"
    oWS_UserProfile.Cells(1, 1).Value = "Parameter"
    oWS_UserProfile.Cells(1, 2).Value = "Value"
    oWS_UserProfile.Cells(2, 1).Value = "JSON_Directory"
    oWS_UserProfile.Cells(2, 2).Value = ""
    ReturnUserProfile = ""
End If

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
Sub SetUserProfile(Parameter As String, Value As String)

'
' Set the value of the given parameter in the worksheet "User Profile"
'

Dim strParameterIndex As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "SetUserProfile"

strParameterIndex = oApp.WorksheetFunction.Match(Parameter, oWB_TaggingWorkbook.Worksheets("User Profile").UsedRange.Columns(1), 0)
oWB_TaggingWorkbook.Worksheets("User Profile").Cells(strParameterIndex, 2).Value = Value

'
' Save Tagging Workbook
'

modMain.SaveTaggingWorkbook

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Function ReturnJSONforTaggedTechniques() As String

'
' Return a string with the required JSON formatting for all active techniques in the tagging worksheet
'

Dim lngRows As Long
Dim lngIndex As Long
Dim i As Integer
Dim dblCount As Double
Dim varValue As Variant
Dim arrTableData() As String
Dim JSON As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnJSONforTaggedTechniques"

'
' Allows you to cancel out of an infinite loop using CTRL-C
'

Application.EnableCancelKey = wdCancelInterrupt

If ReturnNumRowsSummaryRed <= 1 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("You have not tagged any techniques for this document", vbOKCancel + vbInformation, "DISARM: Create Navigator file")
    GoTo FUNC_EXIT
End If

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryRedUnformatted has not been set then create _SRU worksheet. If it has been set check that it points to
' the active document. If not create _SRU worksheet for the active document.
'

If oWS_SummaryRedUnformatted Is Nothing Then
    modMain.CreateSummaryRedUnformatted
ElseIf Mid(oWS_SummaryRedUnformatted.Name, InStrRev(oWS_SummaryRedUnformatted.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryRedUnformatted
End If

lngRows = oWS_SummaryRedUnformatted.UsedRange.Rows.Count
varValue = "Active"
dblCount = oApp.WorksheetFunction.CountIf(oWS_SummaryRedUnformatted.UsedRange.Columns(7).Cells, varValue)
If dblCount > 0 Then ' If there are active techniques
    ReDim arrTableData(1 To dblCount, 1 To 2)
    i = 0
    For lngIndex = 2 To lngRows
      'Add a row to the array
      If oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 7) = "Active" Then
        i = i + 1
        arrTableData(i, 1) = oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 3) ' Technique ID
        arrTableData(i, 2) = ReturnJSONTactic(oWS_SummaryRedUnformatted.UsedRange.Cells(lngIndex, 1)) ' Hyphenated Tactic Name
      End If
    Next lngIndex
End If

'
' Create JSON for first technique
'

JSON = "{""techniqueID"":""" & arrTableData(1, 1) & """,""tactic"":""" & arrTableData(1, 2) & """,""score"":1," & _
"""color"":"""",""comment"":"""",""enabled"":true,""metadata"":[],""links"":[],""showSubtechniques"":false}"

'
' Add comma separator and JSON for each subsequent technique
'

For i = 2 To UBound(arrTableData)
    JSON = JSON & ",{""techniqueID"":""" & arrTableData(i, 1) & """,""tactic"":""" & arrTableData(i, 2) & """,""score"":1," & _
    """color"":"""",""comment"":"""",""enabled"":true,""metadata"":[],""links"":[],""showSubtechniques"":false}"
Next i

'
' Return JSON string
'

ReturnJSONforTaggedTechniques = JSON

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
Function ReturnJSONTactic(TacticID As String) As String

'
' Return the Tactic Name for the given Tactic ID in the format expected by the DISARM Navigator
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnJSONTactic"

Select Case TacticID
Case "TA01"
    ReturnJSONTactic = "plan-strategy"
Case "TA02"
    ReturnJSONTactic = "plan-objectives"
Case "TA13"
    ReturnJSONTactic = "target-audience-analysis"
Case "TA14"
    ReturnJSONTactic = "develop-narratives"
Case "TA06"
    ReturnJSONTactic = "develop-content"
Case "TA15"
    ReturnJSONTactic = "establish-social-assets"
Case "TA16"
    ReturnJSONTactic = "establish-legitimacy"
Case "TA05"
    ReturnJSONTactic = "microtarget"
Case "TA07"
    ReturnJSONTactic = "select-channels-and-affordances"
Case "TA08"
    ReturnJSONTactic = "conduct-pump-priming"
Case "TA09"
    ReturnJSONTactic = "deliver-content"
Case "TA17"
    ReturnJSONTactic = "maximise-exposure"
Case "TA18"
    ReturnJSONTactic = "drive-online-harms"
Case "TA10"
    ReturnJSONTactic = "drive-offline-activity"
Case "TA11"
    ReturnJSONTactic = "persist-in-the-information-environment"
Case "TA12"
    ReturnJSONTactic = "assess-effectiveness"
Case Else
    ReturnJSONTactic = ""
End Select

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub CreateSummaryBlueTable()

'
' Create a table within the Word document to house the list of defender countermeasures identified during tagging.
' Was getting error 4605 when trying to insert a break and the cursor was positioned at the end of a row in the red table so move down if this is the case
'
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateSummaryBlueTable"

If Selection.IsEndOfRowMark = True Then Selection.MoveDown

Selection.InsertBreak Type:=wdSectionBreakNextPage

With Selection.PageSetup
    If .Orientation = wdOrientPortrait Then
        .Orientation = wdOrientLandscape
    End If
    .TopMargin = InchesToPoints(1#) 'amended from 0.5 for ALF/debunk demo
    .BottomMargin = InchesToPoints(0.5)
    .LeftMargin = InchesToPoints(1) 'amended from 0.5 for ALF/debunk demo
    .RightMargin = InchesToPoints(0.5)
End With

'
' Create table. Fit table width within the window
'

Set tblSummaryBlue = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:= _
    3, DefaultTableBehavior:=wdWord8TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
tblSummaryBlue.AllowAutoFit = False
tblSummaryBlue.AutoFormat Format:=0

With tblSummaryBlue

    'If .Style <> "Table Grid" Then
    '    .Style = "Table Grid"
    'End If
    '
    '    This code was bombing out with error 5834 "A megadott nevu elem nem létezik"
    '    ("Item with specified name does not exist") in the Hungarian version of Word
    '    so use the numeric value for Style "Table Grid" so that it works for any language
    '    See https://www.msofficeforums.com/word-vba/46753-install-language-macro-language.html
    '
    
    .Style = -155
    
    .AllowAutoFit = False
    .Columns.PreferredWidthType = wdPreferredWidthPoints
    .Columns(1).PreferredWidth = InchesToPoints(2)
    .Columns(2).PreferredWidth = InchesToPoints(0.9)
    .Columns(3).PreferredWidth = InchesToPoints(7#)
    
    With .Rows(1)
    
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(1).Range.Text = "Countermeasure"
        .Cells(2).Range.Text = "ID"
        .Cells(3).Range.Text = "Use"
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorBlueGray
        
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 14
            .Bold = True
            .ColorIndex = wdWhite
        End With
        
    End With
    
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryRedTableTaskHeader(strTaskHeader As String)

'
' Adds a row in the Word summary table with the task header
'

Dim rowNew As Row
Dim celTable As Cell

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryRedTableTaskheader"

With tblSummaryRed
    Set rowNew = .Rows.Add
    intTableRowNumber = intTableRowNumber + 1
    With rowNew
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorPaleBlue
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 12
            .Bold = True
            .Color = wdColorAutomatic
        End With
        .Cells(1).Range.Text = strTaskHeader
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryBlueTableMetatechniqueHeader(strMetatechniqueHeader As String)

'
' Adds a row to the countermeasures summary table with the metatechnique header
'

Dim rowNew As Row
Dim celTable As Cell

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryBlueTableMetatechniqueHeader"

With tblSummaryBlue
    Set rowNew = .Rows.Add
    intBlueTableRowNumber = intBlueTableRowNumber + 1
    With rowNew
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorPaleBlue
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 12
            .Bold = True
            .Color = wdColorAutomatic
        End With
        .Cells(1).Range.Text = strMetatechniqueHeader
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryRedTableTechniques(strTaskID As Variant, dblTaskIDIndex As Double)

'
' Add a row to the Word summary table for each tag belonging to this task
'

Dim dblRowIndex As Double
Dim rowNew As Row
Dim celTable As Cell
Dim lngLenSentence As Long

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryRedTableTechniques"

dblRowIndex = dblTaskIDIndex

With tblSummaryRed
    ' write first technique
    Set rowNew = .Rows.Add
    intTableRowNumber = intTableRowNumber + 1
    With rowNew
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorAutomatic
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 11
            .Bold = False
            .Color = wdColorAutomatic
        End With
        .Cells(1).PreferredWidth = InchesToPoints(2)
        .Cells(2).PreferredWidth = InchesToPoints(0.9)
        .Cells(3).PreferredWidth = InchesToPoints(7#)  'amended from 5.9 for ALF/debunk demo
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cells(1).Range.Text = oWS_SummaryRedUnformatted.Cells(dblRowIndex, 4)
        .Cells(2).Range.Text = oWS_SummaryRedUnformatted.Cells(dblRowIndex, 3)
        lngLenSentence = Len(oWS_SummaryRedUnformatted.Cells(dblRowIndex, 5))
        If lngLenSentence > 0 Then
            .Cells(3).Range.Text = Left(oWS_SummaryRedUnformatted.Cells(dblRowIndex, 5), lngLenSentence - 1)
        Else
            .Cells(3).Range.Text = ""
        End If
    End With
    dblRowIndex = dblRowIndex + 1
    ' write other techniques if found
    Do While oWS_SummaryRedUnformatted.Cells(dblRowIndex, 1) = strTaskID 'And dblRowIndex <= oWS_RowCount_SRU
        If dblRowIndex > oWS_RowCount_SRU Then Exit Do 'failsafe to prevent infinite loop
        Set rowNew = .Rows.Add
        intTableRowNumber = intTableRowNumber + 1
        With rowNew
            .Cells(1).Range.Text = oWS_SummaryRedUnformatted.Cells(dblRowIndex, 4)
            .Cells(2).Range.Text = oWS_SummaryRedUnformatted.Cells(dblRowIndex, 3)
            lngLenSentence = Len(oWS_SummaryRedUnformatted.Cells(dblRowIndex, 5))
            If lngLenSentence > 0 Then
                .Cells(3).Range.Text = Left(oWS_SummaryRedUnformatted.Cells(dblRowIndex, 5), lngLenSentence - 1)
            Else
                .Cells(3).Range.Text = ""
            End If
            .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        dblRowIndex = dblRowIndex + 1
        'Debug.Print "dblRowIndex = " & dblRowIndex
    Loop
End With
    
'ActiveDocument.Save

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryBlueTableCountermeasures(strMetatechniqueID As Variant, dblMetatechniqueIDIndex As Double)

'
' Add a row to the Word summary table for each countermeasure belonging to this metatechnique
'

Dim dblRowIndex As Double
Dim rowNew As Row
Dim celTable As Cell
Dim lngLenSentence As Long

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryBlueTableCountermeasures"

dblRowIndex = dblMetatechniqueIDIndex

With tblSummaryBlue
    ' write first countermeasure
    Set rowNew = .Rows.Add
    intBlueTableRowNumber = intBlueTableRowNumber + 1
    With rowNew
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorAutomatic
        With .Range.Font
            .Name = "Calibri (Body)"
            .Size = 11
            .Bold = False
            .Color = wdColorAutomatic
        End With
       
        .Cells(1).PreferredWidth = InchesToPoints(2)
        .Cells(2).PreferredWidth = InchesToPoints(0.9)
        .Cells(3).PreferredWidth = InchesToPoints(7#)  'amended from 5.9 for ALF/debunk demo
        .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cells(1).Range.Text = oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 4)
        .Cells(2).Range.Text = oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 3)
        lngLenSentence = Len(oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 5))
        If lngLenSentence > 0 Then
            .Cells(3).Range.Text = Left(oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 5), lngLenSentence - 1)
        Else
            .Cells(3).Range.Text = ""
        End If
    End With
    dblRowIndex = dblRowIndex + 1
    ' write other countermeasures if found
    Do While oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 1) = strMetatechniqueID 'And dblRowIndex <= oWS_RowCount_SRU
        If dblRowIndex > oWS_RowCount_SBU Then Exit Do 'failsafe to prevent infinite loop
        Set rowNew = .Rows.Add
        intBlueTableRowNumber = intBlueTableRowNumber + 1
        With rowNew
            .Cells(1).Range.Text = oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 4)
            .Cells(2).Range.Text = oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 3)
            lngLenSentence = Len(oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 5))
            If lngLenSentence > 0 Then
                .Cells(3).Range.Text = Left(oWS_SummaryBlueUnformatted.Cells(dblRowIndex, 5), lngLenSentence - 1)
            Else
                .Cells(3).Range.Text = ""
            End If
            .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        dblRowIndex = dblRowIndex + 1
        'Debug.Print "dblRowIndex = " & dblRowIndex
    Loop
End With
    
'ActiveDocument.Save

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FormatTaskRow(intRowNumber As Integer)

'
' Merge the cells of the row with the specified number and align the text in the center
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FormatTaskRow"

With tblSummaryRed.Rows(intRowNumber)
    .Cells.Merge
    .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FormatMetatechniqueRow(intRowNumber As Integer)

'
' Merge the cells of the row with the specified number and align the text in the center
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FormatMetatechniqueRow"

With tblSummaryBlue.Rows(intRowNumber)
    .Cells.Merge
    .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub AddSpace()

'
' Add a new line to the document e.g. after inserting a summary table or graphic
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "AddSpace"

Selection.EndKey Unit:=wdLine

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub LoadCountersFromExcel(ByRef oListPassed As Object, ByRef strMetatechniqueName As String)

'Early binding. Project requires a reference to the Excel Object Library _
                see:  http://www.word.mvps.org/FAQs/InterDev/EarlyvsLateBinding.htm
'
' This procedure populates the ListBox oListPassed with all the countermeasures that apply to Metatechnique strMetatechniqueName
'
' This was the original procedure for handling oListPassed when it was a list box.
' It has been replaced by LoadCountersFromExcel2 which treats oListPassed as a list view.
' Keeping the procedure here just in case I need to revert to the old code.
'
                
Dim varData As Variant
Dim lngCount As Long
Dim lngPos As Long
Dim strMetatechniqueID As String
Dim strMetatechniqueIndex As Double
Dim strPath As String
Dim i As Integer
  
If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LoadCountersFromExcel"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    '
    ' Cannot use Vlookup to get MetatechniqueID since this is to the left of MetatechniqueName in the worksheet "metatechniques"
    ' So use Match function to find a match for strMetatechniqueName and get relative position then use Index to retrieve strMetatechniqueID
    '
    
    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    strMetatechniqueIndex = oApp.WorksheetFunction.Match(strMetatechniqueName, oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(2), 0)
    strMetatechniqueID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(1), strMetatechniqueIndex)
    
    '
    ' Find out how many countermeasures there are for Metatechnique strMetatechniqueID to set up size of array arrCountermeasures
    '
    
    If oWS_countermeasures Is Nothing Then
        Set oWS_countermeasures = oWB_FrameworkMaster.Worksheets("countermeasures")
    End If
    
    counter = 0
    For i = 1 To oWS_countermeasures.UsedRange.Columns(3).Rows.Count
        '
        ' Use InStr since some countermeasures fall within more than one metatechnique
        '
        lngPos = InStr(oWS_countermeasures.UsedRange.Columns(3).Rows(i), strMetatechniqueID)
        If lngPos > 0 Then
            counter = counter + 1
        End If
    Next i
    ReDim arrCountermeasures(counter)
    
    '
    ' Now populate the array arrCountermeasures with the countermeasures for Metatechnique strMetatechniqueID
    '
    
    counter = 0
    For i = 1 To oWS_countermeasures.UsedRange.Cells(2, 3).End(xlDown).Row
        lngPos = InStr(oWS_countermeasures.Cells(i + 1, 3).Value, strMetatechniqueID)
        If lngPos > 0 Then
            arrCountermeasures(counter) = oWS_countermeasures.Cells(i + 1, 2).Value
            counter = counter + 1
        End If
    Next i
    
End With

'
' Populate the list box oListPassed with the array of countermeasures
'

With oListPassed
    .Clear
    .List = arrCountermeasures
End With
  
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
  
End Sub
Sub LoadCountersFromExcel2(ByRef oListPassed As Object, ByRef strMetatechniqueName As String)


'Early binding. Project requires a reference to the Excel Object Library _
                see:  http://www.word.mvps.org/FAQs/InterDev/EarlyvsLateBinding.htm
'
' This procedure populates the ListView oListPassed with all the countermeasures that apply to Metatechnique strMetatechniqueName
'
' This is a modified version of LoadCountersFromExcel where oListPassed is a List View
' which was an easier way of handling the detail Summary and Ethics guidance for each countermeasure.
'

Dim varData As Variant
Dim lngCount As Long
Dim lngPos As Long
Dim strMetatechniqueID As String
Dim strMetatechniqueIndex As Double
Dim strPath As String
Dim i As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LoadCountersFromExcel2"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    '
    ' Cannot use Vlookup to get MetatechniqueID since this is to the left of MetatechniqueName in the worksheet "metatechniques"
    ' So use Match function to find a match for strMetatechniqueName and get relative position then use Index to retrieve strMetatechniqueID
    '

    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    strMetatechniqueIndex = oApp.WorksheetFunction.Match(strMetatechniqueName, oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(2), 0)
    strMetatechniqueID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(1), strMetatechniqueIndex)

    '
    ' Find out how many countermeasures there are for Metatechnique strMetatechniqueID to set up size of array arrCountermeasures
    '

    If oWS_countermeasures Is Nothing Then
        Set oWS_countermeasures = oWB_FrameworkMaster.Worksheets("countermeasures")
    End If

    counter = 0
    For i = 1 To oWS_countermeasures.UsedRange.Columns(3).Rows.Count
        '
        ' Use InStr since some countermeasures fall within more than one metatechnique
        '
        lngPos = InStr(oWS_countermeasures.UsedRange.Columns(3).Rows(i), strMetatechniqueID)
        If lngPos > 0 Then
            counter = counter + 1
        End If
    Next i
    ReDim arrCountermeasures(1 To 4, counter) 'name, color, ethics, summary

    '
    ' Now populate the array arrCountermeasures with the countermeasures for Metatechnique strMetatechniqueID
    '

    counter = 0
    For i = 1 To oWS_countermeasures.UsedRange.Cells(2, 3).End(xlDown).Row
        lngPos = InStr(oWS_countermeasures.Cells(i + 1, 3).Value, strMetatechniqueID)
        If lngPos > 0 Then
            arrCountermeasures(1, counter) = oWS_countermeasures.Cells(i + 1, 2).Value 'name
            arrCountermeasures(2, counter) = oWS_countermeasures.Cells(i + 1, 17).Value 'color
            arrCountermeasures(3, counter) = oWS_countermeasures.Cells(i + 1, 16).Value 'ethics
            arrCountermeasures(4, counter) = oWS_countermeasures.Cells(i + 1, 4).Value 'summary
            counter = counter + 1
        End If
    Next i
End With

'
' Populate the list view oListPassed with the array of countermeasures
'

Dim itemX As ListItem

With oListPassed
    For i = 0 To counter - 1
       Set itemX = .ListItems.Add(, , arrCountermeasures(1, i)) 'name
       If arrCountermeasures(2, i) = "g" Then
            itemX.ListSubItems.Add , , , 1 ' green triangle
       ElseIf arrCountermeasures(2, i) = "o" Then
            itemX.ListSubItems.Add , , , 2 ' orange triangle
       ElseIf arrCountermeasures(2, i) = "r" Then
            itemX.ListSubItems.Add , , , 3 ' red triangle
       Else
            'do nothing
       End If
       itemX.ListSubItems.Add , , arrCountermeasures(2, i) 'color
       itemX.ListSubItems.Add , , arrCountermeasures(3, i) 'ethics
       itemX.ListSubItems.Add , , arrCountermeasures(4, i) 'summary
    Next i
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT

End Sub
Sub LoadFromExcel(ByRef oListPassed As Object, ByRef strTacticName As String)

'Early binding. Project requires a reference to the Excel Object Library _
                see:  http://www.word.mvps.org/FAQs/InterDev/EarlyvsLateBinding.htm
'
' This procedure populates the List View oListPassed with all the techniques that apply to tactic strTacticName
'
                
Dim varData As Variant
Dim lngCount As Long
Dim strTacticID As String
Dim strTacticIndex As Double
Dim arrTableData() As String
Dim strPath As String
Dim i As Integer
Dim k As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LoadFromExcel"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    '
    ' Cannot use Vlookup to get TacticID since this is to the left of TacticName in the worksheet "tactics"
    ' So use Match function to find a match for strTacticName and get relative position then use Index to retrieve strTacticID
    '
    
    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    strTacticIndex = oApp.WorksheetFunction.Match(strTacticName, oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(2), 0)
    strTacticID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(1), strTacticIndex)
    
    '
    ' Find out how many techniques there are for tactic strTacticID to set up size of array arrTechniques
    '
    
    If oWS_techniques Is Nothing Then
        Set oWS_techniques = oWB_FrameworkMaster.Worksheets("techniques")
    End If
    
    counter = 0
    For i = 1 To oWS_techniques.UsedRange.Columns(4).Rows.Count
        If oWS_techniques.UsedRange.Columns(4).Rows(i) = strTacticID Then
            counter = counter + 1
        End If
    Next i
    ReDim arrTechniques(1 To 3, counter)
    
    '
    ' Now populate the array arrTechniques with the techniques for tactic strTacticID
    '
    
    counter = 0
    For i = 1 To oWS_techniques.UsedRange.Cells(2, 4).End(xlDown).Row
        If oWS_techniques.Cells(i + 1, 4).Value = strTacticID Then
            arrTechniques(1, counter) = oWS_techniques.Cells(i + 1, 1).Value ' technique ID
            arrTechniques(2, counter) = oWS_techniques.Cells(i + 1, 2).Value ' technique name
            arrTechniques(3, counter) = oWS_techniques.Cells(i + 1, 5).Value ' technique summary
            counter = counter + 1
        End If
    Next i
    
End With

If counter > 0 Then
    ReDim arrTableData(1 To counter, 1 To 3)
    For k = 1 To counter
        arrTableData(k, 1) = arrTechniques(1, k) ' ID of technique
        arrTableData(k, 2) = arrTechniques(2, k) ' name of technique
        arrTableData(k, 3) = arrTechniques(3, k) ' summary of technique
    Next
End If

'
' Populate the list view oListPassed with the array of techniques
'

Dim itemX As ListItem

With oListPassed
    For i = 0 To counter
       Set itemX = .ListItems.Add(, , arrTechniques(1, i)) ' ID
       itemX.ListSubItems.Add , , arrTechniques(2, i) ' Name
       itemX.ListSubItems.Add , , arrTechniques(3, i) ' Summary
    Next i
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FillListTechniques2(strSearch As String, strPhaseName As String, strTacticName As String, oListOrComboBox As Object)
    
'
' Populate oListOrComboBox with all techniques whose name matches the search string and whose phase names and
' tactic names match the parameters specified. The parameters for phase and/or tactic may be ommitted,
' in which case oListOrComboBox is populated with all techniques found regardless of phase or tactic.
'
' This is a modified version of FillListTechniques where oListOrComboBox is a List View
' which was an easier way of handling the detail Summary for each technique.
'

Dim c As Excel.Range
Dim firstAddress As String
Dim firstRow As Long
Dim firstCol As Long
Dim strPath As String
Dim i, j, k As Integer
Dim arrTechniques() As String
Dim arrTableData() As String
Dim strPhaseID As String
Dim rngSearch As Excel.Range
Dim lngRow As Long

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FillListTechniques2"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    
    If oWS_techniques Is Nothing Then
        Set oWS_techniques = oWB_FrameworkMaster.Worksheets("techniques")
    End If
    
    i = 0 ' i is count of entries found using Find method
    j = 0 ' j is count of rows that match all three criteria
    
    If bLookInDescriptions Then
        Set rngSearch = oWS_techniques.UsedRange.Columns("B:E")
    Else
        Set rngSearch = oWS_techniques.UsedRange.Columns("B")
    End If
    
    With rngSearch
        Set c = .Find(strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows)
        If Not c Is Nothing Then
            'Debug.Print "Row ", c.Row, " Column ", c.Column
            i = 1
            firstRow = c.Row ' capture row number for first cell found to avoid infinite loop on wraparound
            Do
                If i > 300 Then GoTo PROC_EXIT
                'Debug.Print i, " ", c.Value
                lngRow = c.Row
                firstCol = c.Column
                With oWS_techniques.Rows(c.Row)
                If strPhaseName = "" Then
                    If strTacticName = "" Or strTacticName = ReturnTacticName(.Columns("D").Value) Then
                        j = j + 1 ' Match on all three criteria
                        ReDim Preserve arrTechniques(1 To 4, 1 To j) ' can only redim last dimension
                        arrTechniques(1, j) = .Columns("A").Value ' Column A is Technique ID
                        arrTechniques(2, j) = .Columns("B").Value ' Column B is Technique Name
                        arrTechniques(3, j) = .Columns("D").Value ' Column D is Tactic ID
                        arrTechniques(4, j) = .Columns("E").Value ' Column E is Technique Summary
                        'Debug.Print arrTechniques(1, j), " ", arrTechniques(2, j), " ", arrTechniques(3, j)
                    End If
                Else
                    strPhaseID = ReturnPhaseID(.Columns("D").Value)
                    If strPhaseName = ReturnPhaseName(strPhaseID) Then
                        If strTacticName = "" Or strTacticName = ReturnTacticName(.Columns("D").Value) Then
                            j = j + 1 ' match on all three criteria
                            ReDim Preserve arrTechniques(1 To 4, 1 To j) ' can only redim last dimension
                            arrTechniques(1, j) = .Columns("A").Value ' Column A is Technique ID
                            arrTechniques(2, j) = .Columns("B").Value ' Column B is Technique Name
                            arrTechniques(3, j) = .Columns("D").Value ' Column D is Tactic ID
                            arrTechniques(4, j) = .Columns("E").Value ' Column E is Technique Summary
                            'Debug.Print arrTechniques(1, j), " ", arrTechniques(2, j), " ", arrTechniques(3, j)
                        End If
                    End If
                End If
                End With
                        
                Do
                    Set c = .FindNext(c)
                    'If Not c Is Nothing Then Debug.Print "Row ", c.Row, " Column ", c.Column
                    i = i + 1
                Loop While Not c Is Nothing And c.Row = lngRow And c.Column <> firstCol 'Keep on finding until the row changes so we don't get duplicates for multi-column search
            
            Loop While Not c Is Nothing And c.Row <> firstRow ' break out of loop once we wrap around
        End If
        If j = 0 Then
            bTechniquesFound = False
            Dim intMsgReturn As Integer
            If strPhaseName = "" And strTacticName = "" Then
                intMsgReturn = MsgBox("No techniques found matching search term " & strSearch, _
                    vbOKCancel + vbInformation, "DISARM: Search Techniques and Insert Red Tag")
            Else
                If strPhaseName = "" And strTacticName <> "" Then
                    intMsgReturn = MsgBox("No techniques found for tactic " & strTacticName & _
                        " matching search term " & strSearch, vbOKCancel + vbInformation, _
                        "DISARM: Search Techniques and Insert Red Tag")
                Else
                    If strPhaseName <> "" And strTacticName = "" Then
                        intMsgReturn = MsgBox("No techniques found for phase " & strPhaseName & _
                        " matching search term " & strSearch, vbOKCancel + vbInformation, _
                        "DISARM: Search Techniques and Insert Red Tag")
                    Else
                        intMsgReturn = MsgBox("No techniques found for phase " & strPhaseName & _
                        " and tactic " & strTacticName & " matching search term " & strSearch, _
                        vbOKCancel + vbInformation, "DISARM: Search Techniques and Insert Red Tag")
                    End If
                End If
            End If
            GoTo PROC_EXIT
        Else
            bTechniquesFound = True
        End If
    End With

End With

'
' Populate the list view oListOrComboBox with the array of techniques
'

Dim itemX As ListItem

With oListOrComboBox
    For i = 1 To j
       Set itemX = .ListItems.Add(, , ReturnPhaseName(ReturnPhaseID(arrTechniques(3, i)))) 'phase name
       itemX.ListSubItems.Add , , ReturnTacticName(arrTechniques(3, i)) 'tactic name
       itemX.ListSubItems.Add , , arrTechniques(1, i) 'technique ID
       itemX.ListSubItems.Add , , arrTechniques(2, i) 'technique name
       itemX.ListSubItems.Add , , arrTechniques(4, i) 'technique summary
    Next i
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FillListTechniques(strSearch As String, strPhaseName As String, strTacticName As String, oListOrComboBox As Object)
    
'
' Populate oListOrComboBox with all techniques whose name matches the search string and whose phase names and
' tactic names match the parameters specified. The parameters for phase and/or tactic may be ommitted,
' in which case oListOrComboBox is populated with all techniques found regardless of phase or tactic.
'
' This was the original procedure for handling oListOrComboBox when it was a list box.
' It has been replaced by FillListTechniques2 which treats oListorComboBox as a list view.
' Keeping the procedure here just in case I need to revert to the old code.
'

Dim c As Excel.Range
Dim firstAddress As String
Dim firstRow As Long
Dim firstCol As Long
Dim strPath As String
Dim i, j, k As Integer
Dim arrTechniques() As String
Dim arrTableData() As String
Dim strPhaseID As String
Dim rngSearch As Excel.Range
Dim lngRow As Long

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FillListTechniques"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    
    If oWS_techniques Is Nothing Then
        Set oWS_techniques = oWB_FrameworkMaster.Worksheets("techniques")
    End If
    
    i = 0 ' i is count of entries found using Find method
    j = 0 ' j is count of rows that match all three criteria
    
    If bLookInDescriptions Then
        Set rngSearch = oWS_techniques.UsedRange.Columns("B:E")
    Else
        Set rngSearch = oWS_techniques.UsedRange.Columns("B")
    End If
    
    With rngSearch
        Set c = .Find(strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows)
        If Not c Is Nothing Then
            'Debug.Print "Row ", c.Row, " Column ", c.Column
            i = 1
            firstRow = c.Row ' capture row number for first cell found to avoid infinite loop on wraparound
            Do
                If i > 300 Then GoTo PROC_EXIT
                'Debug.Print i, " ", c.Value
                lngRow = c.Row
                firstCol = c.Column
                With oWS_techniques.Rows(c.Row)
                If strPhaseName = "" Then
                    If strTacticName = "" Or strTacticName = ReturnTacticName(.Columns("D").Value) Then
                        j = j + 1 ' Match on all three criteria
                        ReDim Preserve arrTechniques(1 To 3, 1 To j) ' can only redim last dimension
                        arrTechniques(1, j) = .Columns("A").Value ' Column A is Technique ID
                        arrTechniques(2, j) = .Columns("B").Value ' Column B is Technique Name
                        arrTechniques(3, j) = .Columns("D").Value ' Column D is Tactic ID
                        'Debug.Print arrTechniques(1, j), " ", arrTechniques(2, j), " ", arrTechniques(3, j)
                    End If
                Else
                    strPhaseID = ReturnPhaseID(.Columns("D").Value)
                    If strPhaseName = ReturnPhaseName(strPhaseID) Then
                        If strTacticName = "" Or strTacticName = ReturnTacticName(.Columns("D").Value) Then
                            j = j + 1 ' match on all three criteria
                            ReDim Preserve arrTechniques(1 To 3, 1 To j) ' can only redim last dimension
                            arrTechniques(1, j) = .Columns("A").Value ' Column A is Technique ID
                            arrTechniques(2, j) = .Columns("B").Value ' Column B is Technique Name
                            arrTechniques(3, j) = .Columns("D").Value ' Column D is Tactic ID
                            'Debug.Print arrTechniques(1, j), " ", arrTechniques(2, j), " ", arrTechniques(3, j)
                        End If
                    End If
                End If
                End With
                        
                Do
                    Set c = .FindNext(c)
                    'If Not c Is Nothing Then Debug.Print "Row ", c.Row, " Column ", c.Column
                    i = i + 1
                Loop While Not c Is Nothing And c.Row = lngRow And c.Column <> firstCol 'Keep on finding until the row changes so we don't get duplicates for multi-column search
            
            Loop While Not c Is Nothing And c.Row <> firstRow ' break out of loop once we wrap around
            
            ' Use subscript k to loop round array arrTableData
            
            If j > 0 Then
                ReDim arrTableData(1 To j, 1 To 3)
                For k = 1 To j
                    arrTableData(k, 3) = arrTechniques(2, k) ' name of technique
                    arrTableData(k, 2) = ReturnTacticName(arrTechniques(3, k)) ' name of tactic
                    strPhaseID = ReturnPhaseID(arrTechniques(3, k))
                    arrTableData(k, 1) = ReturnPhaseName(strPhaseID) ' name of phase
                    'Debug.Print arrTableData(k, 1), " ", arrTableData(k, 2), " ", arrTableData(k, 3)
                Next
            End If
        
        End If
        If j = 0 Then
            bTechniquesFound = False
            Dim intMsgReturn As Integer
            If strPhaseName = "" And strTacticName = "" Then
                intMsgReturn = MsgBox("No techniques found matching search term " & strSearch, _
                    vbOKCancel + vbInformation, "DISARM: Search Techniques and Insert Red Tag")
            Else
                If strPhaseName = "" And strTacticName <> "" Then
                    intMsgReturn = MsgBox("No techniques found for tactic " & strTacticName & _
                        " matching search term " & strSearch, vbOKCancel + vbInformation, _
                        "DISARM: Search Techniques and Insert Red Tag")
                Else
                    If strPhaseName <> "" And strTacticName = "" Then
                        intMsgReturn = MsgBox("No techniques found for phase " & strPhaseName & _
                        " matching search term " & strSearch, vbOKCancel + vbInformation, _
                        "DISARM: Search Techniques and Insert Red Tag")
                    Else
                        intMsgReturn = MsgBox("No techniques found for phase " & strPhaseName & _
                        " and tactic " & strTacticName & " matching search term " & strSearch, _
                        vbOKCancel + vbInformation, "DISARM: Search Techniques and Insert Red Tag")
                    End If
                End If
            End If
            GoTo PROC_EXIT
        Else
            bTechniquesFound = True
        End If
    End With

End With

With oListOrComboBox
    .Clear
    'Create the list entries.
    If UBound(arrTableData) >= 1 Then .List = arrTableData
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
 
Sub FillListCountermeasures2(strSearch As String, strMetatechniqueName As String, oListOrComboBox As Object)
    
'
' Populate oListOrComboBox with all Countermeasures whose name matches the search string and whose
' metatechnique name matches the parameter specified. The parameter for metatechnique may be ommitted,
' in which case oListOrComboBox is populated with all countermeasures found regardless of metatechnique.
'
' This is a modified version of FillListCountermeasures where oListOrComboBox is a List View
' which was an easier way of handling the detail Summary and Ethics guidance for each countermeasure.
'

Dim c As Excel.Range
Dim firstAddress As String
Dim firstRow As Long
Dim firstCol As Long
Dim strPath As String
Dim i, j, k, m As Integer
Dim arrCountermeasures() As String
Dim arrTableData() As String
Dim strPhaseID As String
Dim rngSearch As Excel.Range
Dim lngRow As Long
Dim arrMetatechniques() As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FillListCountermeasures2"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    
    If oWS_countermeasures Is Nothing Then
        Set oWS_countermeasures = oWB_FrameworkMaster.Worksheets("countermeasures")
    End If
    
    i = 0 ' i is count of entries found using Find method
    j = 0 ' j is count of rows that match both
    
    If bLookInDescriptionsCountermeasures Then
        Set rngSearch = oWS_countermeasures.UsedRange.Columns("B:D")
    Else
        Set rngSearch = oWS_countermeasures.UsedRange.Columns("B")
    End If
    
    With rngSearch
        Set c = .Find(strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows)

        If Not c Is Nothing Then
            'Debug.Print "Row ", c.Row, " Column ", c.Column
            i = 1
            firstRow = c.Row ' capture row number for first cell found to avoid infinite loop on wraparound
            Do
                If i > 500 Then GoTo PROC_EXIT
                'Debug.Print i, " ", c.Value
                lngRow = c.Row
                firstCol = c.Column
                With oWS_countermeasures.Rows(c.Row)
                    If strMetatechniqueName = "" Then
                        ' There may be more than one metatechnique associated with this countermeasure
                        arrMetatechniques = ReturnMetatechniques(.Columns("C").Value)
                        ' List each combination of metatechnique + countermeasure in the search results
                        For m = 0 To UBound(arrMetatechniques)
                            j = j + 1 ' match on search term
                            ReDim Preserve arrCountermeasures(1 To 6, 1 To j) ' can only redim last dimension
                            arrCountermeasures(1, j) = .Columns("A").Value ' Column A is Countermeasure ID
                            arrCountermeasures(2, j) = .Columns("B").Value ' Column B is Countermeasure Name
                            arrCountermeasures(3, j) = arrMetatechniques(m)
                            arrCountermeasures(4, j) = .Columns("Q").Value ' Column Q is color
                            arrCountermeasures(5, j) = .Columns("P").Value ' Column P is ethics
                            arrCountermeasures(6, j) = .Columns("D").Value ' Column D is summary
                        'Debug.Print arrCountermeasures(1, j), " ", arrCountermeasures(2, j), " ", arrCountermeasures(3, j)
                        Next m
                    Else
                        ' Check for a match on metatechniquename using upper or lowercase to handle
                        ' Master Excel file prior to or after the change to Title Case in August 2023
                        If InStr(.Columns("C").Value, strMetatechniqueName) > 0 Or _
                            InStr(.Columns("C").Value, LCase(strMetatechniqueName)) > 0 Then
                            j = j + 1 ' Match on both search term plus metatechnique
                            ReDim Preserve arrCountermeasures(1 To 6, 1 To j) ' can only redim last dimension
                            arrCountermeasures(1, j) = .Columns("A").Value ' Column A is Countermeasure ID
                            arrCountermeasures(2, j) = .Columns("B").Value ' Column B is Countermeasure Name
                            arrCountermeasures(3, j) = strMetatechniqueName
                            arrCountermeasures(4, j) = .Columns("Q").Value ' Column Q is color
                            arrCountermeasures(5, j) = .Columns("P").Value ' Column P is ethics
                            arrCountermeasures(6, j) = .Columns("D").Value ' Column D is summary
                            'Debug.Print arrCountermeasures(1, j), " ", arrCountermeasures(2, j), " ", arrCountermeasures(3, j)
                        End If
                    End If
                End With
                        
                Do
                    Set c = .FindNext(c)
                    'If Not c Is Nothing Then Debug.Print "Row ", c.Row, " Column ", c.Column
                    i = i + 1
                Loop While Not c Is Nothing And c.Row = lngRow And c.Column <> firstCol 'Keep on finding until the row changes so we don't get duplicates for multi-column search
            
            Loop While Not c Is Nothing And c.Row <> firstRow ' break out of loop once we wrap around
            
            ' Use subscript k to loop round array arrTableData
            ' May not need to set up arrTableData since the code below now populates the list view
            ' directly from arrCountermeasures
            
            If j > 0 Then
                ReDim arrTableData(1 To j, 1 To 2)
                For k = 1 To j
                    arrTableData(k, 2) = arrCountermeasures(2, k) ' Countermeasure name
                    arrTableData(k, 1) = arrCountermeasures(3, k) ' Metatechniques
                    'Debug.Print arrTableData(k, 1), " ", arrTableData(k, 2)
                Next
            End If
        
        End If
        If j = 0 Then
            bCountermeasuresFound = False
            Dim intMsgReturn As Integer
            If strMetatechniqueName = "" Then
                intMsgReturn = MsgBox("No Countermeasures found matching search term " & strSearch, _
                    vbOKCancel + vbInformation, "DISARM: Search Countermeasures and Insert Blue Tag")
            Else
                intMsgReturn = MsgBox("No Countermeasures found for metatechnique " & strMetatechniqueName & _
                    " matching search term " & strSearch, vbOKCancel + vbInformation, _
                    "DISARM: Search Countermeasures and Insert Blue Tag")
            End If
            GoTo PROC_EXIT
        Else
            bCountermeasuresFound = True
        End If
    End With

End With

'
' Populate the list view oListOrComboBox with the array of countermeasures
'

Dim itemX As ListItem

With oListOrComboBox
    For i = 1 To j
       Set itemX = .ListItems.Add(, , arrCountermeasures(3, i)) 'metatechnique
       itemX.ListSubItems.Add , , arrCountermeasures(2, i) 'name
       If arrCountermeasures(4, i) = "g" Then
            itemX.ListSubItems.Add , , , 1 ' green triangle
       ElseIf arrCountermeasures(4, i) = "o" Then
            itemX.ListSubItems.Add , , , 2 ' orange triangle
       ElseIf arrCountermeasures(4, i) = "r" Then
            itemX.ListSubItems.Add , , , 3 ' red triangle
       Else
            'do nothing
       End If
       itemX.ListSubItems.Add , , arrCountermeasures(4, i) 'color
       itemX.ListSubItems.Add , , arrCountermeasures(5, i) 'ethics
       itemX.ListSubItems.Add , , arrCountermeasures(6, i) 'summary
    Next i
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub FillListCountermeasures(strSearch As String, strMetatechniqueName As String, oListOrComboBox As Object)
    
'
' Populate oListOrComboBox with all Countermeasures whose name matches the search string and whose
' metatechnique name matches the parameter specified. The parameter for metatechnique may be ommitted,
' in which case oListOrComboBox is populated with all countermeasures found regardless of metatechnique.
'
' This was the original procedure for handling oListOrComboBox when it was a list box.
' It has been replaced by FillListCountermeasures2 which treats oListorComboBox as a list view.
' Keeping the procedure here just in case I need to revert to the old code.
'

Dim c As Excel.Range
Dim firstAddress As String
Dim firstRow As Long
Dim firstCol As Long
Dim strPath As String
Dim i, j, k, m As Integer
Dim arrCountermeasures() As String
Dim arrTableData() As String
Dim strPhaseID As String
Dim rngSearch As Excel.Range
Dim lngRow As Long
Dim arrMetatechniques() As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FillListCountermeasures"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    If oWB_FrameworkMaster Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_FrameworkMaster = .Workbooks.Open(strPath & cSourceFrameworkMaster)
    End If
    
    If oWS_countermeasures Is Nothing Then
        Set oWS_countermeasures = oWB_FrameworkMaster.Worksheets("countermeasures")
    End If
    
    i = 0 ' i is count of entries found using Find method
    j = 0 ' j is count of rows that match both
    
    If bLookInDescriptionsCountermeasures Then
        Set rngSearch = oWS_countermeasures.UsedRange.Columns("B:D")
    Else
        Set rngSearch = oWS_countermeasures.UsedRange.Columns("B")
    End If
    
    With rngSearch
        Set c = .Find(strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows)

        If Not c Is Nothing Then
            'Debug.Print "Row ", c.Row, " Column ", c.Column
            i = 1
            firstRow = c.Row ' capture row number for first cell found to avoid infinite loop on wraparound
            Do
                If i > 500 Then GoTo PROC_EXIT
                'Debug.Print i, " ", c.Value
                lngRow = c.Row
                firstCol = c.Column
                With oWS_countermeasures.Rows(c.Row)
                    If strMetatechniqueName = "" Then
                        ' There may be more than one metatechnique associated with this countermeasure
                        arrMetatechniques = ReturnMetatechniques(.Columns("C").Value)
                        ' List each combination of metatechnique + countermeasure in the search results
                        For m = 0 To UBound(arrMetatechniques)
                            j = j + 1 ' match on search term
                            ReDim Preserve arrCountermeasures(1 To 3, 1 To j) ' can only redim last dimension
                            arrCountermeasures(1, j) = .Columns("A").Value ' Column A is Countermeasure ID
                            arrCountermeasures(2, j) = .Columns("B").Value ' Column B is Countermeasure Name
                            arrCountermeasures(3, j) = arrMetatechniques(m)
                        'Debug.Print arrCountermeasures(1, j), " ", arrCountermeasures(2, j), " ", arrCountermeasures(3, j)
                        Next m
                    Else
                        ' Check for a match on metatechniquename using upper or lowercase to handle
                        ' Master Excel file prior to or after the change to Title Case in August 2023
                        If InStr(.Columns("C").Value, strMetatechniqueName) > 0 Or _
                            InStr(.Columns("C").Value, LCase(strMetatechniqueName)) > 0 Then
                            j = j + 1 ' Match on both search term plus metatechnique
                            ReDim Preserve arrCountermeasures(1 To 3, 1 To j) ' can only redim last dimension
                            arrCountermeasures(1, j) = .Columns("A").Value ' Column A is Countermeasure ID
                            arrCountermeasures(2, j) = .Columns("B").Value ' Column B is Countermeasure Name
                            arrCountermeasures(3, j) = strMetatechniqueName
                            'Debug.Print arrCountermeasures(1, j), " ", arrCountermeasures(2, j), " ", arrCountermeasures(3, j)
                        End If
                    End If
                End With
                        
                Do
                    Set c = .FindNext(c)
                    'If Not c Is Nothing Then Debug.Print "Row ", c.Row, " Column ", c.Column
                    i = i + 1
                Loop While Not c Is Nothing And c.Row = lngRow And c.Column <> firstCol 'Keep on finding until the row changes so we don't get duplicates for multi-column search
            
            Loop While Not c Is Nothing And c.Row <> firstRow ' break out of loop once we wrap around
            
            ' Use subscript k to loop round array arrTableData
            
            If j > 0 Then
                ReDim arrTableData(1 To j, 1 To 2)
                For k = 1 To j
                    arrTableData(k, 2) = arrCountermeasures(2, k) ' Countermeasure name
                    arrTableData(k, 1) = arrCountermeasures(3, k) ' Metatechniques
                    'Debug.Print arrTableData(k, 1), " ", arrTableData(k, 2)
                Next
            End If
        
        End If
        If j = 0 Then
            bCountermeasuresFound = False
            Dim intMsgReturn As Integer
            If strMetatechniqueName = "" Then
                intMsgReturn = MsgBox("No Countermeasures found matching search term " & strSearch, _
                    vbOKCancel + vbInformation, "DISARM: Search Countermeasures and Insert Blue Tag")
            Else
                intMsgReturn = MsgBox("No Countermeasures found for metatechnique " & strMetatechniqueName & _
                    " matching search term " & strSearch, vbOKCancel + vbInformation, _
                    "DISARM: Search Countermeasures and Insert Blue Tag")
            End If
            GoTo PROC_EXIT
        Else
            bCountermeasuresFound = True
        End If
    End With

End With

With oListOrComboBox
    .Clear
    'Create the list entries.
    If UBound(arrTableData) >= 1 Then .List = arrTableData
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub InsertRowSummaryBlueUnformatted(ByRef strMetatechniqueID As String, ByRef strMetatechniqueName As String, _
                                   ByRef strCountermeasureID As String, ByRef strCountermeasureName As String, _
                                   ByRef strCountermeasureSentence As String, lngCountermeasureSentenceIndex As Long)

'
' Insert a row into the worksheet SummaryBlueUnformatted for the countermeasure selected by the user
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertRowSummaryBlueUnformatted"

If oWS_SummaryBlueUnformatted Is Nothing Then
    Call CreateSummaryBlueUnformatted
End If

Dim oLastRow As Excel.Range
Set oLastRow = oWS_SummaryBlueUnformatted.UsedRange.Rows(oWS_SummaryBlueUnformatted.UsedRange.Rows.Count)
'note this is the whole row, not just the first cell so we use cells(1,1) to get the first cell of the last row
oLastRow.Cells(1, 1).Offset(1, 0).Value = strMetatechniqueID 'offset(1,0) is one row down, same column i.e. 1st column
oLastRow.Cells(1, 1).Offset(1, 1).Value = strMetatechniqueName 'offset(1,1) is one row down, one column across i.e. 2nd column
oLastRow.Cells(1, 1).Offset(1, 2).Value = strCountermeasureID 'offset(1,2) is one row down, two columns across i.e. 3rd column
oLastRow.Cells(1, 1).Offset(1, 3).Value = strCountermeasureName 'offset(1,3) is one row down, three columns across i.e. 4th column
oLastRow.Cells(1, 1).Offset(1, 4).Value = strCountermeasureSentence 'offset(1,4) is one row down, four columns across i.e. 5th column
oLastRow.Cells(1, 1).Offset(1, 5).Value = lngCountermeasureSentenceIndex 'offset(1,5) is one row down, five columns across i.e. 6th column

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertRowSummaryRedUnformatted(ByRef strTacticID As String, ByRef strTacticName As String, _
                                   ByRef strTechniqueID As String, ByRef strTechniqueTitle As String, _
                                   ByRef strTechniqueSentence As String, lngTechniqueSentenceIndex As Long)

'
' Insert a row into the worksheet SummaryRedUnformatted for the technique selected by the user
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertRowSummaryRedUnformatted"

If oWS_SummaryRedUnformatted Is Nothing Then
    Call CreateSummaryRedUnformatted
End If

Dim oLastRow As Excel.Range
Set oLastRow = oWS_SummaryRedUnformatted.UsedRange.Rows(oWS_SummaryRedUnformatted.UsedRange.Rows.Count)
'note this is the whole row, not just the first cell so we use cells(1,1) to get the first cell of the last row
oLastRow.Cells(1, 1).Offset(1, 0).Value = strTacticID 'offset(1,0) is one row down, same column i.e. 1st column
oLastRow.Cells(1, 1).Offset(1, 1).Value = strTacticName 'offset(1,1) is one row down, one column across i.e. 2nd column
oLastRow.Cells(1, 1).Offset(1, 2).Value = strTechniqueID 'offset(1,2) is one row down, two columns across i.e. 3rd column
oLastRow.Cells(1, 1).Offset(1, 3).Value = strTechniqueTitle 'offset(1,3) is one row down, three columns across i.e. 4th column
oLastRow.Cells(1, 1).Offset(1, 4).Value = strTechniqueSentence 'offset(1,4) is one row down, four columns across i.e. 5th column
oLastRow.Cells(1, 1).Offset(1, 5).Value = lngTechniqueSentenceIndex 'offset(1,5) is one row down, five columns across i.e. 6th column
oLastRow.Cells(1, 1).Offset(1, 6).Value = "Active" 'offset(1,6) is one row down, six columns across i.e. 7th column

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Function NoActiveTechniquesLeft(ByRef strTechniqueID As String) As Boolean

'
' Returns True if there are no active techniques in the tagging worksheet for this technique ID
'

Dim c As Excel.Range
Dim firstAddress As String
Dim firstRow As Long
Dim i As Integer
Dim rngUsed As Excel.Range
Dim lngRow As Long

If gcHandleProcErrors Then On Error GoTo FUNC_ERR
PushCallStack "NoActiveTechniquesLeft"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

With oApp

    If oWS_SummaryRedUnformatted Is Nothing Then
        Call CreateSummaryRedUnformatted
    End If
    
    i = 0 ' i is count of entries found using Find method
    
    Set rngUsed = oWS_SummaryRedUnformatted.UsedRange

    With rngUsed
        Set c = .Columns("C").Find(strTechniqueID, LookIn:=xlValues, LookAt:=xlPart)
        If Not c Is Nothing Then
            i = 1
            firstRow = c.Row ' capture row number for first cell found to avoid infinite loop on wraparound
            Do
                If i > 1000 Then GoTo FUNC_EXIT
                'Debug.Print i, " ", c.Value
                lngRow = c.Row
                With oWS_SummaryRedUnformatted.Rows(c.Row)
                    If .Columns("G").Value = "Active" Then
                        NoActiveTechniquesLeft = False
                        GoTo FUNC_EXIT
                    End If
                End With
                i = i + 1
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Row <> firstRow ' break out of loop once we wrap around
            NoActiveTechniquesLeft = True
        Else
            NoActiveTechniquesLeft = True
        End If
        
    End With

End With

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
Sub UnHighlightTechniqueSummaryRedGraphic(ByRef strTechniqueID As String)
                                   
'
' Remove highlighting for the specified technique in the Summary Red Graphic
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UnHighlightTechniqueSummaryRedGraphic"

If oWS_SummaryRedGraphic Is Nothing Then
    Call CreateSummaryRedGraphic
End If

Dim rgFound As Excel.Range
'Set rgFound = oWS_SummaryRedGraphic.UsedRange.Find(strTechniqueID, LookIn:=xlValues, LookAt:=xlPart)
Set rgFound = oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Find(strTechniqueID, LookIn:=xlValues, LookAt:=xlPart)
'Debug.Print rgFound.Address
If Not rgFound Is Nothing Then
    With rgFound.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub HighlightTechniqueSummaryRedGraphic(ByRef strTacticID As String, ByRef strTacticName As String, _
                                   ByRef strTechniqueID As String, ByRef strTechniqueName As String)
                                 
'
' Highlight the specified technique in the Summary Red Graphic
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "HighlightTechniqueSummaryRedGraphic"

If oWS_SummaryRedGraphic Is Nothing Then
    Call CreateSummaryRedGraphic
End If

With oWS_SummaryRedGraphic
    Dim strTacticIDName As String
    Dim strTechniqueIDName As String
    Dim strTacticIndex As Variant
    Dim strTechniqueIndex As String
    Dim intRowCount As Integer
    Dim intColumnCount As Integer
    Dim oSecondRow As Excel.Range
    Dim oLastRow As Excel.Range
    Dim oUsedRangeWithoutPhases As Excel.Range
    
    '
    ' May 2024. Amended code to replace UsedRange with Range("A1").CurrentRegion
    ' See FormatRedGraphic for explanation
    '
    
    'intRowCount = oWS_SummaryRedGraphic.UsedRange.Rows.Count
    intRowCount = oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Rows.Count
    'Set oSecondRow = oWS_SummaryRedGraphic.UsedRange.Rows(2)
    Set oSecondRow = oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Rows(2)
    'Set oLastRow = oWS_SummaryRedGraphic.UsedRange.Rows(intRowCount)
    Set oLastRow = oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Rows(intRowCount)
    intColumnCount = oSecondRow.Columns.Count
    'Set oUsedRangeWithoutPhases = oWS_SummaryRedGraphic.UsedRange.Range(oSecondRow.Cells(1, 1), oLastRow.Cells(intRowCount, intColumnCount))
    Set oUsedRangeWithoutPhases = oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Range(oSecondRow.Cells(1, 1), oLastRow.Cells(intRowCount, intColumnCount))
    strTacticIDName = strTacticID & ": " & strTacticName
    strTechniqueIDName = strTechniqueID & ": " & strTechniqueName
    'strTacticIndex = oApp.WorksheetFunction.Match(strTacticIDName, oWS_SummaryRedGraphic.UsedRange.Rows(2), 0)
    strTacticIndex = oApp.WorksheetFunction.Match(strTacticIDName, oWS_SummaryRedGraphic.Range("A1").CurrentRegion.Rows(2), 0)
    strTechniqueIndex = oApp.WorksheetFunction.Match(strTechniqueIDName, oUsedRangeWithoutPhases.Columns(strTacticIndex), 0)
    
    With oUsedRangeWithoutPhases.Cells(strTechniqueIndex, strTacticIndex).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        If lngSetRedGraphicColor <> 0 Then
             .Color = lngSetRedGraphicColor
        Else
            .Color = 65535 ' default to yellow highlighting
        End If
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Function ReturnMetatechniqueID(strMetatechniqueName As String) As String

'
' Return the Metatechnique ID for a given Metatechnique Name
'

Dim strMetatechniqueID As String
Dim strMetatechniqueIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnMetatechniqueID"

strMetatechniqueIndex = oApp.WorksheetFunction.Match(strMetatechniqueName, oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(2), 0)
strMetatechniqueID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("metatechniques").UsedRange.Columns(1), strMetatechniqueIndex)

ReturnMetatechniqueID = strMetatechniqueID

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTacticID(strTacticName As String) As String

'
' Return the tactic ID for a given Tactic Name
'

Dim strTacticID As String
Dim strTacticIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTacticID"

strTacticIndex = oApp.WorksheetFunction.Match(strTacticName, oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(2), 0)
strTacticID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(1), strTacticIndex)

ReturnTacticID = strTacticID

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTacticName(strTacticID As String) As String

'
' Return the tactic Name for a given Tactic ID
'

Dim strTacticName As String
Dim strTacticIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTacticName"

strTacticIndex = oApp.WorksheetFunction.Match(strTacticID, oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(1), 0)
strTacticName = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(2), strTacticIndex)

ReturnTacticName = strTacticName

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnPhaseID(strTacticID As String) As String

'
' Return the Phase ID for a given Tactic ID
'

Dim strPhaseID As String
Dim strTacticIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnPhaseID"

strTacticIndex = oApp.WorksheetFunction.Match(strTacticID, oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(1), 0)
strPhaseID = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("tactics").UsedRange.Columns(4), strTacticIndex)

ReturnPhaseID = strPhaseID

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnPhaseName(strPhaseID As String) As String

'
' Return the Phase Name for a given Phase ID
'

Dim strPhaseName As String
Dim strPhaseIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnPhaseName"

strPhaseIndex = oApp.WorksheetFunction.Match(strPhaseID, oWB_FrameworkMaster.Worksheets("phases").UsedRange.Columns(1), 0)
strPhaseName = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("phases").UsedRange.Columns(2), strPhaseIndex)

ReturnPhaseName = strPhaseName

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnCountermeasureID(strCountermeasureName As String, strMetatechniqueID As String) As String

'
' Return the Countermeasure ID for a given Countermeasure Name
'

Dim strCountermeasureID As Variant
Dim strCountermeasureIndex As String
Dim aflUsedRangeforMetatechnique As Excel.AutoFilter
Dim rngUsedRangeforMetatechnique As Excel.Range
Dim strPath As String
Dim strSQL As String
Dim lngPos As Long
Dim i As Long

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnCountermeasureID"

If oWS_countermeasures Is Nothing Then
        Set oWS_countermeasures = oWB_FrameworkMaster.Worksheets("countermeasures")
End If
    
For i = 1 To oWS_countermeasures.UsedRange.Columns(2).Rows.Count
    If oWS_countermeasures.UsedRange.Columns(2).Rows(i) = strCountermeasureName Then
        lngPos = InStr(oWS_countermeasures.UsedRange.Columns(3).Rows(i), strMetatechniqueID)
        If lngPos > 0 Then
            ReturnCountermeasureID = oWS_countermeasures.UsedRange.Columns(1).Rows(i)
            GoTo FUNC_EXIT
        End If
    End If
Next i
    
FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTechniqueID(strTechniqueName As String, strTacticID As String) As String

'
' Return the technique ID for a given Technique Name
'

Dim strTechniqueID As Variant
Dim strTechniqueIndex As String
Dim aflUsedRangeforTactic As Excel.AutoFilter
Dim rngUsedRangeforTactic As Excel.Range
Dim strPath As String
Dim strSQL As String
Dim i As Long

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTechniqueID"

If oWS_techniques Is Nothing Then
        Set oWS_techniques = oWB_FrameworkMaster.Worksheets("techniques")
End If
    
For i = 1 To oWS_techniques.UsedRange.Columns(2).Rows.Count
    If oWS_techniques.UsedRange.Columns(2).Rows(i) = strTechniqueName Then
        If oWS_techniques.UsedRange.Columns(4).Rows(i) = strTacticID Then
            ReturnTechniqueID = oWS_techniques.UsedRange.Columns(1).Rows(i)
            GoTo FUNC_EXIT
        End If
    End If
Next i

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTechniqueName(strTechniqueID As String) As String

'
' Return the technique Name for a given Technique ID
'

Dim strTechniqueName As String
Dim strTechniqueIndex As String

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTechniqueName"

strTechniqueIndex = oApp.WorksheetFunction.Match(strTechniqueID, oWB_FrameworkMaster.Worksheets("techniques").UsedRange.Columns(1), 0)
strTechniqueName = oApp.WorksheetFunction.Index(oWB_FrameworkMaster.Worksheets("techniques").UsedRange.Columns(2), strTechniqueIndex)

ReturnTechniqueName = strTechniqueName

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnCountermeasureSentence(lngCountermeasureSentenceIndex) As String

'
' Return the text for the sentence being tagged with countermeasure(s)
'

Dim selCurrent As Range
Dim bSelIsShape As Boolean

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnCountermeasureSentence"

Set selCurrent = Selection.Range
bSelIsShape = (Selection.Type = wdSelectionShape)

'
' Make sure tagging worksheets are available
'

If oWS_SummaryRedUnformatted Is Nothing Then
    Call CreateSummaryRedUnformatted
End If
If oWS_SummaryBlueUnformatted Is Nothing Then
    Call CreateSummaryBlueUnformatted
End If

'
' Look up the red tagging worksheet to see if the sentence has already been tagged with a red tag. If so return the text of the original sentence.
' If not, look up the blue tagging worksheet to see if the sentence has already been tagged with a blue tag. If so return the original sentence.
' Otherwise find the text of the sentence where the current pointer sits and return that.
'
On Error Resume Next
Dim strSentenceFoundIndex As Variant ' the index for the row in the tagging worksheet where the sentence was found
strSentenceFoundIndex = oApp.WorksheetFunction.Match(lngCountermeasureSentenceIndex, oWS_SummaryRedUnformatted.Columns(6), 0)
If oApp.WorksheetFunction.IsNA(strSentenceFoundIndex) Or oApp.WorksheetFunction.IsNumber(strSentenceFoundIndex) = False Then 'Not found in red tagging worksheet
    strSentenceFoundIndex = oApp.WorksheetFunction.Match(lngCountermeasureSentenceIndex, oWS_SummaryBlueUnformatted.Columns(6), 0)
    If oApp.WorksheetFunction.IsNA(strSentenceFoundIndex) <> 0 Or oApp.WorksheetFunction.IsNumber(strSentenceFoundIndex) = False Then 'not found in blue tagging worksheet
        If Right(ActiveDocument.Sentences(lngCountermeasureSentenceIndex).Text, 1) = "." Then
            ReturnCountermeasureSentence = ActiveDocument.Sentences(lngCountermeasureSentenceIndex).Text
        Else
            ReturnCountermeasureSentence = Left(ActiveDocument.Sentences(lngCountermeasureSentenceIndex).Text, _
                                            Len(ActiveDocument.Sentences(lngCountermeasureSentenceIndex).Text) - 1) ' remove paragraph mark
        End If
                                        
    Else
        ReturnCountermeasureSentence = oWS_SummaryBlueUnformatted.Cells(strSentenceFoundIndex, 5).Value ' found in blue tagging worksheet
    End If
Else
    ReturnCountermeasureSentence = oWS_SummaryRedUnformatted.Cells(strSentenceFoundIndex, 5).Value ' found in red tagging worksheet
End If

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTechniqueSentence(lngTechniqueSentenceIndex) As String

'
' Return the text for the sentence being tagged
'

Dim selCurrent As Range
Dim bSelIsShape As Boolean

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTechniqueSentence"

Set selCurrent = Selection.Range
bSelIsShape = (Selection.Type = wdSelectionShape)

'
' Make sure tagging worksheets are available
'

If oWS_SummaryRedUnformatted Is Nothing Then
    Call CreateSummaryRedUnformatted
End If
If oWS_SummaryBlueUnformatted Is Nothing Then
    Call CreateSummaryBlueUnformatted
End If

'
' Look up the red tagging worksheet to see if the sentence has already been tagged with a red tag. If so return the text of the original sentence.
' If not, look up the blue tagging worksheet to see if the sentence has already been tagged with a blue tag. If so return the original sentence.
' Otherwise find the text of the sentence where the current pointer sits and return that.
'

On Error Resume Next
Dim strSentenceFoundIndex As Variant ' the index for the row in the tagging worksheet where the sentence was found
strSentenceFoundIndex = oApp.WorksheetFunction.Match(lngTechniqueSentenceIndex, oWS_SummaryRedUnformatted.Columns(6), 0)
If oApp.WorksheetFunction.IsNA(strSentenceFoundIndex) Or oApp.WorksheetFunction.IsNumber(strSentenceFoundIndex) = False Then 'not found in red tagging worksheet
    strSentenceFoundIndex = oApp.WorksheetFunction.Match(lngTechniqueSentenceIndex, oWS_SummaryBlueUnformatted.Columns(6), 0)
    If oApp.WorksheetFunction.IsNA(strSentenceFoundIndex) Or oApp.WorksheetFunction.IsNumber(strSentenceFoundIndex) = False Then 'not found in blue tagging worksheet
        If Right(ActiveDocument.Sentences(lngTechniqueSentenceIndex).Text, 1) = "." Then
            ReturnTechniqueSentence = ActiveDocument.Sentences(lngTechniqueSentenceIndex).Text
        Else
            ReturnTechniqueSentence = Left(ActiveDocument.Sentences(lngTechniqueSentenceIndex).Text, _
                                        Len(ActiveDocument.Sentences(lngTechniqueSentenceIndex).Text) - 1) ' remove paragraph mark
        End If
    Else
        ReturnTechniqueSentence = oWS_SummaryBlueUnformatted.Cells(strSentenceFoundIndex, 5).Value ' found in blue tagging worksheet
    End If
Else
    ReturnTechniqueSentence = oWS_SummaryRedUnformatted.Cells(strSentenceFoundIndex, 5).Value ' found in red tagging worksheet
End If

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnCountermeasureSentenceIndex()

'
' Return the sentence index for the sentence in which the countermeasure is being tagged
'

Dim i As Integer

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnCountermeasureSentenceIndex"

For i = 1 To ActiveDocument.Sentences.Count
    If Selection.Range.InRange(ActiveDocument.Sentences(i)) Then
        Exit For
    End If
Next i

ReturnCountermeasureSentenceIndex = i

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnTechniqueSentenceIndex()

'
' Return the sentence index for the sentence in which the technique is being tagged
'

Dim i As Integer

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnTechniqueSentenceIndex"

For i = 1 To ActiveDocument.Sentences.Count
    If Selection.Range.InRange(ActiveDocument.Sentences(i)) Then
        Exit For
    End If
Next i

ReturnTechniqueSentenceIndex = i

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function TaskIDFound(ByVal strTaskID As String) As Boolean

'
' Return true of task found otherwise return false
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "TaskIDFound"

On Error Resume Next
Dim dblRowIndex As Double
dblRowIndex = Excel.WorksheetFunction.Match(strTaskID, oWS_SummaryRedUnformatted.Columns(1), 0)
If Err.Number = 0 Then
    TaskIDFound = True
Else
    TaskIDFound = False
End If

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function TaskIDIndex(ByVal strTaskID As String) As Double

'
' Return the index of the first row in which the Task ID is found or 0 if not found
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "TaskIDIndex"

Dim dblRowIndex As Variant

'
' Assume rows have been sorted by Status column first, so "Active" tabs come before "Deleted" tags.
' Find the first row of deleted tags i.e. where Status = "Active"
'

dblRowIndex = oApp.Match("Deleted", oWS_SummaryRedUnformatted.Columns(7), 0)
If Not IsError(dblRowIndex) Then ' There are deleted tags so find the index of the first row within the Active tags only
    dblRowIndex = oApp.Match(strTaskID, oWS_SummaryRedUnformatted.Range("A1:A" & (dblRowIndex - 1)), 0)
    'Debug.Print "Match A1:A ", strTaskID, " Error Number ", Err.Number, " Description ", Err.Description
    If Not IsError(dblRowIndex) Then
        TaskIDIndex = dblRowIndex
    Else
        TaskIDIndex = 0
    End If
Else ' There are no deleted tags so find the first row within the whole column for TaskID
    dblRowIndex = oApp.Match(strTaskID, oWS_SummaryRedUnformatted.Columns(1), 0)
    'Debug.Print "Match ", strTaskID, " Error Number ", Err.Number, " Description ", Err.Description
    If Not IsError(dblRowIndex) Then
        TaskIDIndex = dblRowIndex
    Else
        TaskIDIndex = 0
    End If
End If

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function MetatechniqueIDIndex(ByVal strMetatechniqueID As String) As Double

'
' Return the index of the first row in which the Metatechnique ID is found or 0 if not found
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "MetatechniqueIDIndex"

On Error Resume Next
Dim dblRowIndex As Double
dblRowIndex = Excel.WorksheetFunction.Match(strMetatechniqueID, oWS_SummaryBlueUnformatted.Columns(1), 0)
If Err.Number = 0 Then
    MetatechniqueIDIndex = dblRowIndex
Else
    MetatechniqueIDIndex = 0
End If

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnNumRowsSummaryRed() As Long

'
' Return the number of rows in the summary blue table
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnNumRowsSummaryRed"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryBlueUnformatted has not been set then create _SBU worksheet. If it has been set check that it points to
' the active document. If not create _SBU worksheet for the active document.
'

If oWS_SummaryRedUnformatted Is Nothing Then
    modMain.CreateSummaryRedUnformatted
ElseIf Mid(oWS_SummaryRedUnformatted.Name, InStrRev(oWS_SummaryRedUnformatted.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryRedUnformatted
End If

ReturnNumRowsSummaryRed = oWS_SummaryRedUnformatted.UsedRange.Rows.Count

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub SortSummaryRedUnformatted()

'
' Sort the list of techniques by task ID, technique ID, and sentence index
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "SortSummaryRedUnformatted"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryRedUnformatted has not been set then create _SRU worksheet. If it has been set check that it points to
' the active document. If not create _SRU worksheet for the active document.
'

If oWS_SummaryRedUnformatted Is Nothing Then
    modMain.CreateSummaryRedUnformatted
ElseIf Mid(oWS_SummaryRedUnformatted.Name, InStrRev(oWS_SummaryRedUnformatted.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryRedUnformatted
End If

oWS_SummaryRedUnformatted.Sort.SortFields.Clear
' Sort first by Status, then by TacticID, TechniqueID, and SentenceIndex
oWS_SummaryRedUnformatted.UsedRange.Sort key1:=oWS_SummaryRedUnformatted.Range("G1"), order1:=xlAscending, Header:=xlYes
Dim dblRowIndex As Double
On Error Resume Next
'
' Assume rows have been sorted by Status column first, so "Active" tabs come before "Deleted" tags.
' Find the first row of deleted tags i.e. where Status = "Active"
'
dblRowIndex = oApp.WorksheetFunction.Match("Deleted", oWS_SummaryRedUnformatted.Columns(7), 0)
If Err.Number = 0 Then ' there are deleted tags
    If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
    If dblRowIndex > 3 Then ' there are at least two Active tags so sort the Active tags
        oWS_SummaryRedUnformatted.Range("A1:G" & (dblRowIndex - 1)).Sort _
        key1:=oWS_SummaryRedUnformatted.Range("A1"), order1:=xlAscending, _
        key2:=oWS_SummaryRedUnformatted.Range("C1"), order2:=xlAscending, _
        key3:=oWS_SummaryRedUnformatted.Range("F1"), order3:=xlAscending, _
        Header:=xlYes
    Else ' too few Active tags to sort
    ' Do nothing
    End If
Else ' no deleted tags so just sort the whole usedrange
    'On Error GoTo 0
    If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0
    oWS_SummaryRedUnformatted.UsedRange.Sort _
        key1:=oWS_SummaryRedUnformatted.Range("A1"), order1:=xlAscending, _
        key2:=oWS_SummaryRedUnformatted.Range("C1"), order2:=xlAscending, _
        key3:=oWS_SummaryRedUnformatted.Range("F1"), order3:=xlAscending, _
        Header:=xlYes
End If
oWS_RowCount_SRU = oWS_SummaryRedUnformatted.UsedRange.Rows.Count

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Function ReturnNumRowsSummaryBlue() As Long

'
' Return the number of rows in the summary blue table
'

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnNumRowsSummaryBlue"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryBlueUnformatted has not been set then create _SBU worksheet. If it has been set check that it points to
' the active document. If not create _SBU worksheet for the active document.
'

If oWS_SummaryBlueUnformatted Is Nothing Then
    modMain.CreateSummaryBlueUnformatted
ElseIf Mid(oWS_SummaryBlueUnformatted.Name, InStrRev(oWS_SummaryBlueUnformatted.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryBlueUnformatted
End If

ReturnNumRowsSummaryBlue = oWS_SummaryBlueUnformatted.UsedRange.Rows.Count

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub SortSummaryBlueUnformatted()

'
' Sort the list of countermeasures by metatechniqueID, countermeasure ID and sentence index
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "SortSummaryBlueUnformatted"

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryBlueUnformatted has not been set then create _SBU worksheet. If it has been set check that it points to
' the active document. If not create _SBU worksheet for the active document.
'

If oWS_SummaryBlueUnformatted Is Nothing Then
    modMain.CreateSummaryBlueUnformatted
ElseIf Mid(oWS_SummaryBlueUnformatted.Name, InStrRev(oWS_SummaryBlueUnformatted.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryBlueUnformatted
End If

oWS_SummaryBlueUnformatted.Sort.SortFields.Clear
oWS_SummaryBlueUnformatted.UsedRange.Sort _
    key1:=oWS_SummaryBlueUnformatted.Range("A1"), order1:=xlAscending, _
    key2:=oWS_SummaryBlueUnformatted.Range("C1"), order2:=xlAscending, _
    key3:=oWS_SummaryBlueUnformatted.Range("F1"), order3:=xlAscending, _
    Header:=xlYes
oWS_RowCount_SBU = oWS_SummaryBlueUnformatted.UsedRange.Rows.Count

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Function ReturnPhaseTaskArray() As Variant

'
' Return an array of tasks by phase sequenced by the DISARM kill chain (not ordered by task ID!)
' Look up the summary red graphic to get the DISARM kill chain
'

Dim arrPhase(1 To 4) As Variant
Dim arrPhaseTask()
Dim strPhase As String
Dim strTask As String
Dim i As Integer
Dim j As Integer

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnPhaseTaskArray"

arrPhase(1) = "Plan"
arrPhase(2) = "Prepare"
arrPhase(3) = "Execute"
arrPhase(4) = "Assess"

With oWS_SummaryRedGraphic
    j = 1
    strPhase = .Cells(1, j).Value
    For i = 1 To 4
        Do While (strPhase = arrPhase(i) Or strPhase = "") And j < 17
            ReDim Preserve arrPhaseTask(1 To 2, 1 To j)
            strTask = .Cells(2, j).Value
            arrPhaseTask(1, j) = arrPhase(i)
            arrPhaseTask(2, j) = strTask
            j = j + 1
            strPhase = .Cells(1, j + 1).Value
        Loop
    Next i
End With

ReturnPhaseTaskArray = arrPhaseTask

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function ReturnMetatechniques(strMetatechniques As String) As String()

'
' Extracts metatechnique names from a comma-separated list of metatechniques
' e.g. "M009 - Dilution, M008 - Data Pollution" would yield the array "Dilution, Data Pollution"
'

Dim arrMetatechniques() As String
Dim arrMetatechniqueNames() As String
Dim strMetatechniqueID As String
Dim strMetatechniqueName As String
Dim i As Integer

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnMetatechniques"

arrMetatechniques = Split(strMetatechniques, ",")
ReDim arrMetatechniqueNames(0 To UBound(arrMetatechniques))

For i = 0 To UBound(arrMetatechniques)
    arrMetatechniqueNames(i) = LTrim(Right(arrMetatechniques(i), Len(arrMetatechniques(i)) - 7))
Next i

ReturnMetatechniques = arrMetatechniqueNames

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function
Function ReturnMetatechniqueArray() As Variant

'
' Return an array of metatechniques with ID and name
'

Dim arrMetatechnique()
Dim strMetatechniqueID As String
Dim strMetatechniqueName As String
Dim i As Integer

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "ReturnMetatechniqueArray"

If oApp Is Nothing Then
    On Error Resume Next
    Set oApp = GetObject(, "Excel.Application")
    If Err Then
        If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0
        bStartApp = True
        Set oApp = New Excel.Application
    Else
        If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0
    End If
End If

If oWB_FrameworkMaster Is Nothing Then
    Dim strPath As String
    strPath = Environ("USERPROFILE") & cPathXlstart
    Set oWB_FrameworkMaster = oApp.Workbooks.Open(strPath & cSourceFrameworkMaster)
End If

If oWS_metatechniques Is Nothing Then
    Set oWS_metatechniques = oWB_FrameworkMaster.Worksheets("metatechniques")
End If
    
With oWS_metatechniques
    For i = 2 To .UsedRange.Cells(1, 1).End(xlDown).Row
        ReDim Preserve arrMetatechnique(1 To 2, 1 To i - 1)
        strMetatechniqueID = .Cells(i, 1).Value
        strMetatechniqueName = .Cells(i, 2).Value
        arrMetatechnique(1, i - 1) = strMetatechniqueID
        arrMetatechnique(2, i - 1) = strMetatechniqueName
    Next i
End With

ReturnMetatechniqueArray = arrMetatechnique

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub FormatRedGraphic()

'
' Format the Red Graphic before displaying it
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "FormatRedGraphic"

Application.ScreenUpdating = False

Dim lngLstCol As Long, lngLstRow As Long

Dim intRowCount As Integer
Dim intColumnCount As Integer
Dim oThirdRow As Excel.Range
Dim oLastRow As Excel.Range
Dim oRngCell As Excel.Range
Dim oUsedRangeTechniquesOnly As Excel.Range

If oApp Is Nothing Then
    modMain.InitializeExcelAndOpenWorkbooks
    modMain.CreateTaggingSheets
End If

'
' If oWS_SummaryRedGraphic has not been set then create _SRG worksheet. If it has been set check that it points to
' the active document. If not create _SRG worksheet for the active document.
'

If oWS_SummaryRedGraphic Is Nothing Then
    modMain.CreateSummaryRedGraphic
ElseIf Mid(oWS_SummaryRedGraphic.Name, InStrRev(oWS_SummaryRedGraphic.Name, "_") - 14, 14) <> Right(ActiveDocument.Variables("DISARM_Name"), 14) Then
    modMain.CreateSummaryRedGraphic
End If

With oWS_SummaryRedGraphic
    .Activate
    oApp.ActiveWindow.DisplayGridLines = False
    
    '
    ' May 2024. Amended the following code to replace UsedRange with Range("A1").CurrentRegion
    ' since for some reason UsedRange kept growing when the user repeatedly inserted the Graphic,
    ' which then kept getting smaller and smaller!
    ' See https://www.mrexcel.com/board/threads/usedrange-adding-range-that-is-not-used.1136929/
    '
    
    'intRowCount = .UsedRange.Rows.Count
    intRowCount = .Range("A1").CurrentRegion.Rows.Count
    'Set oThirdRow = .UsedRange.Rows(3)
    Set oThirdRow = .Range("A1").CurrentRegion.Rows(3)
    'Set oLastRow = .UsedRange.Rows(intRowCount)
    Set oLastRow = .Range("A1").CurrentRegion.Rows(intRowCount)
    'intColumnCount = .UsedRange.Columns.Count
    intColumnCount = .Range("A1").CurrentRegion.Columns.Count
    'Set oUsedRangeTechniquesOnly = .UsedRange.Range(oThirdRow.Cells(1, 1), oLastRow.Cells(intRowCount, intColumnCount))
    Set oUsedRangeTechniquesOnly = .Range("A1").CurrentRegion.Range(oThirdRow.Cells(1, 1), oLastRow.Cells(intRowCount, intColumnCount))
    
    For Each oRngCell In oUsedRangeTechniquesOnly.Cells
        If oRngCell.Value = "" Then
            oRngCell.ClearFormats
        End If
    Next
End With

Application.ScreenUpdating = True

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CopyRedGraphic()

'
' Copy the Red Graphic from Excel to he clipboard
'

Dim intRowCount As Integer
Dim intColumnCount As Integer
Dim oUsedRangeExpanded As Excel.Range

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CopyRedGraphic"

With oWS_SummaryRedGraphic

    '
    ' May 2024. Changed UsedRange to Range("A1").CurrentRegion. See above.
    '
    
    'intRowCount = .UsedRange.Rows.Count
    intRowCount = .Range("A1").CurrentRegion.Rows.Count
    'intColumnCount = .UsedRange.Columns.Count
    intColumnCount = .Range("A1").CurrentRegion.Columns.Count
    'Set oUsedRangeExpanded = .UsedRange.Range(.Cells(1, 1), .Cells(intRowCount + 1, intColumnCount + 1))
    Set oUsedRangeExpanded = .Range("A1").CurrentRegion.Range(.Cells(1, 1), .Cells(intRowCount + 1, intColumnCount + 1))
    oUsedRangeExpanded.Copy
End With

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean

'
' Check if a worksheet exists
'

Dim sht As Worksheet

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
PushCallStack "WorksheetExists"

If wb Is Nothing Then Set wb = oApp.ThisWorkbook
On Error Resume Next
Set sht = wb.Sheets(shtName)

If gcHandleFuncErrors Then On Error GoTo FUNC_ERR Else On Error GoTo 0
WorksheetExists = Not sht Is Nothing

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub SaveTaggingWorkbook()

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "SaveTaggingWorkbook"

On Error Resume Next
oApp.DisplayAlerts = False
oWB_TaggingWorkbook.Save

'Debug.Print "Saving tagging workbook", Err.Number, " ", Err.Description, " (SaveTaggingWorkbook)"
oApp.DisplayAlerts = True

If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CleanUp()

'
' Clean up when closing the last word document being tagged. Close Excel workbooks and close Excel. Was getting a dialog with the message
' "There is a large amount of information on the clipboard. Do you want to be able to paste this information into another program later?"
' when closing the tagging workbook so setting DisplayAlerts to none for Word, closing the workbooks, and then restoring the original setting.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CleanUp"

On Error Resume Next
Dim lngDisplayAlerts As Long
lngDisplayAlerts = Application.DisplayAlerts
Application.DisplayAlerts = wdAlertsNone
oWB_FrameworkMaster.Close Savechanges:=False
Set oWB_FrameworkMaster = Nothing
oWB_TaggingWorkbook.Close Savechanges:=True
'Debug.Print "Tagging Workbook closed and saved (CleanUp)"
Set oWB_TaggingWorkbook = Nothing
oApp.ScreenUpdating = True
'oApp.Calculation = xlAutomatic
oApp.WindowState = iWindowState
oApp.DisplayAlerts = True
If bStartApp Then oApp.Quit
Set oApp = Nothing
'Debug.Print "Set oApp = Nothing"
Application.DisplayAlerts = lngDisplayAlerts

If gcHandleProcErrors Then On Error GoTo PROC_ERR Else On Error GoTo 0

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub Update_DISARM_Red()

'
' Update the worksheets in the Tagging workbook used to create color-coded behavioral profiles
'

Dim strTaskID As Variant
Dim strPrefix As String
Dim strPath As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "Update_DISARM_Red"
  
'
' Allows you to cancel out of an infinite loop using CTRL-C
'
Application.EnableCancelKey = wdCancelInterrupt

Dim arrPhaseTask As Variant
arrPhaseTask = ReturnPhaseTaskArray()

strPrefix = DISARM_Name()

If oWB_TaggingWorkbook Is Nothing Then
        strPath = Environ("USERPROFILE") & cPathXlstart
        Set oWB_TaggingWorkbook = oApp.Workbooks.Open(strPath & cSourceTaggingWorkbook)
        'Debug.Print "New Workbook object oWB_TaggingWorkbook (CreateSummaryRedGraphic)"
End If

'
' Clear contents of all existing techniques
'

Set oWS_DISARM_Red_with_IDs = oWB_TaggingWorkbook.Sheets("DISARM Red with IDs")
Set oWS_DISARM_Red_no_IDs = oWB_TaggingWorkbook.Sheets("DISARM Red no IDs")

For i = 1 To 16
    For k = 3 To oWS_DISARM_Red_with_IDs.UsedRange.Rows.Count
            oWS_DISARM_Red_with_IDs.Cells(k, i).ClearContents
    Next k
Next i

For i = 1 To 16
    For k = 3 To oWS_DISARM_Red_no_IDs.UsedRange.Rows.Count
            oWS_DISARM_Red_no_IDs.Cells(k, i).ClearContents
            oWS_DISARM_Red_no_IDs.Cells(k, i).Borders.LineStyle = xlNone
    Next k
Next i

'
' Populate with new techniques
'

For i = 1 To 16
    strTaskID = Left(arrPhaseTask(2, i), 4)
    k = 3
    For j = 2 To oWS_techniques.UsedRange.Rows.Count
        If oWS_techniques.Cells(j, 4).Value = strTaskID Then
            oWS_DISARM_Red_with_IDs.Cells(k, i).Value = oWS_techniques.Cells(j, 1).Value & _
                ": " & oWS_techniques.Cells(j, 2).Value
            oWS_DISARM_Red_with_IDs.Cells(k, i).WrapText = True
            k = k + 1
        End If
    Next j
Next i

For i = 1 To 16
    strTaskID = Left(arrPhaseTask(2, i), 4)
    k = 3
    For j = 2 To oWS_techniques.UsedRange.Rows.Count
        If oWS_techniques.Cells(j, 4).Value = strTaskID Then
            oWS_DISARM_Red_no_IDs.Cells(k, i).Value = oWS_techniques.Cells(j, 2).Value
            oWS_DISARM_Red_no_IDs.Cells(k, i).Borders.LineStyle = xlContinuous
            oWS_DISARM_Red_no_IDs.Cells(k, i).WrapText = True
            k = k + 1
        End If
    Next j
Next i

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
