Attribute VB_Name = "DISARM"
' This is a macro-enabled global template containing the XML and VBA for the DISARM Word Plug-In.

' Copyright (C) 2023 DISARM Foundation

' This program is free software: you can redistribute it and/or modify it under the terms of the GNU 'General Public License
' as published by the Free Software Foundation, either version 3 of the License, or 'any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License at https://www.gnu.org/licenses/ for more details.

' To report bugs or suggestions for improvement please send email to info@disarm.foundation.
'___________________________________________________________________________________________________________________________

'
' This module houses routines which are called from the Microsoft Word ribbon via procedure RibbonControl.MyBtnMacro
'

Option Explicit

Sub FormatBlue()

'
' Define desired format for Blue tag
'
  Dim oFrm As frmFormatBlue
  
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "FormatBlue"
  
  Set oFrm = New frmFormatBlue
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FormatRed()

'
' Define desired format for Red tag
'
  Dim oFrm As frmFormatRed
  
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "FormatRed"
  
  Set oFrm = New frmFormatRed
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub FormatRedGraphic()
  
'
' Define desired format for selected cells in the Summary Red Graphic
'
  Dim oFrm As frmFormatRedGraphic
  
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "FormatRedGraphic"
  
  Set oFrm = New frmFormatRedGraphic
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertRedTag()

'
' Insert a Red tag top-down i.e. by selecting Phase then Tactic then Technique(s)
'
  Dim oFrm As frmSelectTechniquesTopDown

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "InsertRedTag"
  
  Set oFrm = New frmSelectTechniquesTopDown
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertBlueTag()

'
' Insert a Blue tag top-down i.e. by selecting Metatechnique then Countermeasure(s)
'
  Dim oFrm As frmSelectCountermeasuresTopDown
  
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "InsertBlueTag"
  
  Set oFrm = New frmSelectCountermeasuresTopDown
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing
  
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub SearchTechniques()
  
'
' Search for techniques to include in Red tag
'
  Dim oFrm As frmSearchTechniques

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "SearchTechniques"
  
  Set oFrm = New frmSearchTechniques
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub SearchCountermeasures()
  
'
' Search for countermeasures to include in Blue tag
'
  Dim oFrm As frmSearchCountermeasures

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "SearchCountermeasures"
  
  Set oFrm = New frmSearchCountermeasures
  Load oFrm
  oFrm.Show
  Unload oFrm
  Set oFrm = Nothing

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub CallListTags()

'
' Display the list of red tags from which the user can choose to be marked for deletion
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CallListTags"

'
' Allows you to cancel out of an infinite loop using CTRL-C
'
Application.EnableCancelKey = wdCancelInterrupt

If ReturnNumRowsSummaryRed <= 1 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("You have not tagged any techniques for this document", vbOKCancel + vbInformation, "DISARM: Insert Summary Red Table")
    GoTo PROC_EXIT
End If

Dim oFrm As frmListTags
   Set oFrm = New frmListTags
   oFrm.Show
   Unload oFrm
   Set oFrm = Nothing
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryRedGraphic()

'
' Insert a suummary graphic into the Word document
' highlighting all the techniques and subtechniques that have been tagged
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryRedGraphic"

'
' Place the summary red graphic into the document
'

If ReturnNumRowsSummaryRed <= 1 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("You have not tagged any techniques for this document", vbOKCancel)
    GoTo PROC_EXIT
End If

modMain.FormatRedGraphic
modMain.CopyRedGraphic
Selection.InsertBreak Type:=wdPageBreak
If Selection.PageSetup.Orientation = wdOrientLandscape Then
    Selection.PageSetup.Orientation = wdOrientPortrait
End If

Selection.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
        Placement:=wdInLine, DisplayAsIcon:=False

modMain.AddSpace

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub SpecifyLocationForNavigatorFiles()

'
' Allow user to specify or change the location for storing JSON files for the DISARM Navigator
'

Dim JSONDirectory As String
Dim fso3 As New FileSystemObject
Dim oFD As FileDialog
Dim intMsgReturn As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "SpecifyLocationForNavigatorFiles"

'
' Look up the worksheet "User Profile" to determine the current location
'

JSONDirectory = ReturnUserProfile("JSON_Directory")
If JSONDirectory = "" Or Not fso3.FolderExists(JSONDirectory) Then 'No location specified or location does not exist
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    oFD.Title = "Choose a Location to Save JSON files for the DISARM Navigator"
    oFD.ButtonName = "Choose"
    oFD.InitialFileName = Environ("USERPROFILE") & "\"     'sets the folder e.g. C:\Users\steph\
    oFD.InitialView = msoFileDialogViewLargeIcons
    
    With oFD
        If .Show = -1 Then 'if OK is pressed
            Call SetUserProfile("JSON_Directory", .SelectedItems(1))
            intMsgReturn = MsgBox("Location for DISARM Navigator layer files is " & .SelectedItems(1) & ".", _
                vbOKCancel + vbInformation, "DISARM: Specify Location for Navigator Files")
        Else
            GoTo PROC_EXIT ' If user cancels then we do not know where to save the JSON file so exit
        End If
    End With
Else
    intMsgReturn = MsgBox("Files for the DISARM Navigator are currently saved to " & JSONDirectory & ". Would you like to choose a different location?", _
               vbYesNoCancel, "DISARM: Specify Location for Navigator Files")
    If intMsgReturn = vbYes Then
        Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
        oFD.Title = "Choose a Location to Save JSON files for the DISARM Navigator"
        oFD.ButtonName = "Choose"
        oFD.InitialFileName = Environ("USERPROFILE") & "\"     'sets the folder e.g. C:\Users\steph\
        oFD.InitialView = msoFileDialogViewLargeIcons
        
        With oFD
            If .Show = -1 Then 'if OK is pressed
                Call SetUserProfile("JSON_Directory", .SelectedItems(1))
                intMsgReturn = MsgBox("Location for DISARM Navigator layer files is now " & .SelectedItems(1) & ".", _
                    vbOKCancel + vbInformation, "DISARM: Specify Location for Navigator Files")
            Else
                GoTo PROC_EXIT ' If user cancels then we do not know where to save the JSON file so exit
            End If
        End With
    Else
        GoTo PROC_EXIT
    End If
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub CreateNavigatorFile()

'
' Create a layer file in JSON format for the DISARM Navigator from the techniques tagged by the user for this document
'

Dim JSON As Object
Dim JsonVBA As String
Dim Part1 As String
Dim Part2 As String
Dim Part3 As String
Dim fso As New FileSystemObject
Dim fso2 As New FileSystemObject
Dim fso3 As New FileSystemObject
Dim DocumentName As String
Dim strFolderPath As String
Dim tsout As TextStream
Dim JSONDirectory As String
Dim sFolder As String
Dim oFD As FileDialog
Dim intMsgReturn As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CreateNavigatorFile"

'
' Look up the worksheet "User Profile" to determine the location where the JSON layer file should be saved
'

JSONDirectory = ReturnUserProfile("JSON_Directory")
If JSONDirectory = "" Or Not fso3.FolderExists(JSONDirectory) Then 'No location specified or location does not exist
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    oFD.Title = "Choose a Location to Save JSON files for the DISARM Navigator"
    oFD.ButtonName = "Choose"
    oFD.InitialFileName = Environ("USERPROFILE") & "\"     'sets the folder e.g. C:\Users\steph\
    oFD.InitialView = msoFileDialogViewLargeIcons
    
    With oFD
        If .Show = -1 Then 'if OK is pressed
            strFolderPath = .SelectedItems(1)
            Call SetUserProfile("JSON_Directory", .SelectedItems(1))
        Else
            GoTo PROC_EXIT ' If user cancels then we do not know where to save the JSON file so exit
        End If
    End With
Else
    strFolderPath = JSONDirectory ' location specified and exists
End If

DocumentName = fso2.GetBaseName(Application.ActiveDocument.Name) ' Use the same name for the .json file as the Word document

' Pre-amble contains name of the layer plus some parameters for the ATT&CK Navigator
Part1 = "{""name"":""" & DocumentName & """,""versions"":{""attack"":""1"",""navigator"":""4.8.2"",""layer"":""4.4""}," & _
"""domain"":""DISARM"",""description"":"""",""filters"":{""platforms"":[""Windows"",""Linux"",""Mac""]}" & _
"""sorting"":0,""layout"":{""layout"":""flat"",""aggregateFunction"":""average"",""showID"":true,""showName"":true," & _
"""showAggregateScores"":false,""countUnscored"":false},""hideDisabled"":false,""techniques"":["

' The next part contains the JSON for all the active techniques tagged by the user
Part2 = ReturnJSONforTaggedTechniques

If Part2 = "" Then GoTo PROC_EXIT

' The postscript contains some formatting parameters for the ATT&CK Navigator
Part3 = "],""gradient"":{""colors"":[""#ff6666ff"",""#ffe766ff"",""#8ec843ff""],""minValue"":0,""maxValue"":100},""legendItems"":[],""metadata"":[]," & _
"""links"":[],""showTacticRowBackground"":false,""tacticRowBackground"":""#dddddd"",""selectTechniquesAcrossTactics"":true,""selectSubtechniquesWithParent"":false}"

' Now combine all three parts
JsonVBA = Part1 & Part2 & Part3

' Create a JSON Object from the JsonVBA string, convert this to pretty print JSON, then write to file
Set JSON = JsonConverter.ParseJson(JsonVBA)
Set tsout = fso.CreateTextFile(strFolderPath & "\" & DocumentName & ".json")
Call tsout.WriteLine(JsonConverter.ConvertToJson(JSON, Whitespace:=2))

' Inform user that file created successfully
intMsgReturn = MsgBox("DISARM Navigator layer file " & strFolderPath & "\" & DocumentName & ".json created successfully", _
               vbOKCancel + vbInformation, "DISARM: Create Navigator File")

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryRedTable()

'
' Insert a table into the Word document listing all the techniques that have been tagged
' and the sentences in which they were tagged.
'

Dim strTaskID As Variant
Dim strTaskName As String
Dim i As Integer
Dim j As Integer
Dim counter As Integer
Dim dblCountOfTagsForTask As Double
Dim dblTaskIDIndex As Double

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryRedTable"
  
'
' Allows you to cancel out of an infinite loop using CTRL-C
'
Application.EnableCancelKey = wdCancelInterrupt

If ReturnNumRowsSummaryRed <= 1 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("You have not tagged any techniques for this document", vbOKCancel + vbInformation, "DISARM: Insert Summary Red Table")
    GoTo PROC_EXIT
End If

modMain.CreateSummaryRedTable

modMain.SortSummaryRedUnformatted

Dim arrPhaseTask As Variant
arrPhaseTask = ReturnPhaseTaskArray()
Dim arrTaskRow() As Integer
counter = 0
intTableRowNumber = 1
For i = 1 To 16
    strTaskID = Left(arrPhaseTask(2, i), 4)
    strTaskName = Mid(arrPhaseTask(2, i), 7, Len(arrPhaseTask(2, i)) - 6)
    dblTaskIDIndex = TaskIDIndex(strTaskID)
    If dblTaskIDIndex > 0 Then
        counter = counter + 1
        modMain.InsertSummaryRedTableTaskHeader strTaskName & " [" & strTaskID & "]"
        ReDim Preserve arrTaskRow(counter - 1)
        arrTaskRow(counter - 1) = intTableRowNumber
        modMain.InsertSummaryRedTableTechniques strTaskID, dblTaskIDIndex
    End If
Next i

'
' Loop through the table and format the task headers (tried to do this above by merging cells but this caused alignment problems)
'

For j = 1 To counter
    modMain.FormatTaskRow arrTaskRow(j - 1)
Next j

modMain.AddSpace

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub InsertSummaryBlueTable()

'
' Insert a table into the Word document listing all the countermeasures that have been tagged
' and the sentences in which they were tagged.
'

Dim strMetatechniqueID As Variant
Dim strMetatechniqueName As String
Dim i As Integer
Dim j As Integer
Dim counter As Integer
Dim dblCountOfTagsForMetatechnique As Double
Dim dblMetatechniqueIDIndex As Double

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "InsertSummaryBlueTable"
  
'
' Allows you to cancel out of an infinite loop using CTRL-C
'
Application.EnableCancelKey = wdCancelInterrupt

If ReturnNumRowsSummaryBlue <= 1 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("You have not tagged any countermeasures for this document", vbOKCancel + vbInformation, "DISARM: Insert Summary Blue Table")
    GoTo PROC_EXIT
End If

modMain.CreateSummaryBlueTable

modMain.SortSummaryBlueUnformatted

Dim arrMetatechnique As Variant
arrMetatechnique = ReturnMetatechniqueArray()
Dim arrMetatechniqueRow() As Integer ' used to record the row numbers of those rows with headers for metatechniques
counter = 0
intBlueTableRowNumber = 1
For i = 1 To UBound(arrMetatechnique, 2)
    strMetatechniqueID = arrMetatechnique(1, i)
    strMetatechniqueName = arrMetatechnique(2, i)
    dblMetatechniqueIDIndex = MetatechniqueIDIndex(strMetatechniqueID)
    
    If dblMetatechniqueIDIndex > 0 Then
        counter = counter + 1
        modMain.InsertSummaryBlueTableMetatechniqueHeader strMetatechniqueName & " [" & strMetatechniqueID & "]"
        ReDim Preserve arrMetatechniqueRow(counter - 1)
        arrMetatechniqueRow(counter - 1) = intBlueTableRowNumber
        modMain.InsertSummaryBlueTableCountermeasures strMetatechniqueID, dblMetatechniqueIDIndex
    End If
Next i

'
' Loop through the table and format the Metatechnique headers
'

For j = 1 To counter
    modMain.FormatMetatechniqueRow arrMetatechniqueRow(j - 1)
Next j

modMain.AddSpace

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub LinkToExplorer()

'
' Bring up the DISARM Explorer in the browser
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LinkToExplorer"
  
CreateObject("WScript.Shell").Run ("https://disarmframework.herokuapp.com/")

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Sub LinkToNavigator()

'
' Bring up the DISARM Navigator in the browser
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LinkToNavigator"
  
CreateObject("WScript.Shell").Run ("https://disarmfoundation.github.io/disarm-navigator/")

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub LinkToGitHub()

'
' Bring up the DISARM GitHub pages in the browser
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "LinkToGitHub"
  
CreateObject("WScript.Shell").Run ("https://github.com/DISARMFoundation")

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub ClearHistory()

'
' This routine will clear the tagging history by resetting the unique DISARM Name for the document thereby triggering a new set of tagging worksheets
'

Dim intMsgReturn As Integer

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "ClearHistory"
  
intMsgReturn = MsgBox("Are you sure you want to clear your history? This will delete all the entries in the summary tables and graphic. " _
    & "Press OK to delete. Press Cancel to retain", vbOKCancel + vbQuestion, "DISARM: Clear History")

If intMsgReturn = vbOK Then
    modMain.Reset_DISARM_Name
    intMsgReturn = MsgBox("History cleared successfully", vbOKCancel + vbInformation, "DISARM: Clear History")
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
