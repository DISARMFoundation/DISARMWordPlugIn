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
