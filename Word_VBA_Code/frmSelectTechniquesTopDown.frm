VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectTechniquesTopDown 
   Caption         =   "DISARM: Insert Tag from Red Framework"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   585
   ClientWidth     =   16770
   OleObjectBlob   =   "frmSelectTechniquesTopDown.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectTechniquesTopDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub UserForm_Initialize()

'
' Populate user form with a list of phases to choose from
'

Dim arrPhases(4) As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

arrPhases(0) = "Plan"
arrPhases(1) = "Prepare"
arrPhases(2) = "Execute"
arrPhases(3) = "Assess"

With lstPhases
    .MultiSelect = fmMultiSelectSingle
    .List = arrPhases
End With

' following code positions dialog box in the same monitor screen as the word document
' see https://www.thespreadsheetguru.com/vba/launch-vba-userforms-in-correct-window-with-dual-monitors

Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
Me.chkDetails = False
Me.Height = 245

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub lstPhases_Change()

'
' Provide user with Tactics to choose from based on Phase selected
'

Dim strPlan As String, strPrepare As String, strExecute As String, strAssess As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "1stPhases_Change"

'Define the static variable tactics
strPlan = "Plan Strategy|Plan Objectives|Target Audience Analysis"
strPrepare = "Develop Narratives|Develop Content|Establish Assets|Establish Legitimacy|Microtarget|Select Channels and Affordances"
strExecute = "Conduct Pump Priming|Deliver Content|Maximise Exposure|Drive Online Harms|Drive Offline Activity|Persist in the Information Environment"
strAssess = "Assess Effectiveness"
lstTactics.Clear

'Populate the Tactics listbox
Select Case lstPhases.Value
    Case "Plan": lstTactics.List = Split(strPlan, "|")
    Case "Prepare": lstTactics.List = Split(strPrepare, "|")
    Case "Execute": lstTactics.List = Split(strExecute, "|")
    Case "Assess": lstTactics.List = Split(strAssess, "|")
End Select

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT

End Sub

Private Sub lstTactics_Change()

'
' If user has chosen a tactic then display all the techniques and subtechniques for that tactic
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "1stTactics_Change"

'
' Initialize the list view to show the techniques and subtechniques
'

With Me.lstTechniques2
    .View = lvwReport
    .FullRowSelect = True
    .MultiSelect = True
    .LabelEdit = lvwManual
    .ColumnHeaders.Add , , "ID", 50
    .ColumnHeaders.Add , , "Technique Name", 300
    .ColumnHeaders.Add , , "", 1
    .ListItems.Clear
End With

'
' First check that we are ready to tag i.e. all tagging worksheets created and open
'

CheckReadyToTag

If Not IsNull(lstTactics.Value) Then modMain.LoadFromExcel lstTechniques2, lstTactics.Value

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub chkDetails_Click()

'
' If the user has checked the details checkbox then display the full user form
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "chkDetails_Click"

If chkDetails.Value = True Then
    Me.Height = 354
Else
    Me.Height = 245
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub lstTechniques2_ItemClick(ByVal Item As MSComctlLib.ListItem)

'
' If the user has clicked on a countermeasure then fill in the details. Set ethics rating and color.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "lstTechniques2_ItemClick"

Me.txtTechniqueID = Item.Text 'ID
Me.txtTechniqueName = Item.ListSubItems(1).Text 'name
Me.txtSummary = Item.ListSubItems(2).Text 'summary

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdSelectTechniques_Click()

'
' Add selected technique(s) to sheet SummaryRedUnformatted, highlight techniques in sheet SummaryRedGraphic,
' then create and insert the tag with these techniques into the Word document
'

Dim Tag As String
Dim i As Integer
Dim j As Integer
Dim strTacticID As String
Dim strTacticName As String
Dim strTechniqueID As String
Dim strTechniqueName As String
Dim strTechniqueSentence As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSelectTechniques_Click"

'
' Determine Tactic ID and Name
'

strTacticName = lstTactics.Value
strTacticID = ReturnTacticID(strTacticName)

'
' First check that we are ready to tag i.e. all tagging worksheets created and open
'

CheckReadyToTag

'
' Determine sentence in text that is being tagged
'

Dim lngTechniqueSentenceIndex As Long
lngTechniqueSentenceIndex = ReturnTechniqueSentenceIndex()
strTechniqueSentence = ReturnTechniqueSentence(lngTechniqueSentenceIndex)

'
' Create tag for the techniques(s) selected
'

Dim strTechniqueTitle As String
Dim strParentTechniqueID As String
Dim strParentTechniqueName As String
Dim varPos As Long
Tag = " ("
j = 0
For i = 1 To lstTechniques2.ListItems.Count
    If lstTechniques2.ListItems.Item(i).Selected = True Then
        strTechniqueID = lstTechniques2.ListItems(i).Text ' ID
        strTechniqueName = lstTechniques2.ListItems(i).ListSubItems(1).Text ' Name
        If Right(strTechniqueName, 1) = " " Then
            ' strip off any trailing space
            strTechniqueName = Left(strTechniqueName, Len(strTechniqueName) - 1)
        End If
        ' If this is a subtechnique then add the name of the parent technique to the tag and highlight both the subtechnique and parent technique in the graphic
        varPos = InStr(6, strTechniqueID, ".", vbTextCompare)
        If varPos = 0 Then
            strTechniqueTitle = strTechniqueName
            modMain.InsertRowSummaryRedUnformatted strTacticID, strTacticName, strTechniqueID, strTechniqueTitle, strTechniqueSentence, lngTechniqueSentenceIndex
            modMain.HighlightTechniqueSummaryRedGraphic strTacticID, strTacticName, strTechniqueID, strTechniqueName
        Else
            strParentTechniqueID = Left(strTechniqueID, varPos - 1)
            strParentTechniqueName = ReturnTechniqueName(strParentTechniqueID)
            strTechniqueTitle = strParentTechniqueName & ": " & strTechniqueName
            modMain.InsertRowSummaryRedUnformatted strTacticID, strTacticName, strTechniqueID, strTechniqueTitle, strTechniqueSentence, lngTechniqueSentenceIndex
            modMain.HighlightTechniqueSummaryRedGraphic strTacticID, strTacticName, strTechniqueID, strTechniqueName
            modMain.HighlightTechniqueSummaryRedGraphic strTacticID, strTacticName, strParentTechniqueID, strParentTechniqueName
        End If
        
        ' Now create the inline tag
        j = j + 1
        If j > 1 Then
            Tag = Tag & ", "
        End If
        Tag = Tag & strTechniqueTitle
        Tag = Tag & " [" & strTechniqueID & "]"
    End If
Next i
Tag = Tag & ")"

'
' If no techniques have been selected then prompt the user to select at least one
'

If j = 0 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please select one or more techniques", vbOKCancel + vbInformation, "DISARM: Insert Red Tag")
    GoTo PROC_EXIT
End If

'
' Save Tagging Workbook
'

modMain.SaveTaggingWorkbook

'
' Append tag inline to the Word document
'

modMain.WriteTag Tag, "Red"

'
' Hide the DISARM Tagger dialog box
'

Me.Hide

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

