VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchTechniques 
   Caption         =   "DISARM: Search Techniques then Insert Red Tag"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14625
   OleObjectBlob   =   "frmSearchTechniques.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchTechniques"
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

Private Sub chkDescriptions_Click()

'
' Store the value ofthe checkbox for looking in the technique descriptions
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "chkDescriptions_Click"

bLookInDescriptions = chkDescriptions.Value

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdSearchTechniques_Click()

Dim oFrm As frmSearchResultsTechniques

'
' When user clicks on the Search button store the search criteria in global variables and display the search results
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSearchTechniques_Click"

'
' Check that a search term has been input by user. If so, pass the search criteria to global variables
'

If IsNull(txtSearchTerm.Value) Or txtSearchTerm.Value = "" Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please supply a search term", vbOKCancel + vbInformation, "DISARM: Search Techniques then Insert Red Tag")
    GoTo PROC_EXIT
Else
    gstrSearchTerm = txtSearchTerm.Value
End If

If IsNull(lstPhases.Value) Then gstrPhaseName = "" Else gstrPhaseName = lstPhases.Value
If IsNull(lstTactics.Value) Then gstrTacticName = "" Else gstrTacticName = lstTactics.Value

'
' Close the search dialog and if there are results then display them
'

Unload Me 'frmSearchTechniques
Set oFrm = New frmSearchResultsTechniques
If bTechniquesFound Then oFrm.Show 'frmSearchResultsTechniques

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub lstPhases_Change()

'
' If the user chooses a specific phase then narrow the list of tactics to choose from to that phase
'

Dim strPlan As String, strPrepare As String, strExecute As String, strAssess As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "1stPhases_Change"

'Define the static variable tactics
strPlan = "Plan Strategy|Plan Objectives|Target Audience Analysis"
strPrepare = "Develop Narratives|Develop Content|Establish Social Assets|Establish Legitimacy|Microtarget|Select Channels and Affordances"
strExecute = "Conduct Pump Priming|Deliver Content|Maximise Exposure|Drive Online Harms|Drive Offline Activity|Persist in the Information Environment"
strAssess = "Assess Effectiveness"
lstTactics.Clear

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
' If the user chooses a specific tactic then take no special action. Wait for the user to specify
' search criteria and press the Search Techniques button
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "1stTactics_Change"

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub UserForm_Initialize()

'
' Populate the user form with a list of phases and tactics. The user can search within a specific phase
' and/or within a specific tactic, or search within all phases and/or tactics.
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

strPlan = "Plan Strategy|Plan Objectives|Target Audience Analysis|"
strPrepare = "Develop Narratives|Develop Content|Establish Social Assets|Establish Legitimacy|Microtarget|Select Channels and Affordances|"
strExecute = "Conduct Pump Priming|Deliver Content|Maximise Exposure|Drive Online Harms|Drive Offline Activity|Persist in the Information Environment|"
strAssess = "Assess Effectiveness"

lstTactics.Clear
lstTactics.List = Split(strPlan & strPrepare & strExecute & strAssess, "|")

If IsNull(bLookInDescriptions) Then bLookInDescriptions = False
chkDescriptions.Value = bLookInDescriptions

' following code positions dialog box in the same monitor screen as the word document
' see https://www.thespreadsheetguru.com/vba/launch-vba-userforms-in-correct-window-with-dual-monitors

Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdCancel_Click()

'
' Cancel out of the form
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdCancel_Click"
  
Unload Me

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
