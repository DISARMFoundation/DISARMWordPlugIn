VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCountermeasuresTopDown 
   Caption         =   "DISARM: Insert Tag from Blue Framework"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14430
   OleObjectBlob   =   "frmSelectCountermeasuresTopDown.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectCountermeasuresTopDown"
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

Private Sub UserForm_Initialize()

'
' Populate the user form with metatechniques
'

Dim arrMetatechniques(14) As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

arrMetatechniques(0) = "Resilience"
arrMetatechniques(1) = "Diversion"
arrMetatechniques(2) = "Daylight"
arrMetatechniques(3) = "Friction"
arrMetatechniques(4) = "Removal"
arrMetatechniques(5) = "Scoring"
arrMetatechniques(6) = "Metatechnique"
arrMetatechniques(7) = "Data Pollution"
arrMetatechniques(8) = "Dilution"
arrMetatechniques(9) = "Countermessaging"
arrMetatechniques(10) = "Verification"
arrMetatechniques(11) = "Cleaning"
arrMetatechniques(12) = "Targeting"
arrMetatechniques(13) = "Reduce Resources"

With lstMetatechniques
    .MultiSelect = fmMultiSelectSingle
    .List = arrMetatechniques
End With

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

Private Sub lstMetatechniques_Change()

'
' If the user has chosen a metatechnique then display all countermeasures for that metatechnique
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "lstMetatechniques_Change"

'
' First check that we are ready to tag i.e. all tagging worksheets created and open
'

CheckReadyToTag

If Not IsNull(lstMetatechniques.Value) Then modMain.LoadCountersFromExcel lstCountermeasures, lstMetatechniques.Value

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdSelectCountermeasures_Click()

'
' Add selected countermeasure(s) to sheet SummaryBlueUnformatted. Create and insert the tag with those
' countermeasures into te Word document.
'

Dim Tag As String
Dim i As Integer
Dim j As Integer
Dim strMetatechniqueID As String
Dim strMetatechniqueName As String
Dim strCountermeasureID As String
Dim strCountermeasureName As String
Dim strCountermeasureSentence As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSelectCountermeasures_Click"

'
' Determine Metatechnique ID and Name
'

strMetatechniqueName = lstMetatechniques.Value
strMetatechniqueID = ReturnMetatechniqueID(strMetatechniqueName)

'
' First check that we are ready to tag i.e. all tagging worksheets created and open
'

CheckReadyToTag

'
' Determine sentence in text that is being tagged
'

Dim lngCountermeasureSentenceIndex As Long
lngCountermeasureSentenceIndex = ReturnCountermeasureSentenceIndex()
strCountermeasureSentence = ReturnCountermeasureSentence(lngCountermeasureSentenceIndex)

'
' Create tag for the countermeasures(s) selected
'

Dim strCountermeasureTitle As String
Tag = " ("
j = 0
For i = 1 To lstCountermeasures.ListCount
    If lstCountermeasures.Selected(i - 1) = True Then
        strCountermeasureName = lstCountermeasures.List(i - 1)
        strCountermeasureID = ReturnCountermeasureID(strCountermeasureName, strMetatechniqueID)
        modMain.InsertRowSummaryBlueUnformatted strMetatechniqueID, strMetatechniqueName, strCountermeasureID, strCountermeasureName, strCountermeasureSentence, lngCountermeasureSentenceIndex
        ' Now create the inline tag
        j = j + 1
        If j > 1 Then
            Tag = Tag & ", "
        End If
        Tag = Tag & strCountermeasureName
        Tag = Tag & " [" & strMetatechniqueID & "." & strCountermeasureID & "]"
    End If
Next i
Tag = Tag & ")"

'
' If no countermeasures selected prompt user to choose at least one
'

If j = 0 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please select one or more countermeasures", vbOKCancel + vbInformation, "DISARM: Insert Blue Tag")
    GoTo PROC_EXIT
End If

'
' Save Tagging Workbook
'

modMain.SaveTaggingWorkbook

'
' Append tag inline to the Word document
'

modMain.WriteTag Tag, "Blue"

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
