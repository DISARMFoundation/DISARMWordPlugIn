VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchResultsCountermeasures 
   Caption         =   "DISARM: Search Results Countermeasures"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   OleObjectBlob   =   "frmSearchResultsCountermeasures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchResultsCountermeasures"
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
' Execute the search specified by the search term (and phase and/or Metatechnique if specified) and populate
' this form with the results providing the user with a choice of Countermeasures to select for tagging
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

'e.g. modMain.FillListCountermeasures "*Account*", "Friction", ListBoxCountermeasures
modMain.FillListCountermeasures gstrSearchTermCountermeasures, gstrMetatechniqueName, lstCountermeasures
    
Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

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

Private Sub cmdSelectCountermeasures_Click()

'
' Add chosen Countermeasure(s) to sheet SummaryBlueUnformatted.
' In the future we may highlight Countermeasures in a sheet called SummaryBlueGraphic
'

Dim Tag As String
Dim i As Integer
Dim j As Integer
Dim strMetatechniqueID As String
Dim strMetatechniqueName As String
Dim strCountermeasureID As String
Dim strCountermeasureName As String
Dim strCountermeasureSentence As String
Dim arrResult() As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSelectCountermeasures_Click"

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
' Create tag for the Countermeasures(s) selected
'

Dim strCountermeasureTitle As String
Tag = " ("
j = 0

arrResult = modFunctions.fcnSelectedItems(Controls("lstCountermeasures"))
If modFunctions.IsArrayAllocated(arrResult) Then

    For m_lngIndex = 0 To UBound(arrResult, 1)

        strMetatechniqueName = arrResult(m_lngIndex, 0)
        strMetatechniqueID = ReturnMetatechniqueID(strMetatechniqueName)
        strCountermeasureName = arrResult(m_lngIndex, 1)
        If Right(strCountermeasureName, 1) = " " Then
            strCountermeasureName = Left(strCountermeasureName, Len(strCountermeasureName) - 1)
            strCountermeasureID = ReturnCountermeasureID(strCountermeasureName & " ", strMetatechniqueID)
        Else
            strCountermeasureID = ReturnCountermeasureID(strCountermeasureName, strMetatechniqueID)
        End If
        
        strCountermeasureTitle = strCountermeasureName
        modMain.InsertRowSummaryBlueUnformatted strMetatechniqueID, strMetatechniqueName, strCountermeasureID, strCountermeasureTitle, strCountermeasureSentence, lngCountermeasureSentenceIndex
        ' Anticipating building a graphic for blue
        'modMain.HighlightCountermeasureSummaryBlueGraphic strMetatechniqueID, strMetatechniqueName, strCountermeasureID, strCountermeasureName
        
        ' Now create the inline tag
        j = j + 1
        If j > 1 Then
            Tag = Tag & ", "
        End If
        Tag = Tag & strCountermeasureTitle
        Tag = Tag & " [" & strCountermeasureID & "]"
    
    Next m_lngIndex
Else
    '
    ' If no Countermeasures have been selected then prompt the user to select at least one
    '
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please select one or more Countermeasures", vbOKCancel + vbInformation, "DISARM: Search Results Countermeasures")
    GoTo PROC_EXIT
End If

Tag = Tag & ")"

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

Private Sub cmdMetatechnique_Click()

'
' Sort Countermeasures by Metatechnique name
'
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "cmdMetatechnique_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox lstCountermeasures, 1, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdCountermeasure_Click()

'
' Sort Countermeasures by Countermeasure name
'
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "cmdCountermeasure_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox lstCountermeasures, 2, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

