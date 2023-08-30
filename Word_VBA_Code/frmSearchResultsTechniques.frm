VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchResultsTechniques 
   Caption         =   "DISARM: Search Results Techniques"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16440
   OleObjectBlob   =   "frmSearchResultsTechniques.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchResultsTechniques"
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
' Execute the search specified by the search term (and phase and/or tactic if specified) and populate
' this form with the results providing the user with a choice of techniques to select for tagging
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

'e.g. modMain.FillListTechniques "*Narrative*", "Prepare", "Develop Narratives", ListBoxTechniques
modMain.FillListTechniques gstrSearchTerm, gstrPhaseName, gstrTacticName, lstTechniques
    
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

Private Sub cmdSelectTechniques_Click()

'
' Add chosen technique(s) to sheet SummaryRedUnformatted and highlight techniques in sheet SummaryRedGraphic
'

Dim Tag As String
Dim i As Integer
Dim j As Integer
Dim strTacticID As String
Dim strTacticName As String
Dim strTechniqueID As String
Dim strTechniqueName As String
Dim strTechniqueSentence As String
Dim arrResult() As String

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSelectTechniques_Click"

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

arrResult = modFunctions.fcnSelectedItems(Controls("lstTechniques"))
If modFunctions.IsArrayAllocated(arrResult) Then

    For m_lngIndex = 0 To UBound(arrResult, 1)

        strTacticName = arrResult(m_lngIndex, 1)
        strTacticID = ReturnTacticID(strTacticName)
        strTechniqueName = arrResult(m_lngIndex, 2)
        If Right(strTechniqueName, 1) = " " Then
            strTechniqueName = Left(strTechniqueName, Len(strTechniqueName) - 1)
            strTechniqueID = ReturnTechniqueID(strTechniqueName & " ", strTacticID)
        Else
            strTechniqueID = ReturnTechniqueID(strTechniqueName, strTacticID)
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
    
    Next m_lngIndex
Else
    '
    ' If no techniques have been selected then prompt the user to select at least one
    '
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please select one or more techniques", vbOKCancel + vbInformation, "DISARM: Search Results")
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

Private Sub cmdPhase_Click()

'
' Sort techniques by phase
'
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "cmdPhase_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox lstTechniques, 1, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdTactic_Click()

'
' Sort techniques by tactic name
'
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "cmdTactic_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox lstTechniques, 2, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdTechnique_Click()

'
' Sort techniques by technique name
'
  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "cmdTechnique_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox lstTechniques, 3, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
