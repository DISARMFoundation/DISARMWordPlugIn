VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchResultsTechniques 
   Caption         =   "DISARM: Search Results Techniques"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13290
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

'
' Initialize the list view for the results of the search
'

With Me.lstTechniques2
    .View = lvwReport
    .FullRowSelect = True
    .MultiSelect = True
    .LabelEdit = lvwManual
    .ColumnHeaders.Add , , "Phase", 50 'phase name
    .ColumnHeaders.Add , , "Tactic", 200 'tactic name
    .ColumnHeaders.Add , , "ID", 50 'technique ID
    .ColumnHeaders.Add , , "Technique", 321 'technique name
    .ColumnHeaders.Add , , "", 1 'technique summary - not shown in the list but available for the text box
    .ListItems.Clear
End With

'e.g. modMain.FillListTechniques "*Narrative*", "Prepare", "Develop Narratives", ListBoxTechniques
modMain.FillListTechniques2 gstrSearchTerm, gstrPhaseName, gstrTacticName, lstTechniques2
    
' Following code positions dialog box in the same monitor screen as the word document
' see https://www.thespreadsheetguru.com/vba/launch-vba-userforms-in-correct-window-with-dual-monitors
' By default do not show the complete form with all the details

Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
Me.chkDetails = False
Me.Height = 320

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

Private Sub chkDetails_Click()

'
' If the user has checked the details checkbox then display the full user form
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "chkDetails_Click"

If chkDetails.Value = True Then
    Me.Height = 433
Else
    Me.Height = 320
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
' If the user has clicked on a technique then fill in the text boxes below.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "lstTechniques2_ItemClick"

Me.txtTechniqueID = Item.ListSubItems(2).Text 'technique ID
Me.txtTechniqueName = Item.ListSubItems(3).Text 'technique name
Me.txtSummary = Item.ListSubItems(4).Text 'technique summary

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

For i = 1 To lstTechniques2.ListItems.Count
    If lstTechniques2.ListItems.Item(i).Selected = True Then
        strTechniqueID = lstTechniques2.ListItems(i).ListSubItems(2).Text ' technique ID
        strTechniqueName = lstTechniques2.ListItems(i).ListSubItems(3).Text ' Name
        strTacticName = lstTechniques2.ListItems(i).ListSubItems(1).Text ' tactic name
        strTacticID = ReturnTacticID(strTacticName)
        If Right(strTechniqueName, 1) = " " Then ' strip trailing spaces
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
' If no techniques selected prompt user to choose at least one
'

If j = 0 Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please select one or more techniques", vbOKCancel + vbInformation, "DISARM: Search Results")
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
