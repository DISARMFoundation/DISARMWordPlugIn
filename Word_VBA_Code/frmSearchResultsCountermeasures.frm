VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchResultsCountermeasures 
   Caption         =   "DISARM: Search Results Countermeasures"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13785
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

'
' Load images for warning triangles into image list.
' See VBA A2Z, "Working with ListView Control in Excel VBA", at https://www.youtube.com/watch?v=U1sQ1-Oa0fs
' In the YouTube video they use the LoadPicture function to load images from the local file system.
' This approach certainly worked but would be problematic for packaging. So instead I created three
' separate controls of type image and placed these on the form but with visible set to false.
' I then load the icons using the Picture property.
'
' Note I also explored the use of the PastePicture function available at
' https://stackoverflow.com/questions/25534970/how-do-you-populate-imagelist-with-shape-from-worksheet
' but this was written for a 32-bit Office and it seemed a pain to modify the code to run on 32- or 64-bit.
'

With imlWarningTriangles.ListImages
     '.Add , , LoadPicture(ThisDocument.Path & "\st_green.bmp")
     '.Add , , LoadPicture(ThisDocument.Path & "\st_orange.bmp")
     '.Add , , LoadPicture(ThisDocument.Path & "\st_red.bmp")
     .Add , , imgGreenTriangle.Picture
     .Add , , imgOrangeTriangle.Picture
     .Add , , imgRedTriangle.Picture
End With

'
' Initialize the list view. The width of the form shows only the metatechnique, countermeasure name
' and ethics triangle
'

With Me.lstCountermeasures2
    .View = lvwReport
    .FullRowSelect = True
    .MultiSelect = True
    .LabelEdit = lvwManual
    .ColumnHeaders.Add , , "Metatechnique", 120
    .ColumnHeaders.Add , , "Countermeasure", 480
    .ColumnHeaders.Add , , "Ethics", 30 'triangle
    .ColumnHeaders.Add , , "", 1 'color
    .ColumnHeaders.Add , , "", 1 'ethics
    .ColumnHeaders.Add , , "", 1 'summary
    .ListItems.Clear
    .SmallIcons = imlWarningTriangles
End With

'
' Populate the list view with the results of the search. Keeping the old list box code just in case.
'

'e.g. modMain.FillListCountermeasures "*Account*", "Friction", ListBoxCountermeasures
'modMain.FillListCountermeasures gstrSearchTermCountermeasures, gstrMetatechniqueName, lstCountermeasures
modMain.FillListCountermeasures2 gstrSearchTermCountermeasures, gstrMetatechniqueName, lstCountermeasures2
    
' following code positions dialog box in the same monitor screen as the word document
' see https://www.thespreadsheetguru.com/vba/launch-vba-userforms-in-correct-window-with-dual-monitors
' by default do not show the complete form with all the details

Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
Me.chkDetails = False
Me.Height = 315

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
    Me.Height = 537
Else
    Me.Height = 315
End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
Private Sub lstCountermeasures2_ItemClick(ByVal Item As MSComctlLib.ListItem)

'
' If the user has clicked on a countermeasure then fill in the details. Set ethics rating and color.
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "lstCountermeasures2_ItemClick"

Me.txtCounter = Item.ListSubItems(1).Text 'name
If Item.ListSubItems(3).Text = "g" Then
    Me.txtEthicsRating.BackColor = vbGreen
    Me.txtEthicsRating = "largely unproblematic"
ElseIf Item.ListSubItems(3).Text = "o" Then
    Me.txtEthicsRating.BackColor = RGB(255, 165, 0)
    Me.txtEthicsRating = "potentially problematic"
ElseIf Item.ListSubItems(3).Text = "r" Then
    Me.txtEthicsRating.BackColor = vbRed
    Me.txtEthicsRating = "highly problematic"
Else
    Me.txtEthicsRating.BackColor = vbNone
    Me.txtEthicsRating = ""
End If
Me.txtGuidance = Item.ListSubItems(4).Text 'ethics
Me.txtSummary = Item.ListSubItems(5).Text 'summary

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
' countermeasures into the Word document.
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
For i = 1 To lstCountermeasures2.ListItems.Count
    If lstCountermeasures2.ListItems(i).Selected = True Then
        strMetatechniqueName = lstCountermeasures2.ListItems(i).Text ' name of metatechnique
        strMetatechniqueID = ReturnMetatechniqueID(strMetatechniqueName) ' get ID for metatechnique
        strCountermeasureName = lstCountermeasures2.ListItems(i).ListSubItems(1).Text ' name of counter
        strCountermeasureID = ReturnCountermeasureID(strCountermeasureName, strMetatechniqueID) ' get ID of counter
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
