VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListTags 
   Caption         =   "DISARM: Tag History"
   ClientHeight    =   8790.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23880
   OleObjectBlob   =   "frmListTags.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListTags"
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
' Initialize the userform frmListTags with a list of all tagged techniques that have not been deleted
' i.e. all techniques listed in the SummaryRedUnformatted worksheet with a status of "Active"
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

   With ListBoxTags
    
    'Load ListBoxtags with a list of all active tagged techniques
  
    modMain.FillListTags ListBoxTags
        
  End With

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
Private Sub CommandButtonDelete_Click()

'
' Mark the selected techniques as Deleted in the tagging worksheet then redisplay the list of techniques tagged
'

Dim arrResult() As String
Dim strSQL As String
Dim strParentTechniqueID As String
Dim varPos As Long

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CommandButtonDelete_Click"
  
  arrResult = modFunctions.fcnSelectedItems(Controls("ListBoxTags"))
  If modFunctions.IsArrayAllocated(arrResult) Then
       For m_lngIndex = 0 To UBound(arrResult, 1)
          ' arrResult(m_lngIndex, 2) is the SentenceIndex; arrResult(m_lngIndex, 0) is the TechniqueID
          modMain.MarkTagAsDeleted arrResult(m_lngIndex, 2), arrResult(m_lngIndex, 0)
          ' Now update the summary graphic if there are no active techniques left for this TechniqueID
          If modMain.NoActiveTechniquesLeft(arrResult(m_lngIndex, 0)) Then
            modMain.UnHighlightTechniqueSummaryRedGraphic (arrResult(m_lngIndex, 0))
            ' If this was a subtechnique then update the parent technique in the summary graphic if there are no
            ' active techniques for the technique id of the parent
            varPos = InStr(6, arrResult(m_lngIndex, 0), ".", vbTextCompare)
            If varPos <> 0 Then
                strParentTechniqueID = Left(arrResult(m_lngIndex, 0), varPos - 1)
                If modMain.NoActiveTechniquesLeft(strParentTechniqueID) Then
                    modMain.UnHighlightTechniqueSummaryRedGraphic (strParentTechniqueID)
                End If
            End If
          End If
       Next m_lngIndex
       
       '
       ' Save Tagging Workbook
       '
       
       modMain.SaveTaggingWorkbook
       
       ' Now redisplay the user form to show the updated list of tags
       UserForm_Initialize
  Else
    MsgBox "Nothing selected."
  End If

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub CommandButtonTechniqueID_Click()

'
' Sort tags by technique ID
'

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "CommandButtonTechniqueID_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox ListBoxTags, 1, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub CommandButtonTechnique_Click()

'
' Sort tags by technique name
'

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "CommandButtonTechnique_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox ListBoxTags, 2, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub CommandButtonSentenceID_Click()

'
' Sort tags by sentence ID
'

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "CommandButtonSentenceID_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox ListBoxTags, 3, False, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub CommandButtonSentence_Click()

'
' Sort tags by sentence
'

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "CommandButtonSentence_Click"
  
  'Pass listbox, sort column number, alphabeticall(True)/numerically(False), Ascending(True)/Descending(False)
  
  modFunctions.SortListBox ListBoxTags, 4, True, True
   
PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub CommandButtonCancel_Click()

'
' Cancel out of form
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "CommandButtonCancel_Click"
  
Unload Me

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
