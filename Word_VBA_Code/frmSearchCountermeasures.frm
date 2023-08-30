VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchCountermeasures 
   Caption         =   "DISARM: Search Countermeasures then Insert Blue Tag"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "frmSearchCountermeasures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchCountermeasures"
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

bLookInDescriptionsCountermeasures = chkDescriptions.Value

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdSearchCountermeasures_Click()

Dim oFrm As frmSearchResultsCountermeasures

'
' When user clicks on the Search button store the search criteria in global variables and display the search results
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSearchCountermeasures_Click"

'
' Check that a search term has been input by user. If so, pass the search criteria to global variables
'

If IsNull(txtSearchTerm.Value) Or txtSearchTerm.Value = "" Then
    Dim intMsgReturn As Integer
    intMsgReturn = MsgBox("Please supply a search term", vbOKCancel + vbInformation, "DISARM: Search Countermeasures then Insert Blue Tag")
    GoTo PROC_EXIT
Else
    gstrSearchTermCountermeasures = txtSearchTerm.Value
End If

If IsNull(lstMetatechniques.Value) Then gstrMetatechniqueName = "" Else gstrMetatechniqueName = lstMetatechniques.Value

'
' Close the search dialog and if there are results then display them
'

Unload Me 'frmSearchCountermeasures
Set oFrm = New frmSearchResultsCountermeasures
If bCountermeasuresFound Then oFrm.Show 'frmSearchResultsCountermeasures

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub lstMetatechniques_Change()

'
' If the user chooses a specific metatechnique then take no special action. Wait for the user to specify
' search criteria and press the Search Techniques button
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "1stMetatechniques_Change"

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT

End Sub

Private Sub UserForm_Initialize()

'
' Populate the user form with a list of metatechniques. The user can search within a specific metatechnique
' or search within all metatechniques.
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

If IsNull(bLookInDescriptionsCountermeasures) Then bLookInDescriptionsCountermeasures = False
chkDescriptions.Value = bLookInDescriptionsCountermeasures

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

