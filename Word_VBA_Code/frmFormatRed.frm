VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormatRed 
   Caption         =   "DISARM: Format Red Tags"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "frmFormatRed.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormatRed"
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

'
' The Win64 conditional compiler constant is used to determine which version (32-bit or 64-bit) of Office
' is running. The 64-bit version of the code uses the LongLong and LongPtr data types and the PtrSafe keyword.
'
' For details see https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview
'

Private Sub UserForm_Initialize()

'
' Populate user form with initial values
'

Dim dblProfileRedColor As Double

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "UserForm_Initialize"

'
' If color has already been set in this session then set background color for sample red tag text to chosen color.
' Otherwise check if the background color for red tags has been set in the profile in the Tagging Workbook.
' If not set background to yellow.

If lngSetRedTagColor <> 0 Then
    txtSampleRedText.BackColor = lngSetRedTagColor
Else
    dblProfileRedColor = modMain.GetProfile("Red Color")
    If dblProfileRedColor = 0 Then
        'txtSampleRedText.BackColor = 16777215 ' If no highlighting color for red tags has been set in the profile default to yellow
        'lngSetRedTagColor = 16777215
        txtSampleRedText.BackColor = 65535 ' If no highlighting color for red tags has been set in the profile default to yellow
        lngSetRedTagColor = 65535
    Else
        txtSampleRedText.BackColor = dblProfileRedColor
        lngSetRedTagColor = dblProfileRedColor
    End If
End If

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

Private Sub cmdChooseColorRed_Click()

'
' Display dialog box with colors for user to choose for highlighting Red tag inline. Show sample text with background color
' Write chosen color to profile in Tagging Workbook
'

' experimenting with Win64 compiler constant see https://stackoverflow.com/questions/5506912/how-should-i-make-my-vba-code-compatible-with-64-bit-windows
'
#If Win64 Then
    Dim col As LongLong
#Else
    Dim col As Variant
#End If

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdChooseColorRed_Click"

'
' Set color picker to point to the current setting for Red Color tags
'

If lngSetRedTagColor <> 0 Then ' If color has already been set in this session
    #If Win64 Then
        col = CLngLng(lngSetRedTagColor)
    #Else
        col = lngSetRedTagColor
    #End If
Else
    dblProfileRedColor = modMain.GetProfile("Red Color") 'otherwise get the color from the profile for this document
    If dblProfileRedColor = 0 Then
        ' If no highlighting color for red tags has been set in the profile default to yellow
        #If Win64 Then
            col = CLngLng(65535)
        #Else
            col = 65535
        #End If
    Else
        #If Win64 Then
            col = CLngLng(dblProfileRedColor)
        #Else
            col = dblProfileRedColor
        #End If
    End If
End If

GetColor col
txtSampleRedText.BackColor = CLng(col)

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdSave_Click()

'
' Write color to tagging worksheet
'

If gcHandleProcErrors Then On Error GoTo PROC_ERR
PushCallStack "cmdSave_Click"
  
lngSetRedTagColor = txtSampleRedText.BackColor
modMain.WriteProfile "Red Color", lngSetRedTagColor

Unload Me

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Private Sub cmdCancel_Click()

'
' Cancel out of form
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
