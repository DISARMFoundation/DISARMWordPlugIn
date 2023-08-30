Attribute VB_Name = "modColor"
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
Option Base 0

'
' Code here is from https://stackoverflow.com/questions/36000721/ms-word-vba-i-need-a-color-palette-dialog-box
'
' The DISARM Word Plug-In has been tested on a 64-bit Windows 10 operating system on both
' the 32-bit version of Office and the 64-bit version.
'
' The Win64 conditional compiler constant is used to determine which version (32-bit or 64-bit) of Office
' is running. The 64-bit version of the code uses the LongLong and LongPtr data types and the PtrSafe keyword.
' The Win64 conditional compiler constant is also needed in the choose color routines associated with
' the forms frmFormatBlue, frmFormatRed, and frmFormatRedGraphic.
'
' For details see https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview
'

#If Win64 Then
Private Type CHOOSECOLOR
  lStructSize As LongLong
  hwndOwner As LongPtr
  hInstance As LongPtr
  rgbResult As LongLong
  lpCustColors As LongPtr
  flags As LongLong
  lCustData As LongLong
  lpfnHook As LongLong
  lpTemplateName As String
End Type

Private Declare PtrSafe Function MyChooseColor _
    Lib "comdlg32.dll" Alias "ChooseColorW" _
    (ByRef pChoosecolor As CHOOSECOLOR) As Boolean

Public Function GetColor(ByRef col As LongLong) As _
    Boolean

'
' Function to launch the Color Picker dialog and set the color "col".
' Returns True if user chooses a color or False if user cancels.
'
'

  Static CS As CHOOSECOLOR
  Static CustColor(15) As LongLong

  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "GetColor"

  CS.lStructSize = Len(CS)
  CS.hwndOwner = 0
  CS.flags = &H1 Or &H2
  CS.lpCustColors = VarPtr(CustColor(0))
  CS.rgbResult = col
  CS.hInstance = 0
  GetColor = MyChooseColor(CS)
  If GetColor = False Then GoTo FUNC_EXIT

  GetColor = True
  col = CS.rgbResult
  
FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

#Else

Private Type CHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function MyChooseColor _
    Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpcc As CHOOSECOLOR) As Long
    
Public Function GetColor(ByRef col As Variant) As _
    Long

'
' Function to launch the Color Picker dialog and set the color "col".
' Returns 1 if user chooses a color or 0 if user cancels.
'
'

  Static CS As CHOOSECOLOR
  Static CustColor(15) As Long

  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "GetColor"

  CS.lStructSize = Len(CS)
  CS.hwndOwner = 0
  CS.flags = &H1 Or &H2
  CS.lpCustColors = VarPtr(CustColor(0))
  CS.rgbResult = col
  CS.hInstance = 0
  GetColor = MyChooseColor(CS)
  If GetColor = 0 Then GoTo FUNC_EXIT

  GetColor = 1
  col = CS.rgbResult
  
FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

#End If

Sub MyZoomIn()
    ' https://wordribbon.tips.net/T009831_Zooming_with_the_Keyboard.html
    Dim ZP As Integer
    ZP = Int(ActiveWindow.ActivePane.View.Zoom.Percentage * 1.5)
    If ZP > 500 Then ZP = 500
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZP
End Sub

Sub MyZoomOut()
    ' https://wordribbon.tips.net/T009831_Zooming_with_the_Keyboard.html
    Dim ZP As Integer
    ZP = Int(ActiveWindow.ActivePane.View.Zoom.Percentage / 1.5)
    If ZP < 10 Then ZP = 10
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZP
End Sub
