Attribute VB_Name = "modErrorsAndTracing"
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
' Routines here thanks to FMS at https://www.fmsinc.com/tpapers/vbacode/Debug.asp
'

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer

' Array of procedure names in the call stack
Private mastrCallStack() As String

' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10

' Global constant to control error handling for Procedures
Public Const gcHandleProcErrors As Boolean = True

' Global constant to control error handling for Functions
Public Const gcHandleFuncErrors As Boolean = True

' Global constant to control tracing
Public Const gcTracing As Boolean = False

Sub PushCallStack(strProcName As String)
  ' Comments: Add the current procedure name to the Call Stack.
  '           Should be called whenever a procedure is called

  On Error Resume Next

  ' Verify the stack array can handle the current array element
  If mintStackPointer > UBound(mastrCallStack) Then
    ' If array has not been defined, initialize the error handler
    If Err.Number = 9 Then
      ErrorHandlerInit
    Else
      ' Increase the size of the array to not go out of bounds
      ReDim Preserve mastrCallStack(UBound(mastrCallStack) + mcintIncrementStackSize)
    End If
  End If

  On Error GoTo 0

  mastrCallStack(mintStackPointer) = strProcName

  ' If tracing enabled then print procedure name
  If gcTracing Then Debug.Print strProcName
  
  ' Increment pointer to next element
  mintStackPointer = mintStackPointer + 1
End Sub

Private Sub ErrorHandlerInit()
  mintStackPointer = 1
  ReDim mastrCallStack(1 To mcintIncrementStackSize)
End Sub

Sub PopCallStack()
  ' Comments: Remove a procedure name from the call stack

  If mintStackPointer <= UBound(mastrCallStack) Then
    
    ' If tracing enabled then print procedure name
    'If gcTracing Then Debug.Print "Exiting ", mastrCallStack(mintStackPointer - 1)
    
    mastrCallStack(mintStackPointer) = ""
 
  End If

  ' Reset pointer to previous element
  mintStackPointer = mintStackPointer - 1
End Sub

Sub GlobalErrHandler()
  ' Comments: Main procedure to handle errors that occur.

  Dim strError As String
  Dim lngError As Long
  Dim intErl As Integer
  Dim strMsg As String

  ' Variables to preserve error information
  strError = Err.Description
  lngError = Err.Number

  ' Prompt the user with information on the error:
  strMsg = "Procedure: " & CurrentProcName() & vbCrLf & _
           "Error : (" & lngError & ") " & strError & vbCrLf & _
           "Call Stack: " & GetCallStack & vbCrLf & _
           "PLEASE SCREENSHOT THIS MESSAGE AND EMAIL TO info@disarm.foundation"
  MsgBox strMsg, vbCritical, "DISARM: Error in Word Plug-In"

End Sub

Private Function CurrentProcName() As String
  CurrentProcName = mastrCallStack(mintStackPointer - 1)
End Function

Function GetCallStack() As String
  ' Returns a single string with the call stack
  Dim i As Integer
  Dim CallStack As String
  CallStack = ""
  For i = 1 To UBound(mastrCallStack)
    If mastrCallStack(i) = "" Then Exit For
    If i = 1 Then
        CallStack = mastrCallStack(i)
    Else
        CallStack = CallStack & "->" & mastrCallStack(i)
    End If
  Next i
  GetCallStack = CallStack
  'Debug.Print CallStack
End Function
Sub AdvancedErrorStructure()
  ' Template for using call stack and global error handler

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "AdvancedErrorStructure"

  ' << Your code here >>

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Sub TestAdvancedErrorStructure()
  ' Test for call stack and global error handler

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "TestAdvancedErrorStructure"

  AdvancedErrorStructure

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
