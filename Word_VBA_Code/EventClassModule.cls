VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' This class module is used to clean up before the last open document being tagged is closed. There can be multiple documents open and being tagged
' at any one time. We do not want to close the Excel workbooks until the very last one is being closed. Of course, the user may cancel out of the close
' and save and this routine will still close the Excel workbooks and also close Excel. No worries, the oApp checks will cause Excel to be fired up
' again and the workbooks opened.
'
' For details of using Application object events see:
' https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word.
'
' For explicit guidance on using the Application.DocumentBeforeClose event see:
' https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforeclose
'
' For an example frpom stackoverflow see:
' https://stackoverflow.com/questions/59017097/creating-event-for-word
'

Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)

'
' This event runs before the user closes a Word document. It is used to clean up Excel if this is the last DISARM
' enabled document to be closed.
'

Dim i As Integer
Dim j As Integer
Dim strPrefix As String
On Error Resume Next

j = 0

For i = 1 To Application.Documents.Count
    strPrefix = Application.Documents(i).Variables("DISARM_Name")
    If Err.Number = 0 Then
        j = j + 1
    End If
Next

If j = 1 Then
    modMain.CleanUp
End If

End Sub
