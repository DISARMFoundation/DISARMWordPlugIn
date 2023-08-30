Attribute VB_Name = "RibbonControl"
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
' Thanks to Greg Maxey for the code in this module
' See https://gregmaxey.com/word_tip_pages/customize_ribbon_main.html
'
' To modify the XML code for the Ribbon itself use the Custom UI Editor that Greg recommends at
' https://github.com/fernandreu/office-ribbonx-editor/releases
'

Option Explicit
Public myRibbon As IRibbonUI

Private Sub Onload(ribbon As IRibbonUI)
  'Creates a ribbon instance for use in this project
  'Note that this procedure is referenced in the Ribbon's XML:
  '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonControl.Onload">
  Set myRibbon = ribbon
End Sub

'Callback for Button onAction
Private Sub MyBtnMacro(ByVal control As IRibbonControl)

'
' For those buttons that involve tagging or summaries or the profile of the active document first ensure that the DISARM document name is current
' in case the user has activated a different document. Could use Application.DocumentChange event but just as easy to check the DISARM name
' every time a ribbon button is chosen.
'
' Also for those buttons that involve tagging or summaries or the profile of the active document we need to register the event handler which handles
' events with the Word.Application object so that we can, for example, check for the closing of any document, and if it is a DISARM document then
' clean up Excel when the last DISARM document is being closed.
'

Dim strPrefix As String

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "MyBtnMacro"
  
  Select Case control.ID
    Case Is = "Btn1c"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.InsertRedTag
    Case Is = "Btn1s"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.SearchTechniques
    Case Is = "Btn2c"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.InsertBlueTag
    Case Is = "Btn2s"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.SearchCountermeasures
    Case Is = "Btn3"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.InsertSummaryRedGraphic
    Case Is = "Btn4"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.InsertSummaryRedTable
    Case Is = "Btn5"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.InsertSummaryBlueTable
    Case Is = "Btn6r"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.FormatRed
    Case Is = "Btn6b"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.FormatBlue
    Case Is = "Btn6g"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.FormatRedGraphic
    Case Is = "Btn7"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.ClearHistory
    Case Is = "Btn8"
        strPrefix = modMain.DISARM_Name()
        modMain.Register_Event_Handler
        DISARM.CallListTags
    Case Is = "Btn9"
        DISARM.LinkToExplorer
    Case Is = "Btn10"
        DISARM.LinkToNavigator
    Case Is = "Btn11"
        DISARM.LinkToGitHub
    Case Else
      'Do nothing
  End Select

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub
