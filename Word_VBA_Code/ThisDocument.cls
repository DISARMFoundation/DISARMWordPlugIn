VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This is a macro-enabled global template containing the XML and VBA for the DISARM Word Plug-In.

' Copyright (C) 2023 DISARM Foundation

' This program is free software: you can redistribute it and/or modify it under the terms of the GNU 'General Public License
' as published by the Free Software Foundation, either version 3 of the License, or 'any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License at https://www.gnu.org/licenses/ for more details.

' To report bugs or suggestions for improvement please send email to info@disarm.foundation.
'___________________________________________________________________________________________________________________________

Private Sub Document_Close()

'
' This procedure will only be called if the user has opened a document based on DISARM_TAGGER.dotm and now wishes to close it
' or has opened DISARM_TAGGER.dotm itself and now wishes to close it. See procedure App_DocumentBeforeClose with the EventClassModule
' for the routine that will normally be used to cleanup the Excel workbooks and application - the normal usage is that the macros
' in DISARM_TAGGER.dotm are called via the Word ribbon directly without having to open DISARM_TAGGER.dotm explicitly.
'

modMain.CleanUp

End Sub

