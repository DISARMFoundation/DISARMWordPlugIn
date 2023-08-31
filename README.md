# DISARMWordPlugIn
 DISARM Plug-In for Microsoft Word on Windows desktop.

The DISARM Word Plug-In supports Microsoft Word 32-bit or 64-bit running on Microsoft Windows 10.

It offers an easy way to tag or map adversary behaviors and defender mitigations in the text of a Word document to techniques and countermeasures in DISARM.
The Plug-In keeps a running history of your tags so that you can insert summaries into your document in tabular or graphical format.
To install the Plug-In download and run the MSI file [Install_Files\DISARM_Word_PlugIn.msi](https://github.com/DISARMFoundation/DISARMWordPlugIn/blob/main/Install_Files/DISARM_Word_PlugIn.msi). This installs a macro-enabled global template into the Word STARTUP 
directory and two Excel files into the Excel XLSTART directory in your user profile. These files will be automatically loaded the next time you start Word.
The Plug-In functionality is available from the DISARM menu option in the Microsoft Word ribbon.  

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the   GNU General Public License at https://www.gnu.org/licenses/ for more details.

To modify the ribbon use the Custom UI Editor at https://github.com/fernandreu/office-ribbonx-editor/releases. To modify the VBA code, open the macro-enabled template which is found in the USERPROFILE e.g. C:\Users\bob under \AppData\Roaming\Microsoft\Word\STARTUP, and enable the Developer tab via File->Options->Customize Ribbon. If you would like to contribute your code changes into our repo, clone our repo and update any source files you change by exporting them via the VBA Editor, pushing them up to Github, and creating a pull request (you can export all VBA files by opening (Install_Files/DISARM_TAGGER_INSTALL.docm) and running the exportVBA procedure). 

To report bugs or suggestions for improvement please send email to info@disarm.foundation. We are working on rewriting the Plug-In in Javascript and Python to create an official Microsoft Add-on. 
