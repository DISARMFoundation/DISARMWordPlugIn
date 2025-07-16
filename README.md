# DISARMWordPlugIn
 DISARM Plug-In for Microsoft Word on Windows desktop.

UPDATE July 16 2025.
**This repository is no longer actively maintained**. We have redesigned the software using JavaScript as an official Add-In for Microsoft Word available in
the Microsoft Store [here](https://appsource.microsoft.com/en-us/product/office/wa200008045?tab=overview). The official Add-In runs within an isolated
process which has no access to the local filesystem and is therefore more secure than the DISARM Plug-In for Microsoft Word, which was written using VBA. 
It also supports Microsoft Word running on MacOS and within standard web browsers. We are in the process of making the source code for the official DISARM Add-In 
for Microsoft Word available in GitHub, starting with the server code [here](https://github.com/Digiqal-development/disarm-server). We plan to provide bug fix 
support for the VBA DISARM Plug-In for Microsoft Word until the end of 2025.

------------------------------------------------------------------------------------------------
The DISARM Word Plug-In supports Microsoft Word 32-bit or 64-bit running on Microsoft Windows 10.

It offers an easy way to tag or map adversary behaviors and defender mitigations in the text of a Word document to techniques and countermeasures in DISARM.
The Plug-In keeps a running history of your tags so that you can insert summaries into your document in tabular or graphical format.
To install the Plug-In download and run the MSI file [Install_Files\DISARM_Word_PlugIn.msi](https://github.com/DISARMFoundation/DISARMWordPlugIn/blob/main/Install_Files/DISARM_Word_PlugIn.msi). This installs a macro-enabled global template into the Word STARTUP 
directory and two Excel files into the Excel XLSTART directory in your user profile. These files will be automatically loaded the next time you start Word.
To uninstall the Plug-In press the Windows Start key then select Settings->Apps and scroll down until you see DISARM. Make sure you exit Microsoft Word before uninstalling the Plug-In. 
The Plug-In functionality is available from the DISARM menu option in the Microsoft Word ribbon.  

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the   GNU General Public License at https://www.gnu.org/licenses/ for more details.

To modify the ribbon use the Custom UI Editor at https://github.com/fernandreu/office-ribbonx-editor/releases. To modify the VBA code, open the macro-enabled template which is found in the USERPROFILE e.g. C:\Users\bob under \AppData\Roaming\Microsoft\Word\STARTUP, and enable the Developer tab via File->Options->Customize Ribbon. If you would like to contribute your code changes into our repo, clone our repo and update any source files you change by exporting them via the VBA Editor, pushing them up to Github, and creating a pull request (you can export all VBA files by opening [Install_Files/DISARM_TAGGER_INSTALL.docm](Install_Files/DISARM_TAGGER_INSTALL.docm) and running the exportVBA procedure). 

To report bugs or suggestions for improvement please send email to info@disarm.foundation. We are working on rewriting the Plug-In in Javascript and Python to create an official Microsoft Add-on. 
