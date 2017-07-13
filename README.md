# Reports-Microsoft-Word
This repository contains a complete implementation example/demo of how to use this 'module'

This is a 'module' designed in Microsoft Visual Studio Project using C#. It makes use of the System.Windows.Forms and allows multiple instances of this 'module' to be appended to any other Microsoft Visual Studio Project.

The module consist of 5 parts:
1. The physical GUI that is placed onto a he System.Windows.Forms.Pannel          	=> Report.cs
2. The Control_Events class (the actual code behind the physical GUI              	=> ControlEvents_Report.cs
3. The "DocX.dll" API made by Novacode                                            	=> DocX.dll
4. The interfacing class (a class that easily talks between API and main program) 	=> Reporting_Word.cs
5. The Template.docx document (contains bookmarks as placeholders)					=> BookmarkTemplate.docx

The core files required for this module are located here:
Reports-Microsoft-Word\ReportDemo\ReportDemo\Module_Report

To use this module all you need to do is:
A. Create an instance of 1. above (the main GUI)
B. Create a new instance of a System.Windows.Forms.Pannel
C. Initialise the instance in B with the instance of A (set/link the two panels = to each other)
D. Add the instance in B to your new project.

Congrats, you have now added this 'module' to your Microsoft Visual Studio Project
