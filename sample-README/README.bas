Attribute VB_Name = "README"
'#Module README
'_____________________________________________________________________________________
'DOCUMENTATION
'See git repository at https://...
'_____________________________________________________________________________________
'LICENSE
'All rights reserved. See license file in the git repository at https://...
'_____________________________________________________________________________________
'SUPPORT
'Contact name@company.com
'__________________________________________________________________________________
'MAINTANENCE
'All code is versioned in Git. For change tracking, the source code has to be exported
'before every commit.
'
'This project is combines the following programming patterns:
' - Singleton
' - Dependency Injection (DI)
' - Auto Properties
' - Model-Viewer-Controller (MVC)
' - Data Transfer Object (DTO)
'
'Class dependency tree: Is is important not to violate the singleton dependence tree
'otherwise the app will run out of stack and crash. Singleton dependancies are noted
'clearly in the Initialization of each class.
'_____________________________________________________________________________________
'DEPENDANCIES
'Add lib ref to:
' - Microsoft Forms 2.0 Object Library
' - Microsoft ActiveX Data Objects Recordset 6.0 Library
' - Microsoft Scripting Runtime
' - Microsoft ActiveX Data Object 6.1 Library
'_____________________________________________________________________________________
'ERROR HANDLING
'Error handling is concentrated in the db connection management, and control modules
'_____________________________________________________________________________________
'LIFE CYCLE
'(1) Change expected values in the testing modules/classes
'(2) Run all tests in the test module
'(3) Export code with the code module
'(4) Commit to development branch in git
'(5) Create pull/review request
'(6) The tech lead adds the new manifest into the app registry
'(7) Distribute new version
'_____________________________________________________________________________________
'SECURITY
'(undisclosed)
'_____________________________________________________________________________________
'PUBLISHER
' (c) Company Name






