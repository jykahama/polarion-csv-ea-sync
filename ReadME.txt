// DEPENDENCIES
http://www.codeproject.com/script/articles/download.aspx?file=/KB/database/CsvReader/CsvReader_src.zip&rp=http://www.codeproject.com/Articles/9258/A-Fast-CSV-Reader
http://www.c-sharpcorner.com/UploadFile/mgold/ScrollableMessageBox07292007223713PM/download/ScrollableMessageBox.zip

//START
- Start a new VS2003 Project
- Import the demo Code.

//REFERENCES
- In the Solution viewer right click on "References" and select the COM tab.
- Hit the browse button and navigate to your local EA directory.
- Select the Interop.EA.dll and click open then ok to close the Add Reference Dialog.

//PROPERTIES
- In the VS2003 menu select Project | CS-AddinFramework
- Select Common Properties | References Path and add C:\Program Files\Sparx Systems\EA\ if its not there already.
- Select Common Properties | General | Output Type to be set to Class Library
- Select Configuration Properties | Build | Register for COM Interop set to True.
- Select Configuration Properties | Debugging and Change the debug mode to Program and the start application to "C:\Program Files\Sparx Systems\EA\EA.exe" (or whatever is relevant to you).

//REGISTRY
- To setup the registry, we need to start Regedit and find the HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins directory.
- Add a new key with the name the same as your namespace. In our case it is CS_AddinFramework 
- The value for the key should be your project name followed by your entry point. 
- In our case the value would be: CS_AddinFramework.Main

//REGEDIT EXPORT
 [HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins\CS_AddinFramework]
 @="CS_AddinFramework.Main"
 
//RUN]
- Make sure EA is closed and Hit F5 (Start).
- VS2003 will load EA and all things correct you will see your Add-In under the Add-Ins menu.
- If EA does not load, run EA manually and access your Add-in

