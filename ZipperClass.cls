VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private oFSO As New Scripting.FileSystemObject

Public Function ExtractFile(sZip As String, sFile As String, sOutputPath As String) As Boolean
    Zip_Activity "UNZIP", sZip, oFSO.BuildPath(sOutputPath, sFile)
End Function

Public Function AddFile(sZip As String, sName As String) As Boolean
    Zip_Activity "ZIPFILE", sName, sZip
End Function

