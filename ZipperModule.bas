Attribute VB_Name = "ZipperModule"
Rem http://visualbasic.happycodings.com/files-directories-drives/code8.html

Option Explicit

'//source was in C# from urls:
'//http://www.codeproject.com/csharp/CompressWithWinShellAPICS.asp
'//http://www.codeproject.com/csharp/DecompressWinShellAPICS.asp

'//set reference to "Microsoft Shell Controls and Automation"


'http://forums.microsoft.com/MSDN/ShowPost.aspx?PostID=1090552&SiteID=1
'Be aware when using the shell automation interface to unzip files as it
'leaves copies of the zip files in the temp directory (defined by %TEMP%).
'Folders named "Temporary Directory X for demo.zip" are generated where X
'is a sequential number from 1 - 99.  When it reaches 99 you will then get
'a error dialog saying "The file exists" and it will not continue.
'I 've no idea why Windows doesn't clean up after itself when unzipping files,
'but it is most annoying...


'//CopyHere options
'0 Default. No options specified.
'4 Do not display a progress dialog box.
'8 Rename the target file if a file exists at the target location with the same name.
'16 Click "Yes to All" in any dialog box displayed.
'64 Preserve undo information, if possible.
'128 Perform the operation only if a wildcard file name (*.*) is specified.
'256 Display a progress dialog box but do not show the file names.
'512 Do not confirm the creation of a new directory if the operation requires one to be created.
'1024 Do not display a user interface if an error occurs.
'4096 Disable recursion.
'9182 Do not copy connected files as a group. Only copy the specified files.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Zip_Activity(Action As String, sFileSource As String, sFileDest As String)

    '//copies contents of folder to zip file
    Dim ShellClass  As Shell32.Shell
    Dim Filesource  As Shell32.Folder
    Dim Filedest    As Shell32.Folder
    Dim Folderitems As Shell32.Folderitems
    
    If sFileSource = "" Or sFileDest = "" Then GoTo EH
                
    Select Case UCase$(Action)
        
        Case "ZIPFILE"
            
            If Right$(UCase$(sFileDest), 4) <> ".ZIP" Then
                sFileDest = sFileDest & ".ZIP"
            End If
            
            If Not Create_Empty_Zip(sFileDest) Then
                GoTo EH
            End If
        
            Set ShellClass = New Shell32.Shell
            Set Filedest = ShellClass.NameSpace(sFileDest)
            
            Call Filedest.CopyHere(sFileSource, 1024 + 6 + 4)
                
        Case "ZIPFOLDER"
            
            If Right$(UCase$(sFileDest), 4) <> ".ZIP" Then
                sFileDest = sFileDest & ".ZIP"
            End If
            
            If Not Create_Empty_Zip(sFileDest) Then
                GoTo EH
            End If
        
            '//Copy a folder and its contents into the newly created zip file
            Set ShellClass = New Shell32.Shell
            Set Filesource = ShellClass.NameSpace(sFileSource)
            Set Filedest = ShellClass.NameSpace(sFileDest)
            Set Folderitems = Filesource.Items
            
            Call Filedest.CopyHere(Folderitems, 20)
        
        Case "UNZIP"
            
            If Right$(UCase$(sFileSource), 4) <> ".ZIP" Then
                sFileSource = sFileSource & ".ZIP"
            End If
            
            Set ShellClass = New Shell32.Shell
            Set Filesource = ShellClass.NameSpace(sFileSource)      '//should be zip file
            Set Filedest = ShellClass.NameSpace(sFileDest)          '//should be directory
            Set Folderitems = Filesource.Items                      '//copy zipped items to directory
            
            Call Filedest.CopyHere(Folderitems, 20)
        Case Else
        
    End Select
            
    '//Ziping a file using the Windows Shell API creates another thread where the zipping is executed.
    '//This means that it is possible that this console app would end before the zipping thread
    '//starts to execute which would cause the zip to never occur and you will end up with just
    '//an empty zip file. So wait a second and give the zipping thread time to get started.

    Call Sleep(1000)
    
EH:

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "error"
    End If

    Set ShellClass = Nothing
    Set Filesource = Nothing
    Set Filedest = Nothing
    Set Folderitems = Nothing

End Sub

Public Function Create_Empty_Zip(sFileName As String) As Boolean

    Dim EmptyZip()  As Byte
    Dim J           As Integer

    On Error GoTo EH
    Create_Empty_Zip = False

    '//create zip header
    ReDim EmptyZip(1 To 22)

    EmptyZip(1) = 80
    EmptyZip(2) = 75
    EmptyZip(3) = 5
    EmptyZip(4) = 6
    
    For J = 5 To UBound(EmptyZip)
        EmptyZip(J) = 0
    Next

    '//create empty zip file with header
    Open sFileName For Binary Access Write As #1

    For J = LBound(EmptyZip) To UBound(EmptyZip)
        Put #1, , EmptyZip(J)
    Next
    
    Close #1

    Create_Empty_Zip = True

EH:
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "Error"
    End If
    
End Function

