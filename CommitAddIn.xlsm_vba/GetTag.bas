Attribute VB_Name = "GetTag"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   This module contains the macros and major functions used in the 'Version Laden'
'   button.
'
'   Purpose:
'       The button initiates the process of 'git show'-ing either a single file (that has to be
'       typed in correctly at the moment, it may be nice to see if we can have users select the file
'       from a list?) or the entire repository inside of the "temp" folder that is ignored by
'       the repository.
'
'   Used homemade functions/forms:
'       AnnoyUsers, Pathing, BadCharacterFilter, RetrievalForm, GitVersionCheckForm, UserPromptText,
'       ShellCommand
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Sub GitGetOld(ByVal control As Office.IRibbonControl)

On Error GoTo ErrHandler
    
    Pull
    Dim myForm As Object
    Set myForm = New RetrievalForm
    myForm.Show

ExitSub:

    Exit Sub
    
ErrHandler:

    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume

End Sub

' A function that clones the repository one is currently using into a folder called "temp" next to the workbook, that contains the current repository checked out at a chosen tag.
Public Sub TagFullRetrieval(ByVal version As String)

' Variables:


    Dim fileSystemObject As Object
    Dim tempDirectory As String
    Dim tempSubDirectory As String
    Dim gitURL As String
    Dim gitCommand As String
    Dim versionTag As String
    
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    versionTag = version
    tempDirectory = "temp"
    tempSubDirectory = "temp_" & versionTag
    
    
'-------------------------------------------------------
' Getting the right path:

    Pathing
    
'-------------------------------------------------------
' Creating the new folders if they don't exist yet.

    If Not fileSystemObject.FolderExists(tempDirectory) Then
        fileSystemObject.CreateFolder tempDirectory
    End If
    
                
    If Not fileSystemObject.FolderExists(tempDirectory & "\" & tempSubDirectory) Then
        fileSystemObject.CreateFolder tempDirectory & "\" & tempSubDirectory
    End If
    
'----------------------------------------------------------
' Retrieve adress of the current repository in order to clone it into the specified subdirectory

    gitURL = GetShellOutput("git config --get remote.origin.url")
    gitURL = Replace(gitURL, vbLf, vbNullString)

'----------------------------------------------------------
' Execute the shell command
    
    gitCommand = "git clone --branch " & versionTag & " --single-branch " & gitURL & " " & tempDirectory & "\" & tempSubDirectory
    
    ShellCommand gitCommand, "Das Repository wurde in den Ordner " & tempDirectory & "\" & tempSubDirectory & "geladen.", "Die ältere Version des Repositorys konnte nicht geladen werden.", "TagFullRetrieval"
        
End Sub

' A function that checks out a specific file at a certain tag and saves in the temp subdirectory.
Public Sub TagFileRetrieval(ByVal version As String)

    Dim gitCommand As String
    Dim versionTag As String
    Dim oldFile As String
    Dim tempFile As String
    Dim fileSystemObject As Object
    Dim tempDirectory As String
    
'-------------------------------------------------------
' Getting the right location:
    Pathing
    
'-------------------------------------------------------
' Creating the temp directory:

    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    tempDirectory = "temp"
    
    If Not fileSystemObject.FolderExists(tempDirectory) Then
                fileSystemObject.CreateFolder tempDirectory
    End If
    
'-----------------------------------------------------
' Getting the desired file to be checked out:
    
    oldFile = UserPromptText("Welche Datei möchten Sie laden?", "Datei auswählen", vbNullString, "Filename")
       
    
    versionTag = version
    tempFile = tempDirectory & "\" & Replace(versionTag, ".", "_") & "_" & oldFile
    
'-----------------------------------------------------
' Executing the desired command:

    gitCommand = "cmd.exe /C git show " & versionTag & ":" & oldFile & " > " & tempFile
    
    'Debug.Print gitCommand
    
    ShellCommand gitCommand, "Die alte Version von " & oldFile & " wurde erfolgreich im Ordner " & tempDirectory & " abgelegt.", "Der Vorgang ist gescheitert, versuchen Sie es erneut oder über die Commandline.", "TagFileRetrieval"

End Sub


