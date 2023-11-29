Option Explicit

Sub GitGetOld(ByRef control As Office.IRibbonControl)

   
    Dim myForm As Object
    
    Set myForm = New RetrievalForm
    
    myForm.Show


End Sub

' A function that clones the repository one is currently using into a folder called "temp" next to the workbook, that contains the current repository checked out at a chosen tag.
Function TagFullRetrieval(ByVal version As String)

' Variables:

    Dim temp As Integer
    Dim fs As Object
    Dim tempDirectory As String
    Dim tempSubDirectory As String
    Dim tempBranch As String
    Dim gitURL As String
    Dim gitCommand As String
    Dim versionTag As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    versionTag = version
    tempDirectory = "temp"
    tempSubDirectory = "temp_" & versionTag
    tempBranch = "tempBranch"
    
    
'-------------------------------------------------------
' Getting the right path:

    Pathing
    
'-------------------------------------------------------
' Creating the new folders if they don't exist yet.

    If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
    End If
    
                
    If Not fs.FolderExists(tempDirectory & "\" & tempSubDirectory) Then
        fs.CreateFolder tempDirectory & "\" & tempSubDirectory
    End If
    
'----------------------------------------------------------
' Retrieve adress of the current repository in order to clone it into the specified subdirectory

    gitURL = GetShellOutput("git config --get remote.origin.url")
    gitURL = Replace(gitURL, vbLf, "")

'----------------------------------------------------------
' Execute the shell command
    
    gitCommand = "git clone --branch " & versionTag & " --single-branch " & gitURL & " " & tempDirectory & "\" & tempSubDirectory
    
    temp = ShellCommand(gitCommand, "Das Repository wurde in den Ordner " & tempDirectory & "\" & tempSubDirectory & "geladen.", "Die ältere Version des Repositorys konnte nicht geladen werden.")
        
End Function

' A function that checks out a specific file at a certain tag and saves in the temp subdirectory.
Function TagFileRetrieval(ByVal version As String)

    Dim temp As Integer
    Dim gitCommand As String
    Dim versionTag As String
    Dim oldFile As String
    Dim tempFile As String
    Dim tempVersionTag As String
    Dim fs As Object
    Dim tempDirectory As String
    
'-------------------------------------------------------
' Getting the right location:
    Pathing
    
'-------------------------------------------------------
' Creating the temp directory:

    Set fs = CreateObject("Scripting.FileSystemObject")
    tempDirectory = "temp"
    
    If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
    End If
    
'-----------------------------------------------------
' Getting the desired file to be checked out:
    
    oldFile = UserPromptText("Welche Datei möchten Sie laden?", "Datei auswählen", "", "Filename")
    
   
    
    versionTag = version
    tempFile = tempDirectory & "\" & Replace(versionTag, ".", "_") & "_" & oldFile
    tempVersionTag = Replace(versionTag, ".", "_")
    
'-----------------------------------------------------
' Executing the desired command:

    gitCommand = "cmd.exe /C git show " & versionTag & ":" & oldFile & " > " & tempFile
    
    'Debug.Print gitCommand
    
    temp = ShellCommand(gitCommand, "Die alte Version von " & oldFile & " wurde erfolgreich im Ordner " & tempDirectory & " abgelegt.", "Der Vorgang ist gescheitert, versuchen Sie es nochmal oder manuell.")


End Function

' A function that retrieves the tags that exist in the current repository.
Function FindTags()

    Dim existingTagsRaw As String
    Dim existingTags() As String
    Dim i As Integer
    
    Pathing
    
    existingTagsRaw = GetShellOutput("git tag")
    
    existingTags = Split(existingTagsRaw, vbLf)
    
    ReDim Preserve existingTags(UBound(existingTags) - 1)
    
    FindTags = existingTags
    
End Function