Option Explicit

Sub GitGetOld(ByRef control As Office.IRibbonControl)

    Dim myForm As Object
    
    Set myForm = New RetrievalForm
    
    myForm.Show


End Sub


Function TagFullRetrieval(ByVal version As String)

' Variablen:
    Dim temp As Integer
    Dim fs As Object
    Dim tempDirectory As String
    Dim tempSubDirectory As String
    Dim tempBranch As String
    Dim gitURL As String
    Dim versionTag As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    versionTag = version
    tempDirectory = "temp"
    tempSubDirectory = "temp_" & versionTag
    tempBranch = "tempBranch"
    
    
'-------------------------------------------------------
' Pfad:
    Pathing
    
'-------------------------------------------------------
' Creating the new folders if they don't exist yet.
    If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
    End If
    
                
    If Not fs.FolderExists(tempDirectory & "\" & tempSubDirectory) Then
        fs.CreateFolder tempDirectory & "\" & tempSubDirectory
    End If
    
    
    gitURL = GetShellOutput("git config --get remote.origin.url")
    gitURL = Replace(gitURL, vbLf, "")
    
    temp = ShellCommand("git clone --branch " & versionTag & " --single-branch " & gitURL & " " & tempDirectory & "\" & tempSubDirectory, "Das Repository wurde in den Ordner " & tempDirectory & "\" & tempSubDirectory & "geladen.", "Die ältere Version des Repositorys konnte nicht geladen werden.")
        
End Function

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
' Pfad:
    Pathing
    
'-------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    tempDirectory = "temp"
    
    If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
    End If
    
    oldFile = "CommitAddIn.xlsm"
    versionTag = version
    tempFile = tempDirectory & "\" & Replace(versionTag, ".", "_") & "_" & oldFile
    tempVersionTag = Replace(versionTag, ".", "_")
    
    gitCommand = "cmd.exe /C git show " & versionTag & ":" & oldFile & " > " & tempFile
    
    Debug.Print gitCommand
    
    temp = ShellCommand(gitCommand, "Die alte Version von " & oldFile & " wurde erfolgreich im Ordner " & tempDirectory & " abgelegt.", "Der Vorgang ist gescheitert, versuchen Sie es nochmal oder manuell.")


End Function

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