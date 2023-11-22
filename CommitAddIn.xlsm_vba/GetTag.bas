Option Explicit

Sub GitGetOld(ByRef control As Office.IRibbonControl)

    ' Get Userinput as to whether we want to retrieve the entire repo at version vX.XX
    ' or whether it should only retrieve a certain file

End Sub


Sub TagFullRetrieval()

' Variablen:
    Dim temp As Integer
    Dim fs As Object
    Dim tempDirectory As String
    Dim tempBranch As String
    Dim gitURL As String
    Dim versionTag As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    versionTag = "v1.0"
    tempDirectory = "temp_" & versionTag
    tempBranch = "tempBranch"
    
    
'-------------------------------------------------------
' Pfad:
    Pathing
    
'-------------------------------------------------------
    If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
    End If
    
    gitURL = GetShellOutput("git config --get remote.origin.url")
    gitURL = Replace(gitURL, vbLf, "")
    
    temp = ShellCommand("git clone --branch " & versionTag & " --single-branch " & gitURL & " " & tempDirectory, "Das Repository wurde in den Ordner " & tempDirectoy & "geladen.", "Die ältere Version des Repositorys konnte nicht geladen werden.")
        
End Sub

Sub TagFileRetrieval()

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
    versionTag = "v1.0"
    tempFile = tempDirectory & "\" & Replace(versionTag, ".", "_") & "_" & oldFile
    tempVersionTag = Replace(versionTag, ".", "_")
    
    gitCommand = "cmd.exe /C git show " & versionTag & ":" & oldFile & " > " & tempFile
    
    Debug.Print gitCommand
    
    temp = ShellCommand(gitCommand, "Die alte Version von " & oldFile & " wurde erfolgreich im Ordner " & tempDirectory & " abgelegt.", "Der Vorgang ist gescheitert, versuchen Sie es nochmal oder manuell.")


End Sub

Sub TagCleanUp()

' Variablen:

    
'-----------------------------------
' Pfad:
    Pathing
    
'-----------------------------------
' Suche nach "temp_*" Ordnern


'------------------------------------
' die gefundenen ordner löschen

'----------------------------------
' Suche nach temp Ordner


End Sub