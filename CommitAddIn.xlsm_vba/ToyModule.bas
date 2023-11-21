'''
'
' A module to contain temporary tests that are used to test functions of my own or ones that are new to me,
' this module is regularly reset and cleared.
'
'''

Option Explicit


Sub TagRetrieval()

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
    
    temp = ShellCommand("git clone --branch " & versionTag & " --single-branch " & gitURL & " " & tempDirectory, "The cloning and checkout worked!", "We might have a problem.")
        
End Sub

Sub Testing()

    tempBranch = "tempBranch"
    temp = ShellCommand("git checkout -b " & tempBranch & " tags/v1.0", "checkout temp 1", "checkout temp 0")
    tempDirectory = "temp_v1-0"
    If temp = 0 Then
    
        temp = ShellCommand("git checkout FeatureTagging", "In Feature branch", "Not in Feature Branch")
        
        If temp = 0 Then
            If Not fs.FolderExists(tempDirectory) Then
                fs.CreateFolder tempDirectory
            End If
            
            temp = ShellCommand("git worktree add ./" & tempDirectory & " " & tempBranch, "Added tempBranch to tempDirectory", "Last step fail")
        End If
    End If
    
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




End Sub

