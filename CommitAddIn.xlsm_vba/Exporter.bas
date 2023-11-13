'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Exportieren der Module im VBA Projekt übernimmt
'
'
'
'''

Option Explicit


Sub ExtractModulesFromWorkbook(control As Office.IRibbonControl)

    Export
    
End Sub

Function Export()

    Dim wb As Workbook
    Dim WorkbookName As String
    Dim vbComp As Object
    Dim vbProj As Object
    Dim moduleName As String
    Dim moduleCode As String
    Dim outPath As String
    Dim modulePath As String
    Dim fileSysObj As Object
    Dim fs As Object

'---------------------------------------------------------------------------------------------
' Der Pfad zum Exportordner wird gefunden

    Set wb = ActiveWorkbook
    WorkbookName = wb.Name


    outPath = wb.path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim vbaDirectory As String
    vbaDirectory = outPath & "\" & WorkbookName & "_vba\"
    
'---------------------------------------------------------------------------------------------
' Exportordner wird erstellt falls noch nicht vorhanden

    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If

'---------------------------------------------------------------------------------------------
' Die Module des VBA Projekts werden an den gewünschten Ort Exportiert


    ' Es wird durch alle Komponenten des VBA Projekts durch iteriert und alle Module werden exportiert.
    For Each vbProj In wb.VBProject.VBComponents
        If vbProj.Type = 1 Then ' Module
            moduleName = vbProj.Name
            
            ' Prüfen ob das Modul nicht einfach leer ist.
            If vbProj.CodeModule.CountOfLines > 0 Then
            
                moduleCode = vbProj.CodeModule.Lines(1, vbProj.CodeModule.CountOfLines)
            
                ' Inhalt des Moduls wird als String Variable geladen
                modulePath = vbaDirectory & moduleName & ".bas"
                
                ' Prüfen ob das Modul oder ein Modul mit diesem Namen bereits im Exportordner existiert
                If fs.FileExists(modulePath) Then
                
                    ' Inhalt der gleichnamigen Datei einladen
                    Dim textStream As Object
                    Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading

                    Dim fileContent As String
                    fileContent = textStream.ReadAll
                    textStream.Close

                    ' Prüfen ob der Inhalt der Datei und der des Moduls sich unterscheiden falls ja wird die Datei überschrieben
                    If fileContent <> moduleCode Then
                        
                        Dim textStreamOverwrite As Object
                        Set textStreamOverwrite = fs.CreateTextFile(modulePath, True)
                        textStreamOverwrite.Write moduleCode
                        textStreamOverwrite.Close
                    End If
                ' Modul wurde unter dem jetzigen Namen noch nicht exportiert, dementsprechend einfach exportieren.
                Else
                
                ' Neue .bas Datei wird erstellt und mit dem Modul inhalt gefüllt
                Dim textStreamNew As Object
                Set textStreamNew = fs.CreateTextFile(modulePath, True)
            
                textStreamNew.Write moduleCode
                textStreamNew.Close
            
                Debug.Print "Module Name: " & moduleName
                Debug.Print moduleCode
                End If
            End If
        End If
    Next vbProj
    
'---------------------------------------------------------------------------------------------
' Aufräumen
    
    Set fs = Nothing
    Set vbComp = Nothing
    Set wb = Nothing
    Set vbProj = Nothing
    Set fileSysObj = Nothing
End Function
