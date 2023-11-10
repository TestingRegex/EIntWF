# Excel Integrierter Workflow
Ein erster Ansatz einen Arbeitsprozess für DevSecOps möglichst in Excel abzuwickeln

Bisher beinhaltet die CommitAddIn.xlsm Datei ein Workbook und VBA-Module die zu einem Excel Add-In (.xlam) umgespeichtert werden können mit den folgenden Funktionen:

1. VBA-Module exportieren
2. VBA-Module importieren
3. Änderung im gegebenen Git Repo commiten
4. Bestehende Commits auf Remote pushen
5. Updates von Remote pullen
6. In einem Schritt exportieren, committen, und pushen
7. In einem Schritten pullen und den Importprozes starten.



## Genaueres zu den Funktionen

### Export Prozess

#### Bestehende Funktion

In diesem Prozess werden die VBA-Module des aktiven Workbooks in einen Ordner namens "ActiveWorkbook.Name"_vba , der sich neben dem Workbook befindet, abgelegt als .bas-Dateien.

#### Gewünschte Funktionalitäten:

Es wäre wünschenswert die Möglichkeit zuhaben die Module gesammelt an einen anderen Ort ablegen, da das Workbook wahrscheinlich nicht immer im Git Repository liegen wird?

#### Bemerkungen

Momentan werden beim Export Module im Ordner überschrieben. Dazu könnte man sich auch einen alternativen Prozess überlegen.

### Import Prozess

#### Bestehende Funktion

In diesem Prozess werden alle .bas-Dateien, die sich in einem vom Benutzer ausgewählten Ordner befinden, in das VBA-Projekt des aktiven Workbooks importiert.

Beim Importieren eines Moduls mit gleichem Namen wird der User darum gebeten sich zu entscheiden soll das alte Modul überschrieben werden, das neue Modul nicht importiert werden, oder soll das neue Modul unter einem anderen Namen abgelegt werden im VBA-Projekt.

#### Gewünschte Funktionalitäten:

-

### Git Commit

#### Bestehende Funktion
In diesem Prozess werden 

#### Gewünschte Funktionalitäten:

### Git Push

#### Bestehende Funktion
In diesem Prozess werden

#### Gewünschte Funktionalitäten:

### Git Pull

#### Bestehende Funktion
In diesem Prozess werden

#### Gewünschte Funktionalitäten:
