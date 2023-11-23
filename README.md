# Excel Integrierter Workflow
Ein erster Ansatz einen Arbeitsprozess für DevSecOps möglichst in Excel abzuwickeln

Bisher beinhaltet die CommitAddIn.xlsm Datei ein Workbook und VBA-Module die zu einem Excel Add-In (.xlam) umgespeichtert werden können mit den folgenden Funktionen:

    1. VBA-Module exportieren
    2. VBA-Module importieren
    3. Änderung im gegebenen Git Repo commiten
    4. Bestehende Commits auf Remote pushen
    5. Updates von Remote pullen
    6. In einem Schritt exportieren, commiten, und pushen
    7. In einem Schritten pullen und den Importprozes starten.



## Genaueres zu den Funktionen

### Export Prozess

#### Bestehende Funktionen:

In diesem Prozess werden die VBA-Module des aktiven Workbooks in einen Ordner namens "ActiveWorkbook.Name"_vba , der sich neben dem Workbook befindet, abgelegt als .bas-Dateien.

#### Gewünschte Funktionen:

Es wäre wünschenswert die Möglichkeit zuhaben die Module gesammelt an einen anderen Ort ablegen, da das Workbook wahrscheinlich nicht immer im Git Repository liegen wird?

#### Bemerkungen

Momentan werden beim Export Module im Ordner überschrieben. Dazu könnte man sich auch einen alternativen Prozess überlegen.

### Import Prozess

#### Bestehende Funktionen:

In diesem Prozess werden alle .bas-Dateien, die sich in einem vom Benutzer ausgewählten Ordner befinden, in das VBA-Projekt des aktiven Workbooks importiert.

Beim Importieren eines Moduls mit gleichem Namen wird der User darum gebeten sich zu entscheiden soll das alte Modul überschrieben werden, das neue Modul nicht importiert werden, oder soll das neue Modul unter einem anderen Namen abgelegt werden im VBA-Projekt.

#### Gewünschte Funktionen:

-

### Git Commit

#### Bestehende Funktionen:

In diesem Prozess werden die gespeicherten Änderungen in zu Git commitet. 
Es besteht die Option eine eigene Commit-Nachricht zu schreiben die am Ende mit "- $Office_Username" unterschrieben wird, oder eine standardisierte Nachricht (auch mit dem Benutzernamen unterschrieben) abzuschicken. 

#### Gewünschte Funktionen:

-

### Git Tag

#### Bestehende Funktionen:

Es wird ein neuer Commit erstellt und dieser wird dann mit einem neuen Tag versehen.

#### Gewünschte Funktionen:

-

### Git Tag-Retreival

#### Bestehende Funktionen:

Wir können entweder einzelne Dateien oder das Ganze Repository in dem Zustand eines bestimmten Tags zurückholen. Die zurückgeholten Dateien werden in einem Temporären Ordner abgelegt.

<mark> Man merke hier an, dass der Benutzer sich selber darum kümmern muss die alten Versionen zu entsorgen, wenn sie nich mehr benötigt sind. </mark>

#### Gewünschte Funktionen:

### Git Push

#### Bestehende Funktionen:

Die gemachten Commits werden an das remote Repository gepusht.
Unter der Annahme, dass das Workbook im lokalen Repository liegt.

#### Gewünschte Funktionen:

-

### Git Pull

#### Bestehende Funktionen:

Die neusten Updates werden vom Remote Repo gezogen.

#### Gewünschte Funktionen:
