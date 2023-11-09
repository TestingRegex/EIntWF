#!/bin/bash

find . -type d -name "*_vba" > vbaDirectories.txt
cat vbaDirectories.txt


phrase="password"
echo "Hier suchen wir innerhalb der _vba Ordner nach dem String '$phrase'."
vbaDirectories=$(cat vbaDirectories.txt)
rm -f vbaDirectories.txt
echo $vbaDirectories

# Loop through the list of "_vba" directories and scan files for "password"
touch passwordFiles.txt
for vbaDirectory in $vbaDirectories; do
    echo "Der Ordner '$vbaDirectory' wird geprüft."
    grep -Rnwl $vbaDirectory -e $phrase |            
        while read -r file; do
            echo "Die Datei '$file' enthält den String '$phrase'"
            echo "$file" >> passwordFiles.txt
        done
done
if [ -s passwordFiles.txt ] ; then
    echo "Diese Dateien enthalten den String '$phrase':"
    cat passwordFiles.txt
    rm  -f passwordFiles.txt
    exit 1
else
    echo "Keine Dateien enthalten den String '$phrase'."
    rm  -f passwordFiles.txt 
    exit 0
fi