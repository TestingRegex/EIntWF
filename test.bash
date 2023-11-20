#!/bin/bash

PHRASE="password Passwort pwd passcode Passcode Password passwort pwt"

rm -f passwordFiles.txt
find . -type d -name "*_vba" > vbaDirectories.txt

vbaDirectories=$(cat vbaDirectories.txt)
rm -f vbaDirectories.txt

# Loop through the list of "_vba" directories and scan files for "password"
touch passwordFiles.txt

for phrase in $PHRASE; do
    echo $'\n'"Dateien mit dem Ausdruch: '$phrase'"$'\n'" ">> passwordFiles.txt
    for vbaDirectory in $vbaDirectories; do
        grep -Rnwl $vbaDirectory -e ".*$phrase.*" |            
        while read -r file; do
            echo "      $file" >> passwordFiles.txt
            foundFiles=true
        done || true
    done
done

if $foundFiles; then
    echo "Es wurden auffällige Dateien gefunden."
    cat passwordFiles.txt
fi