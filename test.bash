#!/bin/bash


rm -f passwordFiles.log
rm -f protectUnprotect.log
find . -type d -name "*_vba" > vbaDirectories.txt

vbaDirectories=$(cat vbaDirectories.txt)
#echo $vbaDirectories
rm -f vbaDirectories.txt

PHRASE="password\|Passwort\|pwd\|passcode\|Passcode\|Password\|passwort\|pwt"

# Loop through the list of "_vba" directories and scan files for "password"
#touch passwordFiles.log

foundFiles=false

#for phrase in $PHRASE; do
#    echo $'\n'"Dateien mit dem Ausdruch: '$phrase'"$'\n'" ">> passwordFiles.log
for vbaDirectory in $vbaDirectories; do
    echo "$vbaDirectory"
    grep -rle '.*\(password\|Passwort\|pwd\|passcode\|Passcode\|Password\|passwort\|pwt\).*=.*\".*\"' "$vbaDirectory" >> passwordFiles.log            
#    while read -r file; do
#        echo $foundFiles
#        echo "$file contains a suspicious use of one of the phrases."
#        echo "$file" >> passwordFiles.log
#        foundFiles=true
#    done || true
done
#done

if $foundFiles; then
    echo "Es wurden auffällige Dateien gefunden."
    cat passwordFiles.log
    foundFiles=false
fi



#Using Regex with grep.
touch protectUnprotect.log
foundProtect=false

for vbaDirectory in $vbaDirectories; do
    grep -Rnwl $vbaDirectory -e '\(Protect\|Unprotect\) \".*\"' |            
    while read -r file; do
        echo "$file" >> protectUnprotect.log
        foundProtect=true
    done || true
done

if [ $foundProtect ]; then
    echo "Es wurden auffällige Uses von Protect/Unprotect in den folgenden Macros gefunden:"
    cat protectUnprotect.log
fi

