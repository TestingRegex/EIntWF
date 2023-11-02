#!/bin/bash

# GPG Key ID (replace with your key ID)
KEY_ID="6611F2830A4B3C794DD9FF18D7DB7E9EAAEADD8C"


# Use find to locate subdirectories ending with "_vba"
#find . -type d -name "*_vba" | while read -r SUBDIR; do
#    echo "Signing files in $SUBDIR"


    # Loop through files in the subdirectory and sign them
#    find "$SUBDIR" -type f -not -name "*.sig" | while read -r file; do
#        gpg --local-user "$KEY_ID" --detach-sign "$file"

#    done
#done

# Get a list of staged files in Git
staged_files=$(git diff --name-only --cached)
echo $staged_files
# Loop through the staged files and process corresponding directories
for file in $staged_files; do
    # Check if there's a corresponding directory in the structure
    SUBDIR="$file"_vba
    if [ -d "$SUBDIR" ]; then
        echo "Signing files in $SUBDIR"

        # Loop through files in the subdirectory and sign them
        find "$SUBDIR" -type f -not -name "*.gpg" | grep -v "UnsignedModule.bas" | while read -r vba_file; do
            gpg --sign "$vba_file"
	     git add $(vba_file + ".gpg")
        done
    fi
    # Externally sign the workbooks as well
    if [[ "$file" =~ \.xlsm$ ]]; then
	echo "Signing the workbook " + $file
	gpg --sign "$file" 
	git add $(file + ".gpg")
    fi
done
