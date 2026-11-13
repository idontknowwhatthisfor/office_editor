#!/bin/sh
# Opens a .docx in Microsoft Word and saves Word's own PDF of it next to the file: the check that
# Word reads an edited file without complaint, and a picture of how Word lays it out.
#
#     sh test-files/word_to_pdf.sh test-files/x.docx
#
# `open -a` hands the file to Word the way Finder does, so Word's sandbox lets it in without
# asking. Word can only write to ~/Downloads, so the PDF goes there and is then moved.
set -e
file=$(cd "$(dirname "$1")" && pwd)/$(basename "$1")
name=$(basename "$file")
pdf="${name%.*}.pdf"
open -a "Microsoft Word" "$file"
sleep 6
osascript -e "tell application \"Microsoft Word\"
save as document \"$name\" file name (POSIX file ((POSIX path of (path to downloads folder)) & \"$pdf\")) file format format PDF
close document \"$name\" saving no
end tell"
mv "$HOME/Downloads/$pdf" "$(dirname "$file")/$pdf"
echo "$(dirname "$file")/$pdf"
