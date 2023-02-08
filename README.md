# Docx_DecryptToHtml
Node.js script that decompresses .docx, deletes some files and then converts it back again to docx and finally converts it to html with it's settings.
Nodejs installation required.
To use this you should either have an archive with multiple or 1 .docx inside.
Usage: 

node.exe Converter.js
It will promt you to set the archive's name, you should set it like this - archive.zip.
After that it will extract the archive into the folder docx that it creates if it does not exist and then will also create a temp_docx folder where temporary files will be left. After the work it will delete the files in the temp_docx folder and will exit.

