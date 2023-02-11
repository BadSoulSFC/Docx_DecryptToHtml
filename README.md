# Docx_DecryptToHtml
Node.js script that extracts .docx, deletes some parameters and then converts it back again to docx and finally converts it to html in a different folder with it's settings.
Nodejs installation required.
To use this you should either have an archive with multiple or 1 .docx inside.
You can also modify the code so it searches for the archive by itself by deleting userinput
Usage: node.exe Converter.js

It will promt you to set the archive's name, you should set it like this - archive.zip, do this only if the archive is in the same directory, if not, you should set the path manually like this - C:\Path\archive.zip.
After that it will extract the archive into the folder docx that it creates if it does not exist and then will also create a temp_ folder where temporary files will be before deletion.

