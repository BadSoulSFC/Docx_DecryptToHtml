const fs = require('fs');
const AdmZip = require('adm-zip');
const mammoth = require('mammoth');
const folderDocx = 'docx_decrypted/';
const folderHtml = 'Html/';

async function DocxExtraction(){
  const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
  });
  readline.question("Enter the name of the zip archive: ", (zipFileName) => {readline.close();
    // Extracting files to a directory   
    //const zip = new AdmZip('example.zip');
    const zip = new AdmZip(zipFileName);
    zip.extractAllTo(folderDocx, /*overwrite*/true);
  
    // Process all docx files in the folder
    const files = fs.readdirSync(folderDocx);
    files.forEach(file => {if (file.endsWith('.docx')) {
      console.log(`Processing file: ${file}`);
      processDocx(folderDocx + file);
      console.log(`Finished processing ${file}`);
    }
  });
  }); 
}

async function processDocx(file) {
  // Extract docx to a temporary folder
  const zip = new AdmZip(file);
  const tempFolder = 'temp_';
  zip.extractAllTo(tempFolder, true);
  
  // Remove w:documentProtection from settings.xml
  const settingsXml = tempFolder + '/word/settings.xml';
  let settings = fs.readFileSync(settingsXml, 'utf8');
  settings = settings.replace(/<w:documentProtection[\s\S]*?<\/w:documentProtection>/g, '');
  fs.writeFileSync(settingsXml, settings);
  
  // Remove all w:p nodes that contain w:pict from document.xml
  const documentXml = tempFolder + '/word/document.xml';
  let document = fs.readFileSync(documentXml, 'utf8');
  document = document.replace(/<w:p[\s\S]*?<w:pict[\s\S]*?<\/w:p>/g, '');
  fs.writeFileSync(documentXml, document);

  // Repack and save as decrypted docx
  zip.addLocalFile(settingsXml);
  zip.addLocalFile(documentXml);
  zip.writeZip(folderDocx.replace('/', '')) + file.replace(folderDocx, '');
  //zip.writeZip(folderDocx.replace('/', '')) + file.replace(folderDocx, '') + file.replace('.docx', '_decrypted.docx');

  // Convert to HTML and apply custom styles
  const result = await mammoth.convertToHtml({
    path: folderDocx + file.replace(folderDocx, ''),
    styleMap: [
      {match: 'p[style-name="Заголовок 1"]', tag: 'h2'},
      {match: 'p[style-name="Заголовок 2"]', tag: 'h3'},
      {match: 'span[style-name="consolas"]', tag: 'pre'}]});

  // Save HTML to html folder
  fs.writeFileSync(folderHtml + file.replace(folderDocx, '').replace('.docx', '') + '.html', result.value);
  fs.rmSync(tempFolder, {recursive: true})
  if(fs.existsSync(tempFolder)){fs.rmdirSync(tempFolder, {recursive: true});}
 
}

//Theoretical int main(){}

if(!fs.existsSync(folderDocx)){fs.mkdirSync(folderDocx);}
if(!fs.existsSync(folderHtml)){fs.mkdirSync(folderHtml);}

DocxExtraction()