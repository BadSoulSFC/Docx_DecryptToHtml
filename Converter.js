const fs = require('fs');
const AdmZip = require('adm-zip');
const mammoth = require('mammoth');

async function processDocx(file) {
  // Extract docx to a temporary folder
  const zip = new AdmZip(file);
  const tempFolder = 'temp_' + file;
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
  //const tempArchFolder = 'decrypt/';
  // Repack and save as decrypted docx
  zip.addLocalFile(settingsXml);
  zip.addLocalFile(documentXml);
  zip.writeZip('docx') + file.replace('docx', '');
  //zip.writeZip('tempArchFolder' + file.replace('.docx', '_decrypted.docx'));

  // Convert to HTML and apply custom styles
  const result = await mammoth.convertToHtml({
    path: 'docx' + file.replace('docx', ''),
    styleMap: [
      {
        match: 'p[style-name="Заголовок 1"]',
        tag: 'h2'
      },
      {
        match: 'p[style-name="Заголовок 2"]',
        tag: 'h3'
      },
      {
        match: 'span[style-name="consolas"]',
        tag: 'pre'
      }
    ]
  });
  // Save HTML to html folder
  fs.writeFileSync('html/' + file.replace('.docx', '').replace('docx/', '') + '.html', result.value);
  fs.rmSync(tempFolder, {recursive: true});
    //fs.rmSync('decrypt/docx_decrypted', {recursive: true});
  
}

if(!fs.existsSync('docx/')){fs.mkdirSync('docx/');}
if(!fs.existsSync('html/')){fs.mkdirSync('html/');}
if(!fs.existsSync('decrypt/')){fs.mkdirSync('decrypt/');}
// Uncomment this if userinput is required.
/*
  const readline = require('readline').createInterface({
  input: process.stdin,
  output: process.stdout
});
readline.question("Enter the name of the zip archive: ", (zipFileName) => {
  readline.close();
  const zip = new AdmZip(zipFileName);
  // Also add }); to the end of the script
  */

  const zip = new AdmZip('docx.zip') 

  // Extracting files to a directory
  // You can create a global variable
  // with the path in it and change the code a little.
  // const FolderPath = 'SamplePath/';
  zip.extractAllTo("docx/", /*overwrite*/true);

  // Process all docx files in the folder
  const folder = 'docx/';
  const files = fs.readdirSync(folder);
  files.forEach(file => {
    if (file.endsWith('.docx')) {
      processDocx(folder + file);
    }});
  

    