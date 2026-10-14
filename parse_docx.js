const fs = require('fs');
const xml = fs.readFileSync('D:/Guru/docx_raw.xml', 'utf8');

// Replace </w:p> with newline, strip XML tags
const text = xml
  .replace(/<\/w:p>/g, '\n')
  .replace(/<[^>]+>/g, '')
  .replace(/&amp;/g, '&')
  .replace(/&lt;/g, '<')
  .replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'");

fs.writeFileSync('D:/Guru/extracted_doc_text.txt', text, 'utf8');
console.log('Finished. Sample text:');
console.log(text.substring(0, 2000));
