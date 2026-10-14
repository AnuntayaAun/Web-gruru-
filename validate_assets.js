const fs = require('fs');
const path = require('path');
const html = fs.readFileSync('D:/Guru/presentation.html', 'utf8');

const regex = /src=["']([^"']+)["']/g;
let match;
let missing = 0;
console.log('Validating asset references:');
while ((match = regex.exec(html)) !== null) {
  const src = match[1];
  const absPath = path.resolve('D:/Guru', src);
  const exists = fs.existsSync(absPath);
  console.log(` - ${src} : ${exists ? 'OK' : 'MISSING'}`);
  if (!exists) missing++;
}

if (missing === 0) {
  console.log('SUCCESS: All image and asset paths in presentation.html are valid!');
} else {
  console.error(`ERROR: Found ${missing} missing assets.`);
}
