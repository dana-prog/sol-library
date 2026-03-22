const fs = require('fs');
const path = require('path');

const typesDir = './types-tmp';
const output = './types/sol-library.d.ts';

const files = fs.readdirSync(typesDir)
  .filter(f => f.endsWith('.d.ts') && f !== 'sol-library.d.ts');

let content = files
  .map(f => fs.readFileSync(path.join(typesDir, f), 'utf8'))
  .join('\n')
  // remove "declare function"
  .replace(/declare function/g, '')
  // convert variables/constants
  .replace(/declare (const|let|var) (\w+): ([^;]+);/g,
    (_, kind, name, type) => `  ${name}: ${type};`
  )
  // remove extra "declare"
  .replace(/declare /g, '')
  // add newlines after semicolons
  .replace(/;\n\n/g, ';\n')
  .replace(/;\n/g, ';\n\n');

const wrapped = `declare const SOLLibrary: {
${_removePrivateFunctions(content)}
};\n`;

fs.writeFileSync(output, wrapped);

function _removePrivateFunctions(content) {
  const lines = content.split('\n');

  let skip = false;
  let inJSDoc = false;
  let buffer = [];

  const result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // detect start of JSDoc
    if (line.trim().startsWith('/**')) {
      inJSDoc = true;
      buffer = [line];
      continue;
    }

    if (inJSDoc) {
      buffer.push(line);

      // end of JSDoc
      if (line.includes('*/')) {
        inJSDoc = false;

        const isPrivate = buffer.some(l => l.includes('@private'));

        if (isPrivate) {
          skip = true; // skip next declaration
          buffer = [];
          continue;
        } else {
          result.push(...buffer);
          buffer = [];
        }
      }

      continue;
    }

    // skip function after @private
    if (skip) {
      if (line.includes(';')) {
        skip = false;
      }
      continue;
    }

    result.push(line);
  }

  return result.join('\n');
}