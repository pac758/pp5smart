const fs = require('fs');
const path = require('path');

function walk(dir, exts) {
  const out = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const e of entries) {
    if (e.name === '.git' || e.name === 'node_modules') continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...walk(p, exts));
    else if (exts.has(path.extname(e.name).toLowerCase())) out.push(p);
  }
  return out;
}

function extractRunCalls(html) {
  const calls = [];
  // Extract terminal GAS function calls on google.script.run chains.
  // Key detail: ignore any `.something(` inside callback bodies by only collecting
  // chained method names that appear at top-level of the chain (parenDepth === 0).

  const runRe = /google\s*\.\s*script\s*\.\s*run\b/gi;

  function findStatementEnd(src, startIdx, maxLen) {
    const endLimit = Math.min(src.length, startIdx + (maxLen || 20000));
    let paren = 0;
    let bracket = 0;
    let brace = 0;
    let inS = false,
      inD = false,
      inT = false;
    let inLine = false,
      inBlock = false;

    for (let i = startIdx; i < endLimit; i++) {
      const c = src[i];
      const n = src[i + 1];

      if (inLine) {
        if (c === '\n') inLine = false;
        continue;
      }
      if (inBlock) {
        if (c === '*' && n === '/') {
          inBlock = false;
          i++;
        }
        continue;
      }

      if (!inS && !inD && !inT) {
        if (c === '/' && n === '/') {
          inLine = true;
          i++;
          continue;
        }
        if (c === '/' && n === '*') {
          inBlock = true;
          i++;
          continue;
        }
      }

      if (!inD && !inT && c === '\'' && src[i - 1] !== '\\') {
        inS = !inS;
        continue;
      }
      if (!inS && !inT && c === '"' && src[i - 1] !== '\\') {
        inD = !inD;
        continue;
      }
      if (!inS && !inD && c === '`' && src[i - 1] !== '\\') {
        inT = !inT;
        continue;
      }
      if (inS || inD || inT) continue;

      if (c === '(') paren++;
      else if (c === ')') paren = Math.max(0, paren - 1);
      else if (c === '[') bracket++;
      else if (c === ']') bracket = Math.max(0, bracket - 1);
      else if (c === '{') brace++;
      else if (c === '}') brace = Math.max(0, brace - 1);

      // End statement only when we're at top-level
      if (paren === 0 && bracket === 0 && brace === 0 && c === ';') {
        return i;
      }
    }

    return endLimit - 1;
  }

  function scanTopLevelMethods(chunk) {
    const methods = [];
    let paren = 0;
    let bracket = 0;
    let brace = 0;
    let inS = false,
      inD = false,
      inT = false;
    let inLine = false,
      inBlock = false;

    const isIdStart = (c) => /[A-Za-z_$]/.test(c);
    const isId = (c) => /[A-Za-z0-9_$]/.test(c);

    for (let i = 0; i < chunk.length; i++) {
      const c = chunk[i];
      const n = chunk[i + 1];

      if (inLine) {
        if (c === '\n') inLine = false;
        continue;
      }
      if (inBlock) {
        if (c === '*' && n === '/') {
          inBlock = false;
          i++;
        }
        continue;
      }

      if (!inS && !inD && !inT) {
        if (c === '/' && n === '/') {
          inLine = true;
          i++;
          continue;
        }
        if (c === '/' && n === '*') {
          inBlock = true;
          i++;
          continue;
        }
      }

      // string toggles
      if (!inD && !inT && c === '\'' && chunk[i - 1] !== '\\') {
        inS = !inS;
        continue;
      }
      if (!inS && !inT && c === '"' && chunk[i - 1] !== '\\') {
        inD = !inD;
        continue;
      }
      if (!inS && !inD && c === '`' && chunk[i - 1] !== '\\') {
        inT = !inT;
        continue;
      }
      if (inS || inD || inT) continue;

      // depth tracking
      if (c === '(') paren++;
      else if (c === ')') paren = Math.max(0, paren - 1);
      else if (c === '[') bracket++;
      else if (c === ']') bracket = Math.max(0, bracket - 1);
      else if (c === '{') brace++;
      else if (c === '}') brace = Math.max(0, brace - 1);

      // Only consider method chaining when we're at top-level of the chain.
      // (Outside any () argument list)
      if (paren !== 0 || bracket !== 0 || brace !== 0) continue;

      if (c === '.') {
        // skip whitespace
        let j = i + 1;
        while (j < chunk.length && /\s/.test(chunk[j])) j++;
        if (j >= chunk.length || !isIdStart(chunk[j])) continue;

        let k = j + 1;
        while (k < chunk.length && isId(chunk[k])) k++;
        const name = chunk.slice(j, k);

        // Ensure it's a call `.name(` at top-level
        let p = k;
        while (p < chunk.length && /\s/.test(chunk[p])) p++;
        if (chunk[p] === '(') methods.push(name);
      }

      // Support bracket-notation calls at top-level: run[fnName](...)
      if (c === '[') {
        let j = i + 1;
        while (j < chunk.length && /\s/.test(chunk[j])) j++;
        if (j >= chunk.length) continue;

        let name = '';
        if (chunk[j] === '\'' || chunk[j] === '"') {
          const q = chunk[j];
          j++;
          let k = j;
          while (k < chunk.length && chunk[k] !== q) k++;
          name = chunk.slice(j, k);
          j = k + 1;
        } else {
          // dynamic [fnName] can't be resolved statically; skip
          continue;
        }

        // find closing ]
        while (j < chunk.length && chunk[j] !== ']') j++;
        if (chunk[j] !== ']') continue;
        j++;
        while (j < chunk.length && /\s/.test(chunk[j])) j++;
        if (chunk[j] === '(' && name) methods.push(name);
      }
    }
    return methods;
  }

  let m;
  while ((m = runRe.exec(html))) {
    const start = m.index;
    const end = findStatementEnd(html, start, 30000);
    const chunk = html.slice(start, end + 1);

    const chainMethods = scanTopLevelMethods(chunk);
    const skip = new Set([
      'withSuccessHandler',
      'withFailureHandler',
      'withUserObject',
      'withFinalizer',
      // common non-GAS methods that can be accidentally captured if a statement isn't terminated cleanly
      'addEventListener',
      'removeEventListener',
      'apply',
      'call',
      'bind'
    ]);

    const terminal = [...chainMethods].reverse().find((fn) => !skip.has(fn));

    if (terminal) calls.push(terminal);
  }

  return [...new Set(calls)];
}

function indexJsFunctions(jsText) {
  const idx = new Map();
  const patterns = [
    // async function foo(...) {
    /(^|\n)\s*async\s+function\s+([A-Za-z_$][\w$]*)\s*\(/g,
    // function foo(...) {
    /(^|\n)\s*function\s+([A-Za-z_$][\w$]*)\s*\(/g,
    // const foo = function(...) {
    /(^|\n)\s*(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*function\s*\(/g,
    // const foo = (...) => {
    /(^|\n)\s*(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?\([^)]*\)\s*=>\s*\{/g,
    // const foo = x => {
    /(^|\n)\s*(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?[A-Za-z_$][\w$]*\s*=>\s*\{/g,
    // foo: function(...) {
    /(^|\n)\s*([A-Za-z_$][\w$]*)\s*:\s*function\s*\(/g,
    // foo: (...) => {
    /(^|\n)\s*([A-Za-z_$][\w$]*)\s*:\s*(?:async\s*)?\([^)]*\)\s*=>\s*\{/g
  ];

  for (const re of patterns) {
    let m;
    while ((m = re.exec(jsText))) {
      const name = m[2];
      const pos = m.index + (m[1] ? 1 : 0);
      if (!idx.has(name)) idx.set(name, pos);
    }
  }
  return idx;
}

function sliceFunctionBody(jsText, startIdx) {
  const braceIdx = jsText.indexOf('{', startIdx);
  if (braceIdx < 0) return '';
  let i = braceIdx;
  let depth = 0;
  let inS = false, inD = false, inT = false;
  let inLine = false, inBlock = false;
  for (; i < jsText.length; i++) {
    const c = jsText[i];
    const n = jsText[i + 1];

    if (inLine) {
      if (c === '\n') inLine = false;
      continue;
    }
    if (inBlock) {
      if (c === '*' && n === '/') { inBlock = false; i++; }
      continue;
    }

    if (!inS && !inD && !inT) {
      if (c === '/' && n === '/') { inLine = true; i++; continue; }
      if (c === '/' && n === '*') { inBlock = true; i++; continue; }
    }

    if (!inD && !inT && c === '\'' && jsText[i - 1] !== '\\') { inS = !inS; continue; }
    if (!inS && !inT && c === '"' && jsText[i - 1] !== '\\') { inD = !inD; continue; }
    if (!inS && !inD && c === '`' && jsText[i - 1] !== '\\') { inT = !inT; continue; }
    if (inS || inD || inT) continue;

    if (c === '{') depth++;
    else if (c === '}') {
      depth--;
      if (depth === 0) return jsText.slice(braceIdx, i + 1);
    }
  }
  return jsText.slice(braceIdx);
}

function detectSheets(fnBody) {
  const sheets = new Set();
  const patterns = [
    /getSheetByName\(\s*['"]([^'"]+)['"]\s*\)/g,
    /insertSheet\(\s*['"]([^'"]+)['"]\s*\)/g,
    /deleteSheet\(\s*['"]([^'"]+)['"]\s*\)/g,
    /_readSheetToObjects\(\s*['"]([^'"]+)['"]\s*\)/g,
    /_writeObjectsToSheet\(\s*['"]([^'"]+)['"]\s*[,)]/g
  ];
  for (const re of patterns) {
    let m;
    while ((m = re.exec(fnBody))) sheets.add(m[1]);
  }
  return [...sheets];
}

function detectAuth(fnName, fnBody) {
  const name = fnName.toLowerCase();
  if (name.includes('login') || name.includes('token') || name.includes('passwordreset')) return 'public';
  if (/verifyAuthToken\s*\(|validateToken\s*\(|getActiveSession\s*\(|isUserLoggedIn\s*\(/.test(fnBody)) return 'token_checked';
  return 'unknown';
}

function toCsv(rows) {
  const headers = ['page','function','serverFile','auth','sheets'];
  const esc = (v) => {
    const s = String(v ?? '');
    if (/[\n",]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  };
  const lines = [headers.join(',')];
  for (const r of rows) {
    lines.push(headers.map(h => esc(r[h])).join(','));
  }
  return lines.join('\n');
}

function main() {
  const root = process.argv[2] ? path.resolve(process.argv[2]) : process.cwd();
  const htmlFiles = walk(root, new Set(['.html']));
  const jsFiles = walk(root, new Set(['.js']));

  const jsIndex = new Map();
  for (const f of jsFiles) {
    const txt = fs.readFileSync(f, 'utf8');
    const fns = indexJsFunctions(txt);
    for (const [fn, pos] of fns) {
      if (!jsIndex.has(fn)) jsIndex.set(fn, []);
      jsIndex.get(fn).push({ file: f, pos, txt });
    }
  }

  const rows = [];
  for (const hf of htmlFiles) {
    const html = fs.readFileSync(hf, 'utf8');
    const calls = extractRunCalls(html);
    const page = path.relative(root, hf).replace(/\\/g, '/');

    for (const fn of calls) {
      const defs = jsIndex.get(fn) || [];
      if (defs.length === 0) {
        rows.push({ page, function: fn, serverFile: '', auth: 'unknown', sheets: '' });
        continue;
      }
      for (const d of defs) {
        const body = sliceFunctionBody(d.txt, d.pos);
        const sheets = detectSheets(body);
        const auth = detectAuth(fn, body);
        rows.push({
          page,
          function: fn,
          serverFile: path.relative(root, d.file).replace(/\\/g, '/'),
          auth,
          sheets: sheets.join('|')
        });
      }
    }
  }

  rows.sort((a, b) => (a.page + '::' + a.function).localeCompare(b.page + '::' + b.function));

  fs.writeFileSync(path.join(root, 'api_inventory.json'), JSON.stringify({ root, generatedAt: new Date().toISOString(), rows }, null, 2), 'utf8');
  fs.writeFileSync(path.join(root, 'api_inventory.csv'), toCsv(rows), 'utf8');

  const md = [];
  md.push('| page | function | serverFile | auth | sheets |');
  md.push('|---|---|---|---|---|');
  for (const r of rows) {
    md.push(`| ${r.page} | ${r.function} | ${r.serverFile || ''} | ${r.auth} | ${r.sheets || ''} |`);
  }
  fs.writeFileSync(path.join(root, 'api_inventory.md'), md.join('\n'), 'utf8');

  const missing = rows.filter(r => !r.serverFile).length;
  console.log(`Done. HTML files: ${htmlFiles.length}, JS files: ${jsFiles.length}, rows: ${rows.length}, missing server mapping: ${missing}`);
  console.log('Wrote: api_inventory.json, api_inventory.csv, api_inventory.md');
}

main();
