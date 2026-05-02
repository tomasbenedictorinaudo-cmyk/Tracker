#!/usr/bin/env node
// Bundles index.html + styles.css + app.js + favicon.svg into a single
// standalone dist/cockpit.html. No runtime dependencies — open the
// bundled file with file:// or double-click; everything works offline.
//
// Usage:   node build.js
// Output:  dist/cockpit.html
//
// Run after every change to index.html / styles.css / app.js / favicon.svg.
'use strict';
const fs = require('fs');
const path = require('path');

const ROOT  = __dirname;
const SRC   = (name) => path.join(ROOT, name);
const OUT   = path.join(ROOT, 'dist', 'cockpit.html');

function read(name) {
  return fs.readFileSync(SRC(name), 'utf8');
}

function bytes(name) {
  return fs.readFileSync(SRC(name));
}

function fmtBytes(n) {
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1024 / 1024).toFixed(2) + ' MB';
}

function bundle() {
  const html  = read('index.html');
  const css   = read('styles.css');
  const js    = read('app.js');
  const fav   = bytes('favicon.svg').toString('base64');
  const stamp = new Date().toISOString();

  // Sanity guards — inlining `</script>` or `</style>` literals would close
  // the host tag prematurely. The current sources are clean, but bail out
  // loudly if a future edit introduces one.
  if (js.includes('</script>')) {
    throw new Error('app.js contains a literal </script> — escape it before inlining.');
  }
  if (css.includes('</style>')) {
    throw new Error('styles.css contains a literal </style> — escape it before inlining.');
  }

  let out = html;

  // String.prototype.replace treats `$$`, `$&`, `$\``, `$'`, `$<name>`, and
  // `$1`..`$9` as substitution patterns inside the replacement string. The
  // app's source code uses `$$` as a real identifier (querySelectorAll
  // helper), so we MUST pass the replacement as a function — function
  // returns are not subject to that substitution.
  const literal = (s) => () => s;

  // 1. Replace the stylesheet <link> with an inline <style>.
  const linkRE = /<link\s+rel=["']stylesheet["']\s+href=["']styles\.css["']\s*\/?>/i;
  if (!linkRE.test(out)) throw new Error('Could not find <link rel="stylesheet" href="styles.css">');
  out = out.replace(linkRE, literal(`<style>\n${css}\n  </style>`));

  // 2. Replace the favicon <link> with a data URI so the bundle stays
  //    standalone (otherwise file:// would 404 the favicon).
  const favRE = /<link\s+rel=["']icon["'][^>]*href=["']favicon\.svg["'][^>]*\/?>/i;
  if (favRE.test(out)) {
    out = out.replace(favRE, literal(`<link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,${fav}" />`));
  }

  // 3. Replace the <script src="app.js"> with an inline <script>.
  const scriptRE = /<script\s+src=["']app\.js["']\s*><\/script>/i;
  if (!scriptRE.test(out)) throw new Error('Could not find <script src="app.js">');
  out = out.replace(scriptRE, literal(`<script>\n${js}\n  </script>`));

  // 4. Stamp a build header so the file's provenance is obvious.
  out = out.replace(
    /<head>/i,
    `<head>\n  <!-- Built ${stamp} by build.js — single-file standalone bundle -->`,
  );

  // Round-trip integrity check: extract the inlined <script> and <style>
  // and confirm every byte of the source survived. Catches silent damage
  // from replacement-string substitutions ($$, $&, etc.) or stray edits to
  // build.js itself, before the broken bundle gets shipped.
  const inlinedScript = (out.match(/<script>\n([\s\S]*?)\n  <\/script>/) || [])[1];
  if (inlinedScript !== js) {
    throw new Error('Inlined <script> body does not match app.js byte-for-byte (build corruption).');
  }
  const inlinedStyle = (out.match(/<style>\n([\s\S]*?)\n  <\/style>/) || [])[1];
  if (inlinedStyle !== css) {
    throw new Error('Inlined <style> body does not match styles.css byte-for-byte (build corruption).');
  }

  // Ensure dist/ exists, write atomically-ish.
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT + '.tmp', out);
  fs.renameSync(OUT + '.tmp', OUT);

  const cssBytes = Buffer.byteLength(css, 'utf8');
  const jsBytes  = Buffer.byteLength(js, 'utf8');
  const totalBytes = Buffer.byteLength(out, 'utf8');
  console.log(`✓ dist/cockpit.html  (${fmtBytes(totalBytes)} total — css ${fmtBytes(cssBytes)}, js ${fmtBytes(jsBytes)}, built ${stamp})`);
}

try {
  bundle();
} catch (err) {
  console.error('✗ build failed:', err.message);
  process.exitCode = 1;
}
