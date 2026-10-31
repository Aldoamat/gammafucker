#!/usr/bin/env node
// ═══════════════════════════════════════════════════════════════════
//  GammaFucker — script di rilascio
//
//  node release.mjs build [N]   → bump versione (v26→v27 o vN), sincronizza
//                                 getScript() da apps-script/Codice.js,
//                                 prepara dist/ per Netlify. Nessun deploy.
//  node release.mjs deploy      → git commit + push (Netlify pubblica dal repo),
//                                 clasp push + deploy. Solo dopo il build.
//                                 --netlify: forza netlify deploy via CLI; --clasp: forza clasp da qui
//                                 (di norma lo script lo deploya la GitHub Action dopo il push).
//  node release.mjs status      → mostra versioni correnti e coerenza.
//
//  Fonte di verità del backend: apps-script/Codice.js
//  Il testo dentro getScript() in index.html viene rigenerato da lì.
// ═══════════════════════════════════════════════════════════════════
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const P = {
  index: path.join(ROOT, 'index.html'),
  sw: path.join(ROOT, 'sw.js'),
  script: path.join(ROOT, 'apps-script', 'Codice.js'),
  cfg: path.join(ROOT, 'release.config.json'),
  dist: path.join(ROOT, 'dist'),
  distCollab: path.join(ROOT, 'dist-collab'),
};
const PUBLISH = ['index.html', 'sw.js', 'manifest.json', 'icon-192.png', 'icon-512.png'];

const read = f => fs.readFileSync(f, 'utf8');
const write = (f, s) => fs.writeFileSync(f, s, 'utf8');
const sh = (cmd, cwd = ROOT) => { console.log('  $ ' + cmd); execSync(cmd, { cwd, stdio: 'inherit', shell: true }); };
const has = cmd => { try { execSync(`${cmd} --version`, { stdio: 'ignore', shell: true }); return true; } catch { return false; } };

// ── versioni ──────────────────────────────────────────────────────
function versions() {
  const idx = read(P.index), sw = read(P.sw), scr = read(P.script);
  const app = [...idx.matchAll(/GammaFucker v(\d+)/g)].map(m => +m[1]);
  const hdr = idx.match(/font-weight:normal;">v(\d+)</);
  const cache = sw.match(/const CACHE='teatro-v([0-9a-z]+)'/);
  const emb = idx.match(/SPETTACOLI APP SCRIPT v(\d+)/);
  const script = scr.match(/SPETTACOLI APP SCRIPT v(\d+)/);
  return {
    app: app.length ? Math.max(...app) : null,
    appAll: app, header: hdr ? +hdr[1] : null,
    cache: cache ? cache[1] : null,
    scriptEmbedded: emb ? +emb[1] : null,
    script: script ? +script[1] : null,
  };
}

function status() {
  const v = versions();
  console.log('Versioni correnti');
  console.log(`  app (footer/report): v${v.app}   header: v${v.header}   sw cache: teatro-v${v.cache}`);
  console.log(`  Apps Script: Codice.js v${v.script}`);
  const ok = v.app === v.header && String(v.app) === v.cache && !!v.script;
  console.log(ok ? '  ✓ coerente' : '  ✗ versioni non allineate — lancia "node release.mjs build"');
  const cfg = fs.existsSync(P.cfg) ? JSON.parse(read(P.cfg)) : {};
  console.log(`  deploymentId Apps Script: ${cfg.deploymentId || '(mancante — vedi release.config.json)'}`);
  return ok;
}

// ── build ─────────────────────────────────────────────────────────
function embed(scriptSrc) {
  // Codice.js → template literal JS (stesso escaping che si farebbe a mano)
  return scriptSrc.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$\{/g, '\\${');
}

function build(target) {
  const v = versions();
  const next = target ? +target : (v.app || 0) + 1;
  const scriptNext = (v.script || 0) + 1;
  console.log(`Build: app v${v.app} → v${next}, Apps Script v${v.script} → v${scriptNext}`);

  let idx = read(P.index);
  idx = idx.replace(/GammaFucker v\d+/g, `GammaFucker v${next}`);
  idx = idx.replace(/font-weight:normal;">v\d+</, `font-weight:normal;">v${next}<`);

  let scr = read(P.script);
  scr = scr.replace(/SPETTACOLI APP SCRIPT v\d+/, `SPETTACOLI APP SCRIPT v${scriptNext}`);
  write(P.script, scr);

  // dalla v108 lo script non è più incluso in index.html (vive solo in apps-script/Codice.js)
  write(P.index, idx);

  let sw = read(P.sw);
  sw = sw.replace(/const CACHE='teatro-v[0-9a-z]+'/, `const CACHE='teatro-v${next}'`);
  write(P.sw, sw);

  fs.rmSync(P.dist, { recursive: true, force: true });
  fs.mkdirSync(P.dist);
  for (const f of PUBLISH) fs.copyFileSync(path.join(ROOT, f), path.join(P.dist, f));

  // App collaboratori: stesso sorgente, APP_MODE='collab', cache e manifest propri → dist-collab/
  fs.rmSync(P.distCollab, { recursive: true, force: true });
  fs.mkdirSync(P.distCollab);
  const marker = "const APP_MODE = 'admin'; // __APP_MODE__";
  if (!idx.includes(marker)) throw new Error('marker APP_MODE non trovato in index.html');
  write(path.join(P.distCollab, 'index.html'), idx.replace(marker, "const APP_MODE = 'collab'; // __APP_MODE__").replace(/<title>[^<]*<\/title>/, '<title>GammaFucker Collab</title>'));
  write(path.join(P.distCollab, 'sw.js'), sw.replace(/const CACHE='teatro-v[0-9a-z]+'/, `const CACHE='collab-v${next}'`));
  const man = JSON.parse(read(path.join(ROOT, 'manifest.json')));
  man.name = 'GammaFucker Collab'; man.short_name = 'GF Collab';
  write(path.join(P.distCollab, 'manifest.json'), JSON.stringify(man, null, 2));
  for (const f of ['icon-192.png', 'icon-512.png']) fs.copyFileSync(path.join(ROOT, f), path.join(P.distCollab, f));

  console.log(`✓ index.html, sw.js, Codice.js aggiornati; dist/ e dist-collab/ pronte`);
  console.log('  Controlla il diff, poi: node release.mjs deploy');
}

// ── deploy ────────────────────────────────────────────────────────
function deploy() {
  if (!status()) process.exit(1);
  if (!fs.existsSync(P.dist)) { console.error('dist/ mancante: lancia prima "node release.mjs build"'); process.exit(1); }
  const v = versions();
  const cfg = fs.existsSync(P.cfg) ? JSON.parse(read(P.cfg)) : {};
  const tag = `v${v.app}`;

  console.log('\n[1/3] git');
  let pushed = false;
  if (!has('git')) console.log('  git non installato — salto (winget install Git.Git)');
  else {
    if (!fs.existsSync(path.join(ROOT, '.git'))) sh('git init');
    try { execSync('git rm -r -q --cached --ignore-unmatch "Claude outputs"', { cwd: ROOT, stdio: 'ignore', shell: true }); } catch {}
    sh('git add -A');
    try { sh(`git commit -m "release ${tag}"`); } catch { console.log('  (niente da committare)'); }
    try { sh(`git tag -f ${tag}`); } catch {}
    let remote = '';
    try { remote = execSync('git remote get-url origin', { cwd: ROOT, stdio: ['ignore', 'pipe', 'ignore'], shell: true }).toString().trim(); } catch {}
    if (remote) {
      try { sh('git pull --rebase --autostash origin main'); } catch { console.log('  (pull non riuscito: continuo con il push)'); }
      sh('git push origin HEAD');
      try { sh(`git push -f origin refs/tags/${tag}`); } catch { console.log('  (tag non aggiornato, non è un problema)'); }
      pushed = true; console.log('  → push su ' + remote + ' (Netlify pubblica l\'app, la GitHub Action lo script)');
    }
  }

  console.log('\n[2/3] Apps Script');
  const hasAction = fs.existsSync(path.join(ROOT, '.github', 'workflows', 'deploy-apps-script.yml'));
  if (pushed && hasAction && !process.argv.includes('--clasp')) {
    console.log('  saltato: lo script viene deployato dalla GitHub Action dopo il push (usa "deploy --clasp" per farlo da qui)');
  } else {
  sh('clasp push -f', path.join(ROOT, 'apps-script'));
  if (cfg.deploymentId) sh(`clasp deploy -i ${cfg.deploymentId} -d "${tag}"`, path.join(ROOT, 'apps-script'));
  else {
    console.log('  ⚠ deploymentId mancante in release.config.json: codice caricato ma NON ridistribuito.');
    console.log('    Recuperalo con: cd apps-script && clasp deployments   (quello con l\'URL /exec in uso)');
  }
  }

  console.log('\n[3/3] Netlify');
  if (pushed && !process.argv.includes('--netlify')) console.log('  saltato: il sito è collegato a GitHub e si aggiorna dal push (usa "deploy --netlify" per forzare la CLI)');
  else sh(`netlify deploy --prod --dir=dist --message "${tag}"`);

  console.log(`\n✓ Rilascio ${tag} completato`);
}

// ── main ──────────────────────────────────────────────────────────
const [cmd, arg] = process.argv.slice(2);
try {
  if (cmd === 'build') build(arg);
  else if (cmd === 'deploy') deploy();
  else if (cmd === 'status') status();
  else { console.log('uso: node release.mjs build [N] | deploy | status'); process.exit(1); }
} catch (e) { console.error('✗ ' + e.message); process.exit(1); }
