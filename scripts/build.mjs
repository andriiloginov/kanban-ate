// Збирає Tailwind CSS і вбудовує його в index.html (замість <!-- tailwind -->).
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';

const bin = new URL('../node_modules/.bin/tailwindcss', import.meta.url).pathname;
const css = execFileSync(bin, ['-i', 'src/input.css', '--minify'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit'] })
  .replace(/url\(["']?[^"')]*?fonts\/([\w-]+\.woff2)["']?\)/g, (_, f) => `url(data:font/woff2;base64,${readFileSync(`fonts/${f}`).toString('base64')})`);
const src = readFileSync('src/index.src.html', 'utf8');
if (!src.includes('<!-- tailwind -->')) throw new Error('У src/index.src.html немає <!-- tailwind -->');
writeFileSync('index.html', src.replace('<!-- tailwind -->', () => `<style>${css.trim()}</style>`));
console.log(`index.html зібрано (${(css.length / 1024).toFixed(1)} KB CSS)`);
