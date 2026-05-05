import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { resolve } from 'node:path';

const repoRoot = resolve(process.cwd());
const packageJsonPath = resolve(repoRoot, 'package.json');

const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8'));
const packageVersion = String(packageJson.version || '').trim();

if (!packageVersion) {
  console.error('Keine gueltige Version in package.json gefunden.');
  process.exit(1);
}

const manifestVersion = `${packageVersion}.0`;
const manifestFiles = ['manifest.prod.xml', 'manifest.dev.xml', 'manifest.xml'];

for (const fileName of manifestFiles) {
  const filePath = resolve(repoRoot, fileName);
  if (!existsSync(filePath)) {
    continue;
  }

  const content = readFileSync(filePath, 'utf8');
  if (!content.includes('<Version>')) {
    console.error(`Keine <Version>-Angabe gefunden in ${fileName}.`);
    process.exit(1);
  }

  const updated = content.replace(/<Version>[^<]+<\/Version>/, `<Version>${manifestVersion}</Version>`);

  if (updated !== content) {
    writeFileSync(filePath, updated, 'utf8');
    console.log(`${fileName} -> ${manifestVersion}`);
  } else {
    console.log(`${fileName} bereits auf ${manifestVersion}`);
  }
}