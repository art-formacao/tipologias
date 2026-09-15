import fs from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';

const root = process.cwd();
const areas = [
  { directory: 'Tipo de Reboque', kind: 'Reboque' },
  { directory: 'Tipo de Rígido', kind: 'Rígido' }
];
const imageExtensions = new Set(['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif']);
const natural = new Intl.Collator('pt', { numeric: true, sensitivity: 'base' });

function normalize(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function toWebPath(absolutePath) {
  return './' + path.relative(root, absolutePath).split(path.sep).join('/');
}

async function directoriesAt(directory) {
  try {
    const entries = await fs.readdir(directory, { withFileTypes: true });
    return entries.filter(entry => entry.isDirectory()).map(entry => ({
      name: entry.name,
      absolute: path.join(directory, entry.name)
    }));
  } catch {
    return [];
  }
}

async function filesAt(directory) {
  try {
    const entries = await fs.readdir(directory, { withFileTypes: true });
    return entries.filter(entry => entry.isFile()).map(entry => ({
      name: entry.name,
      absolute: path.join(directory, entry.name)
    }));
  } catch {
    return [];
  }
}

async function findNamedDirectory(directory, expectedName) {
  const expected = normalize(expectedName);
  return (await directoriesAt(directory)).find(item => normalize(item.name) === expected);
}

async function readImages(directory) {
  if (!directory) return [];
  return (await filesAt(directory.absolute))
    .filter(file => imageExtensions.has(path.extname(file.name).toLowerCase()))
    .sort((a, b) => natural.compare(a.name, b.name));
}

async function readSpreadsheet(profileDirectory) {
  const excel = (await filesAt(profileDirectory)).find(file => {
    const extension = path.extname(file.name).toLowerCase();
    const base = normalize(path.basename(file.name, extension));
    return ['.xlsx', '.xls', '.xlsm'].includes(extension) && base === 'especificacoes tecnicas e descricao';
  });
  if (!excel) return { specifications: [], description: '' };

  const excelBuffer = await fs.readFile(excel.absolute);
  const workbook = XLSX.read(excelBuffer, { type: 'buffer', cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
  const specifications = [];
  let description = '';
  for (const row of rows) {
    const name = String(row[0] ?? '').trim();
    const value = String(row[1] ?? '').trim();
    if (!name) continue;
    if (normalize(name) === 'descricao') description = value;
    else specifications.push({ name, value });
  }
  return { specifications, description };
}

async function isProfileDirectory(directory) {
  const dirs = await directoriesAt(directory);
  const files = await filesAt(directory);
  const hasAssetFolder = dirs.some(item => ['fotos', 'equipamentos associados'].includes(normalize(item.name)));
  const hasExcel = files.some(file => {
    const extension = path.extname(file.name).toLowerCase();
    return ['.xlsx', '.xls', '.xlsm'].includes(extension) &&
      normalize(path.basename(file.name, extension)) === 'especificacoes tecnicas e descricao';
  });
  return hasAssetFolder || hasExcel;
}

async function scanArea(area) {
  const areaRoot = path.join(root, area.directory);
  const profiles = [];

  async function walk(directory) {
    if (await isProfileDirectory(directory)) {
      const relativeParts = path.relative(areaRoot, directory).split(path.sep).filter(Boolean);
      if (relativeParts.length < 2) return;
      const cat = relativeParts[0];
      const profileName = relativeParts.at(-1);
      const hierarchy = relativeParts.slice(1, -1);
      const folder = hierarchy.join(' / ');
      const photosDirectory = await findNamedDirectory(directory, 'Fotos');
      const equipmentDirectory = await findNamedDirectory(directory, 'Equipamentos Associados');
      const photos = await readImages(photosDirectory);
      const equipmentImages = await readImages(equipmentDirectory);
      const excelData = await readSpreadsheet(directory);
      profiles.push({
        kind: area.kind,
        cat,
        folder,
        shortTitle: folder ? profileName : '',
        title: folder ? `${folder} — ${profileName}` : profileName,
        sourcePath: toWebPath(directory),
        photos: photos.map(file => toWebPath(file.absolute)),
        equipment: equipmentImages.map(file => ({
          name: path.basename(file.name, path.extname(file.name)),
          image: toWebPath(file.absolute)
        })),
        specifications: excelData.specifications,
        description: excelData.description
      });
      return;
    }
    for (const child of await directoriesAt(directory)) await walk(child.absolute);
  }

  try { await fs.access(areaRoot); await walk(areaRoot); } catch { /* optional area */ }
  return profiles;
}

const profiles = (await Promise.all(areas.map(scanArea))).flat().sort((a, b) =>
  natural.compare(`${a.kind}/${a.cat}/${a.folder}/${a.title}`, `${b.kind}/${b.cat}/${b.folder}/${b.title}`)
);

await fs.writeFile(path.join(root, 'data-manifest.json'), JSON.stringify({
  generatedAt: new Date().toISOString(),
  profiles
}, null, 2) + '\n', 'utf8');
console.log(`data-manifest.json criado com ${profiles.length} perfis.`);
