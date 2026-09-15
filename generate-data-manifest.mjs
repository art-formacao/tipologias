import fs from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';

const root = process.cwd();
const typologiesRoot = path.join(root, 'Tipologias');
const imageExtensions = new Set(['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif']);
const natural = new Intl.Collator('pt', { numeric: true, sensitivity: 'base' });

function normalize(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}
function toWebPath(absolutePath) {
  return './' + path.relative(root, absolutePath).split(path.sep).join('/');
}
async function entriesAt(directory) {
  try { return await fs.readdir(directory, { withFileTypes: true }); } catch { return []; }
}
async function directoriesAt(directory) {
  return (await entriesAt(directory)).filter(entry => entry.isDirectory()).map(entry => ({
    name: entry.name, absolute: path.join(directory, entry.name)
  })).sort((a, b) => natural.compare(a.name, b.name));
}
async function filesAt(directory) {
  return (await entriesAt(directory)).filter(entry => entry.isFile()).map(entry => ({
    name: entry.name, absolute: path.join(directory, entry.name)
  }));
}
async function findNamedDirectory(directory, name) {
  return (await directoriesAt(directory)).find(item => normalize(item.name) === normalize(name));
}
async function readImages(directory) {
  if (!directory) return [];
  return (await filesAt(directory.absolute))
    .filter(file => imageExtensions.has(path.extname(file.name).toLowerCase()))
    .sort((a, b) => natural.compare(a.name, b.name));
}
async function findExcel(directory) {
  return (await filesAt(directory)).find(file => {
    const extension = path.extname(file.name).toLowerCase();
    return ['.xlsx', '.xls', '.xlsm'].includes(extension) &&
      normalize(path.basename(file.name, extension)) === 'especificacoes tecnicas e descricao';
  });
}
async function readSpreadsheet(directory) {
  const excel = await findExcel(directory);
  if (!excel) return { specifications: [], description: '' };
  const buffer = await fs.readFile(excel.absolute);
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
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
async function isProfile(directory) {
  const folders = await directoriesAt(directory);
  return folders.some(folder => ['fotos', 'equipamentos associados'].includes(normalize(folder.name))) || Boolean(await findExcel(directory));
}

const types = [];
const profiles = [];
for (const typeDirectory of await directoriesAt(typologiesRoot)) {
  const categories = (await directoriesAt(typeDirectory.absolute)).map(category => category.name);
  types.push({ name: typeDirectory.name, categories });

  async function walk(directory) {
    if (await isProfile(directory)) {
      const parts = path.relative(typeDirectory.absolute, directory).split(path.sep).filter(Boolean);
      if (parts.length < 2) return;
      const cat = parts[0];
      const profileName = parts.at(-1);
      const hierarchy = parts.slice(1, -1);
      const folder = hierarchy.join(' / ');
      const photosFolder = await findNamedDirectory(directory, 'Fotos');
      const equipmentFolder = await findNamedDirectory(directory, 'Equipamentos Associados');
      const photos = await readImages(photosFolder);
      const equipmentPhotos = await readImages(equipmentFolder);
      const excel = await readSpreadsheet(directory);
      profiles.push({
        kind: typeDirectory.name,
        cat,
        folder,
        shortTitle: folder ? profileName : '',
        title: folder ? `${folder} — ${profileName}` : profileName,
        sourcePath: toWebPath(directory),
        photos: photos.map(photo => toWebPath(photo.absolute)),
        equipment: equipmentPhotos.map(photo => ({
          name: path.basename(photo.name, path.extname(photo.name)),
          image: toWebPath(photo.absolute)
        })),
        specifications: excel.specifications,
        description: excel.description
      });
      return;
    }
    for (const child of await directoriesAt(directory)) await walk(child.absolute);
  }
  for (const category of await directoriesAt(typeDirectory.absolute)) await walk(category.absolute);
}

profiles.sort((a, b) => natural.compare(
  `${a.kind}/${a.cat}/${a.folder}/${a.title}`,
  `${b.kind}/${b.cat}/${b.folder}/${b.title}`
));
await fs.writeFile(path.join(root, 'data-manifest.json'), JSON.stringify({
  generatedAt: new Date().toISOString(), types, profiles
}, null, 2) + '\n', 'utf8');
console.log(`Manifesto criado: ${types.length} tipologias principais e ${profiles.length} perfis.`);
