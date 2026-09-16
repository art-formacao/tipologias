import fs from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';

const root = process.cwd();
const tipologiasRoot = path.join(root, 'Tipologias');
const imageExtensions = new Set(['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif']);
const spreadsheetExtensions = new Set(['.xlsx', '.xls', '.xlsm']);
const natural = new Intl.Collator('pt', { numeric: true, sensitivity: 'base' });

function normalize(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function toWebPath(absolutePath, withDot = true) {
  const relative = path.relative(root, absolutePath).split(path.sep).join('/');
  return withDot ? `./${relative}` : relative;
}

async function entriesAt(directory) {
  try {
    return await fs.readdir(directory, { withFileTypes: true });
  } catch {
    return [];
  }
}

async function directoriesAt(directory) {
  return (await entriesAt(directory))
    .filter(entry => entry.isDirectory())
    .map(entry => ({ name: entry.name, absolute: path.join(directory, entry.name) }))
    .sort((a, b) => natural.compare(a.name, b.name));
}

async function filesAt(directory) {
  return (await entriesAt(directory))
    .filter(entry => entry.isFile())
    .map(entry => ({ name: entry.name, absolute: path.join(directory, entry.name) }))
    .sort((a, b) => natural.compare(a.name, b.name));
}

function isImage(fileName) {
  return imageExtensions.has(path.extname(fileName).toLowerCase());
}

function isSpreadsheet(fileName) {
  const extension = path.extname(fileName).toLowerCase();
  return spreadsheetExtensions.has(extension) &&
    normalize(path.basename(fileName, extension)) === 'especificacoes tecnicas e descricao';
}

async function findNamedDirectory(directory, expectedName) {
  const expected = normalize(expectedName);
  return (await directoriesAt(directory)).find(item => normalize(item.name) === expected);
}

async function readImages(directory) {
  if (!directory) return [];
  return (await filesAt(directory.absolute)).filter(file => isImage(file.name));
}

async function findFolderIcon(directory) {
  return (await filesAt(directory))
    .find(file => isImage(file.name) && normalize(path.basename(file.name, path.extname(file.name))) === 'icone');
}

async function findSpreadsheet(profileDirectory) {
  return (await filesAt(profileDirectory)).find(file => isSpreadsheet(file.name));
}

async function readSpreadsheet(excel) {
  if (!excel) return { specifications: [], description: '' };
  try {
    const workbook = XLSX.readFile(excel.absolute, { cellDates: false });
    const firstSheetName = workbook.SheetNames?.[0];
    if (!firstSheetName) return { specifications: [], description: '' };
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      header: 1,
      raw: false,
      defval: ''
    });
    const specifications = [];
    let description = '';
    for (const row of rows) {
      const name = String(row?.[0] ?? '').trim();
      const value = String(row?.[1] ?? '').trim();
      if (!name) continue;
      if (normalize(name) === 'descricao') description = value;
      else specifications.push({ name, value });
    }
    return { specifications, description };
  } catch (error) {
    console.warn(`Não foi possível ler "${toWebPath(excel.absolute, false)}": ${error.message}`);
    return { specifications: [], description: '' };
  }
}

async function isProfileDirectory(directory) {
  const directories = await directoriesAt(directory);
  const files = await filesAt(directory);
  return directories.some(item => ['fotos', 'equipamentos associados'].includes(normalize(item.name))) ||
    files.some(file => isSpreadsheet(file.name));
}

async function scanType(typeDirectory) {
  const typeIconFile = await findFolderIcon(typeDirectory.absolute);
  const categoryDirectories = await directoriesAt(typeDirectory.absolute);
  const categories = categoryDirectories.map(directory => directory.name);
  const categoryIcons = {};
  for (const category of categoryDirectories) {
    const iconFile = await findFolderIcon(category.absolute);
    if (iconFile) categoryIcons[category.name] = toWebPath(iconFile.absolute);
  }
  const profiles = [];

  async function walk(directory) {
    if (await isProfileDirectory(directory)) {
      const parts = path.relative(typeDirectory.absolute, directory).split(path.sep).filter(Boolean);
      if (parts.length < 2) return;

      const category = parts[0];
      const profileName = parts.at(-1);
      const hierarchy = parts.slice(1, -1);
      const folder = hierarchy.join(' / ');
      const photosDirectory = await findNamedDirectory(directory, 'Fotos');
      const equipmentDirectory = await findNamedDirectory(directory, 'Equipamentos Associados');
      const photos = await readImages(photosDirectory);
      const equipmentImages = await readImages(equipmentDirectory);
      const excel = await findSpreadsheet(directory);
      const excelData = await readSpreadsheet(excel);

      profiles.push({
        kind: typeDirectory.name,
        cat: category,
        folder,
        shortTitle: folder ? profileName : '',
        title: folder ? `${folder} — ${profileName}` : profileName,
        sourcePath: toWebPath(directory),
        photos: photos.map(file => toWebPath(file.absolute)),
        equipment: equipmentImages.map(file => ({
          name: path.basename(file.name, path.extname(file.name)),
          image: toWebPath(file.absolute)
        })),
        excelPath: excel ? toWebPath(excel.absolute, false) : '',
        specifications: excelData.specifications,
        description: excelData.description,
        detailsLoaded: true
      });
      return;
    }

    for (const child of await directoriesAt(directory)) {
      await walk(child.absolute);
    }
  }

  for (const category of categoryDirectories) {
    await walk(category.absolute);
  }

  return {
    type: {
      name: typeDirectory.name,
      categories,
      icon: typeIconFile ? toWebPath(typeIconFile.absolute) : '',
      categoryIcons
    },
    profiles
  };
}

const typeDirectories = await directoriesAt(tipologiasRoot);
const scanned = await Promise.all(typeDirectories.map(scanType));
const types = scanned.map(result => result.type);
const profiles = scanned
  .flatMap(result => result.profiles)
  .sort((a, b) => natural.compare(
    `${a.kind}/${a.cat}/${a.folder}/${a.title}`,
    `${b.kind}/${b.cat}/${b.folder}/${b.title}`
  ));

const manifest = {
  generatedAt: new Date().toISOString(),
  types,
  profiles
};

await fs.writeFile(
  path.join(root, 'data-manifest.json'),
  `${JSON.stringify(manifest, null, 2)}\n`,
  'utf8'
);

const typeIconCount = types.filter(type => type.icon).length;
const categoryIconCount = types.reduce(
  (total, type) => total + Object.keys(type.categoryIcons).length,
  0
);

console.log(
  `data-manifest.json criado com ${types.length} tipos, ${profiles.length} perfis, ` +
  `${typeIconCount} ícones principais e ${categoryIconCount} ícones de categoria.`
);
