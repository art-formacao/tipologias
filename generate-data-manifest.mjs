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

function spreadsheetStem(fileName) {
  const extension = path.extname(fileName).toLowerCase();
  return spreadsheetExtensions.has(extension)
    ? normalize(path.basename(fileName, extension))
    : '';
}

function isSpreadsheet(fileName) {
  return spreadsheetStem(fileName) === 'especificacoes tecnicas e descricao';
}

function isLinkSpreadsheet(fileName) {
  return spreadsheetStem(fileName) === 'link';
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

async function findLinkSpreadsheet(profileDirectory) {
  return (await filesAt(profileDirectory)).find(file => isLinkSpreadsheet(file.name));
}

function isWebLink(value) {
  return /^https?:\/\//i.test(String(value || '').trim());
}

async function readVideoLinks(excel) {
  if (!excel) return [];
  try {
    const workbook = XLSX.readFile(excel.absolute, { cellDates: false });
    const firstSheetName = workbook.SheetNames?.[0];
    if (!firstSheetName) return [];
    const sheet = workbook.Sheets[firstSheetName];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
    const links = [];
    for (let row = range.s.r; row <= range.e.r; row++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
      const value = String(cell?.l?.Target || cell?.v || '').trim();
      if (isWebLink(value) && !links.includes(value)) links.push(value);
    }
    return links;
  } catch (error) {
    console.warn(`Não foi possível ler "${toWebPath(excel.absolute, false)}": ${error.message}`);
    return [];
  }
}

async function readSpreadsheet(excel) {
  const emptyData = {
    specifications: [],
    textSections: []
  };
  if (!excel) return emptyData;
  try {
    const workbook = XLSX.readFile(excel.absolute, { cellDates: false });
    const plan1Name = workbook.SheetNames?.find(name => normalize(name) === 'plan1') || workbook.SheetNames?.[0];
    const plan2Name = workbook.SheetNames?.find(name => normalize(name) === 'plan2');
    if (!plan1Name) return emptyData;
    const rowsFor = sheetName => XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: false,
      defval: ''
    });
    const plan1Rows = rowsFor(plan1Name);
    const plan2Rows = plan2Name ? rowsFor(plan2Name) : [];
    const specifications = [];
    const plan1TextSections = [];
    const knownTextLabels = new Set([
      'descricao',
      'cuidados a ter antes e depois da utilizacao',
      'pontos de atencao',
      'aspetos a ter em conta durante o servico com este tipo de veiculo'
    ]);
    let activeSection = null;
    for (const row of plan1Rows) {
      const name = String(row?.[0] ?? '').trim();
      const value = String(row?.[1] ?? '').trim();
      if (!name) {
        if (value && activeSection) activeSection.value += `\n${value}`;
        continue;
      }
      const normalizedName = normalize(name).replace(/:\s*$/, '').replace(/\s+/g, ' ');
      if (knownTextLabels.has(normalizedName)) {
        activeSection = { title: name, value };
        plan1TextSections.push(activeSection);
      } else {
        activeSection = null;
        specifications.push({ name, value });
      }
    }
    const textSections = [];
    let plan2Section = null;
    for (const row of plan2Rows) {
      const name = String(row?.[0] ?? '').trim();
      const value = String(row?.[1] ?? '').trim();
      if (!name) {
        if (value && plan2Section) plan2Section.value += `\n${value}`;
        continue;
      }
      plan2Section = { title: name, value };
      textSections.push(plan2Section);
    }
    return {
      specifications,
      textSections: textSections.length ? textSections : plan1TextSections
    };
  } catch (error) {
    console.warn(`Não foi possível ler "${toWebPath(excel.absolute, false)}": ${error.message}`);
    return emptyData;
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
      const linkExcel = await findLinkSpreadsheet(directory);
      const [excelData, videos] = await Promise.all([
        readSpreadsheet(excel),
        readVideoLinks(linkExcel)
      ]);

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
        linkExcelPath: linkExcel ? toWebPath(linkExcel.absolute, false) : '',
        videos,
        videosLoaded: true,
        specifications: excelData.specifications,
        textSections: excelData.textSections,
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
