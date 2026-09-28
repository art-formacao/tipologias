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
    .map(entry => ({
      name: entry.name,
      absolute: path.join(directory, entry.name)
    }))
    .sort((a, b) => natural.compare(a.name, b.name));
}

async function filesAt(directory) {
  return (await entriesAt(directory))
    .filter(entry => entry.isFile())
    .map(entry => ({
      name: entry.name,
      absolute: path.join(directory, entry.name)
    }))
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

  return (await directoriesAt(directory))
    .find(item => normalize(item.name) === expected);
}

async function readImages(directory) {
  if (!directory) return [];

  return (await filesAt(directory.absolute))
    .filter(file => isImage(file.name));
}

async function findFolderIcon(directory) {
  return (await filesAt(directory))
    .find(
      file =>
        isImage(file.name) &&
        normalize(
          path.basename(file.name, path.extname(file.name))
        ) === 'icone'
    );
}

async function findSpreadsheet(profileDirectory) {
  return (await filesAt(profileDirectory))
    .find(file => isSpreadsheet(file.name));
}

async function findLinkSpreadsheet(profileDirectory) {
  return (await filesAt(profileDirectory))
    .find(file => isLinkSpreadsheet(file.name));
}

function isWebLink(value) {
  return /^https?:\/\//i.test(String(value || '').trim());
}

async function readVideoLinks(excel) {
  if (!excel) return [];

  try {
    const workbook = XLSX.readFile(excel.absolute, {
      cellDates: false
    });

    const firstSheetName = workbook.SheetNames?.[0];

    if (!firstSheetName) return [];

    const sheet = workbook.Sheets[firstSheetName];

    const range = XLSX.utils.decode_range(
      sheet['!ref'] || 'A1:A1'
    );

    const links = [];

    for (
      let row = range.s.r;
      row <= range.e.r;
      row++
    ) {
      const cell = sheet[
        XLSX.utils.encode_cell({
          r: row,
          c: 0
        })
      ];

      const value = String(
        cell?.l?.Target ||
        cell?.v ||
        ''
      ).trim();

      if (
        isWebLink(value) &&
        !links.includes(value)
      ) {
        links.push(value);
      }
    }

    return links;

  } catch (error) {

    console.warn(
      `Não foi possível ler "${toWebPath(
        excel.absolute,
        false
      )}": ${error.message}`
    );

    return [];
  }
}

async function readSpreadsheet(excel) {

  const emptyData = {
    specifications: [],
    description: '',
    careBeforeAfter: '',
    attentionPoints: '',
    serviceConsiderations: ''
  };

  if (!excel) return emptyData;

  try {

    const workbook = XLSX.readFile(
      excel.absolute,
      {
        cellDates: false
      }
    );

    const firstSheetName =
      workbook.SheetNames?.[0];

    if (!firstSheetName)
      return emptyData;

    const rows =
      XLSX.utils.sheet_to_json(
        workbook.Sheets[firstSheetName],
        {
          header: 1,
          raw: false,
          defval: ''
        }
      );

    const specifications = [];

    let description = '';
    let careBeforeAfter = '';
    let attentionPoints = '';
    let serviceConsiderations = '';

    for (const row of rows) {

      const name = String(
        row?.[0] ?? ''
      ).trim();

      const value = String(
        row?.[1] ?? ''
      ).trim();

      if (!name) continue;

      const normalizedName =
        normalize(name)
          .replace(/:\s*$/, '')
          .replace(/\s+/g, ' ');

      if (
        normalizedName === 'descricao'
      ) {

        description = value;

      } else if (
        normalizedName ===
        'cuidados a ter antes e depois da utilizacao'
      ) {

        careBeforeAfter = value;

      } else if (
        normalizedName ===
        'pontos de atencao'
      ) {

        attentionPoints = value;

      } else if (
        normalizedName ===
        'aspetos a ter em conta durante o servico com este tipo de veiculo'
      ) {

        serviceConsiderations = value;

      } else {

        specifications.push({
          name,
          value
        });

      }
    }

    return {
      specifications,
      description,
      careBeforeAfter,
      attentionPoints,
      serviceConsiderations
    };

  } catch (error) {

    console.warn(
      `Não foi possível ler "${toWebPath(
        excel.absolute,
        false
      )}": ${error.message}`
    );

    return emptyData;
  }
}

async function isProfileDirectory(directory) {

  const directories =
    await directoriesAt(directory);

  const files =
    await filesAt(directory);

  if (
    directories.some(
      item =>
        [
          'fotos',
          'equipamentos associados'
        ].includes(
          normalize(item.name)
        )
    )
  ) {
    return true;
  }

  if (
    files.some(
      file =>
        isSpreadsheet(file.name)
    )
  ) {
    return true;
  }

  return false;
}

/*
 * Verifica se uma pasta contém os dados
 * necessários para ser considerada uma
 * variante, por exemplo:
 *
 * 2 Eixos/
 * 3 Eixos/
 * 4 Eixos/
 *
 * Cada uma pode conter:
 *
 * Fotos/
 * Equipamentos Associados/
 * Especificações Técnicas e Descrição.xlsx
 */
async function hasVariantDataDirectory(directory) {

  const directories =
    await directoriesAt(directory);

  const files =
    await filesAt(directory);

  return (
    directories.some(
      item =>
        [
          'fotos',
          'equipamentos associados'
        ].includes(
          normalize(item.name)
        )
    )
    ||
    files.some(
      file =>
        isSpreadsheet(file.name)
    )
  );
}

/*
 * Procura automaticamente as variantes
 * existentes dentro da pasta da tipologia.
 *
 * Exemplo:
 *
 * Bi Temperatura/
 * ├── Fotos/
 * ├── Equipamentos Associados/
 * ├── Especificações Técnicas e Descrição.xlsx
 * │
 * ├── 2 Eixos/
 * │   ├── Fotos/
 * │   ├── Equipamentos Associados/
 * │   └── Especificações Técnicas e Descrição.xlsx
 * │
 * └── 3 Eixos/
 *     ├── Fotos/
 *     ├── Equipamentos Associados/
 *     └── Especificações Técnicas e Descrição.xlsx
 *
 * O manifesto ficará com:
 *
 * variants: [
 *   {
 *     name: "2 Eixos",
 *     ...
 *   },
 *   {
 *     name: "3 Eixos",
 *     ...
 *   }
 * ]
 */
async function findVariants(profileDirectory) {

  const variants = [];

  for (
    const directory
    of await directoriesAt(profileDirectory)
  ) {

    const name =
      normalize(directory.name);

    /*
     * Estas duas pastas pertencem
     * diretamente à tipologia e não
     * são variantes.
     */
    if (
      name === 'fotos' ||
      name === 'equipamentos associados'
    ) {
      continue;
    }

    /*
     * Se a pasta tiver estrutura de dados,
     * será considerada uma variante.
     */
    if (
      await hasVariantDataDirectory(
        directory.absolute
      )
    ) {
      variants.push(directory);
    }
  }

  return variants;
}

async function readFolderData(directory) {

  const photosDirectory =
    await findNamedDirectory(
      directory,
      'Fotos'
    );

  const equipmentDirectory =
    await findNamedDirectory(
      directory,
      'Equipamentos Associados'
    );

  const photos =
    await readImages(
      photosDirectory
    );

  const equipmentImages =
    await readImages(
      equipmentDirectory
    );

  const excel =
    await findSpreadsheet(
      directory
    );

  const linkExcel =
    await findLinkSpreadsheet(
      directory
    );

  const [
    excelData,
    videos
  ] = await Promise.all([
    readSpreadsheet(excel),
    readVideoLinks(linkExcel)
  ]);

  return {

    sourcePath:
      toWebPath(directory),

    photos:
      photos.map(
        file =>
          toWebPath(file.absolute)
      ),

    equipment:
      equipmentImages.map(
        file => ({
          name: path.basename(
            file.name,
            path.extname(file.name)
          ),
          image:
            toWebPath(file.absolute)
        })
      ),

    excelPath:
      excel
        ? toWebPath(
            excel.absolute,
            false
          )
        : '',

    linkExcelPath:
      linkExcel
        ? toWebPath(
            linkExcel.absolute,
            false
          )
        : '',

    videos,

    videosLoaded: true,

    specifications:
      excelData.specifications,

    description:
      excelData.description,

    careBeforeAfter:
      excelData.careBeforeAfter,

    attentionPoints:
      excelData.attentionPoints,

    serviceConsiderations:
      excelData.serviceConsiderations,

    detailsLoaded: true
  };
}

async function scanProfile(
  directory,
  typeName,
  category,
  folder,
  profileName
) {

  /*
   * Dados da tipologia principal
   */
  const data =
    await readFolderData(
      directory.absolute
    );

  /*
   * Procura variantes como:
   * 2 Eixos
   * 3 Eixos
   * 4 Eixos
   * etc.
   */
  const variantDirectories =
    await findVariants(
      directory.absolute
    );

  const variants = [];

  for (
    const variantDirectory
    of variantDirectories
  ) {

    const variantData =
      await readFolderData(
        variantDirectory.absolute
      );

    variants.push({

      /*
       * Nome apresentado no botão
       * Exemplo: "2 Eixos"
       */
      name:
        variantDirectory.name,

      /*
       * Fotos, equipamentos,
       * especificações e restantes
       * dados dessa variante.
       */
      ...variantData

    });
  }

  return {

    kind:
      typeName,

    cat:
      category,

    folder,

    shortTitle:
      folder
        ? profileName
        : '',

    title:
      folder
        ? `${folder} — ${profileName}`
        : profileName,

    /*
     * Dados da tipologia principal
     */
    ...data,

    /*
     * Novos dados das variantes
     */
    variants

  };
}

async function scanType(
  typeDirectory
) {

  const typeIconFile =
    await findFolderIcon(
      typeDirectory.absolute
    );

  const categoryDirectories =
    await directoriesAt(
      typeDirectory.absolute
    );

  const categories =
    categoryDirectories.map(
      directory =>
        directory.name
    );

  const categoryIcons = {};

  for (
    const category
    of categoryDirectories
  ) {

    const iconFile =
      await findFolderIcon(
        category.absolute
      );

    if (iconFile) {

      categoryIcons[
        category.name
      ] =
        toWebPath(
          iconFile.absolute
        );

    }
  }

  const profiles = [];

  async function walk(directory) {

    if (
      await isProfileDirectory(
        directory
      )
    ) {

      const parts =
        path.relative(
          typeDirectory.absolute,
          directory
        )
        .split(path.sep)
        .filter(Boolean);

      if (
        parts.length < 2
      ) {
        return;
      }

      const category =
        parts[0];

      const profileName =
        parts.at(-1);

      const hierarchy =
        parts.slice(
          1,
          -1
        );

      const folder =
        hierarchy.join(
          ' / '
        );

      profiles.push(
        await scanProfile(
          directory,
          typeDirectory.name,
          category,
          folder,
          profileName
        )
      );

      return;
    }

    for (
      const child
      of await directoriesAt(
        directory
      )
    ) {

      await walk(child);

    }
  }

  for (
    const category
    of categoryDirectories
  ) {

    await walk(category);

  }

  return {

    type: {

      name:
        typeDirectory.name,

      categories,

      icon:
        typeIconFile
          ? toWebPath(
              typeIconFile.absolute
            )
          : '',

      categoryIcons

    },

    profiles

  };
}

const typeDirectories =
  await directoriesAt(
    tipologiasRoot
  );

const scanned =
  await Promise.all(
    typeDirectories.map(
      scanType
    )
  );

const types =
  scanned.map(
    result =>
      result.type
  );

const profiles =
  scanned
    .flatMap(
      result =>
        result.profiles
    )
    .sort(
      (a, b) =>
        natural.compare(
          `${a.kind}/${a.cat}/${a.folder}/${a.title}`,
          `${b.kind}/${b.cat}/${b.folder}/${b.title}`
        )
    );

const manifest = {

  generatedAt:
    new Date().toISOString(),

  types,

  profiles

};

await fs.writeFile(

  path.join(
    root,
    'data-manifest.json'
  ),

  `${JSON.stringify(
    manifest,
    null,
    2
  )}\n`,

  'utf8'

);

const typeIconCount =
  types.filter(
    type =>
      type.icon
  ).length;

const categoryIconCount =
  types.reduce(
    (total, type) =>
      total +
      Object.keys(
        type.categoryIcons
      ).length,
    0
  );

const variantCount =
  profiles.reduce(
    (total, profile) =>
      total +
      (profile.variants?.length || 0),
    0
  );

console.log(

  `data-manifest.json criado com ${types.length} tipos, ` +
  `${profiles.length} perfis, ` +
  `${variantCount} variantes internas, ` +
  `${typeIconCount} ícones principais e ` +
  `${categoryIconCount} ícones de categoria.`

);
