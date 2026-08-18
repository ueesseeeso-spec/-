import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import { chromium } from 'playwright';
import sharp from 'sharp';

const ORIGINAL_PRODUCTS = [
  { sku: '1825205889', expected: 'Отпугиватель мышей ультразвуковой комплект 2 шт' },
  { sku: '1663754371', expected: 'Плата управления бесщеточной литиевой угловой шлифовальной машины' },
  { sku: '1724766531', expected: 'Ограничитель для двери 4 шт' },
  { sku: '723815479', expected: 'Глайды для Logitech G Pro X Superlight ножки для игровой мыши GPX FINALGEAR' },
  { sku: '2026793880', expected: 'Проушина и накладка для навесного замка петли для навесного замка RAMI' },
  { sku: '1950955543', expected: 'Точилка для сверл подходит для угловой шлифовальной машины 125 для 2-13 мм' },
  { sku: '1945180142', expected: 'Держатели для простыни фиксаторы для простыни на матрас набор 4 шт' },
  { sku: '1769339465', expected: 'Инструмент для установки резьбовых заклепок адаптер для шуруповерта заклепки M6' },
  { sku: '1849833851', expected: 'Ролик щетка для снятия ворсинок 1 шт' },
  { sku: '1967664346', expected: 'Контроллер платы управления бесщеточной литиевой электрической угловой шлифовальной машиной 60А' },
  { sku: '1675807532', expected: 'Т образная угловая линейка уголок строительный 169 мм для деревообработки 45 90' },
  { sku: '3210992423', expected: 'Воронка для холдера 51 мм дозирующее кольцо для портафильтра' },
  { sku: '3230034311', expected: 'Кисти для моделирования миниатюр 11 шт набор тонких кистей' },
  { sku: '2862477326', expected: 'Органайзер для хранения вещей 12.5х7.5' },
  { sku: '2400149322', expected: 'Мини CD альбом чехол брелок для ключей DIY рюкзак' },
  { sku: '907234276', expected: 'Брелок для ключей Микки со стразами кольцо с карабином' }
];

const HOME = 'https://www.ozon.ru/';
const COMPOSER = 'https://www.ozon.ru/api/composer-api.bx/page/json/v2?url=';
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36';
const NAV_TIMEOUT = 90_000;
const CHALLENGE_WAIT = 15_000;
const MAX_SEARCH_RESULTS = 36;
const MAX_DETAIL_CANDIDATES = 8;
const MAX_IMAGES_PER_CARD = 5;
const WAIT_BETWEEN_REQUESTS = 450;

const STOPWORDS = new Set([
  'для', 'и', 'на', 'в', 'с', 'из', 'по', 'к', 'от', 'под', 'над', 'или', 'шт', 'штук', 'набор',
  'товар', 'универсальный', 'профессиональный', 'электрический', 'машины', 'машина', 'инструмент',
  'черный', 'белый', 'серебристый', 'комплект'
]);

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const log = (...args) => console.error(new Date().toISOString(), ...args);

function normalizeText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[×хx*]/g, 'x')
    .replace(/[^a-zа-я0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokens(value) {
  return normalizeText(value)
    .split(' ')
    .filter((token) => token.length > 1 && !STOPWORDS.has(token));
}

function jaccard(a, b) {
  const aa = new Set(tokens(a));
  const bb = new Set(tokens(b));
  if (!aa.size || !bb.size) return 0;
  let intersection = 0;
  for (const token of aa) if (bb.has(token)) intersection += 1;
  return intersection / (aa.size + bb.size - intersection);
}

function containment(a, b) {
  const aa = new Set(tokens(a));
  const bb = new Set(tokens(b));
  if (!aa.size || !bb.size) return 0;
  let intersection = 0;
  for (const token of aa) if (bb.has(token)) intersection += 1;
  return intersection / Math.min(aa.size, bb.size);
}

function textSimilarity(a, b) {
  const na = normalizeText(a);
  const nb = normalizeText(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (na.length >= 40 && nb.length >= 40 && (na.includes(nb) || nb.includes(na))) return 0.98;
  return Math.max(jaccard(na, nb), containment(na, nb) * 0.92);
}

function searchQuery(title, fallback) {
  const source = title || fallback;
  const meaningful = tokens(source)
    .filter((token) => !/^\d+$/.test(token) || token.length >= 2)
    .slice(0, 11);
  return meaningful.join(' ') || normalizeText(fallback).slice(0, 100);
}

function widgetName(key) {
  return String(key).split('-')[0];
}

function widget(page, name) {
  const states = page?.widgetStates || {};
  const key = Object.keys(states).find((item) => widgetName(item) === name);
  if (!key) return null;
  try {
    return JSON.parse(states[key]);
  } catch {
    return null;
  }
}

function widgets(page, name) {
  const states = page?.widgetStates || {};
  return Object.keys(states)
    .filter((item) => widgetName(item) === name)
    .map((key) => {
      try {
        return JSON.parse(states[key]);
      } catch {
        return null;
      }
    })
    .filter(Boolean);
}

function priceToNumber(text) {
  if (typeof text !== 'string') return null;
  const digits = text.replace(/[^\d]/g, '');
  return digits ? Number.parseInt(digits, 10) : null;
}

function absoluteUrl(value) {
  if (!value) return null;
  let url = String(value).trim();
  if (url.startsWith('//')) url = `https:${url}`;
  if (url.startsWith('/')) url = `https://www.ozon.ru${url}`;
  return url;
}

function cleanUrl(value) {
  const url = absoluteUrl(value);
  return url ? url.split('?')[0] : null;
}

function skuFromUrl(value) {
  const match = String(value || '').match(/-(\d{6,})\/?(?:\?|$)/) || String(value || '').match(/(\d{6,})/);
  return match ? match[1] : null;
}

function parseSearchItem(item) {
  if (!item) return null;
  const mainState = Array.isArray(item.mainState) ? item.mainState : [];
  const priceBlock = mainState.find((state) => state.type === 'priceV2')?.priceV2;
  const prices = priceBlock?.price || [];
  const price = priceToNumber(prices.find((part) => part.textStyle === 'PRICE')?.text);
  const name = mainState.find((state) => state.id === 'name')?.textDS?.text || null;
  const url = cleanUrl(item.action?.link);
  const sku = String(item.sku || item.id || skuFromUrl(url) || '') || null;
  const image = absoluteUrl(
    item.tileImage?.items?.find((part) => part.image?.link)?.image?.link ||
    item.tileImage?.coverImage ||
    null
  );
  if (!sku || !name) return null;
  return { sku, name, price, url: url || `https://www.ozon.ru/product/${sku}/`, image };
}

function parseSearch(page) {
  const grids = widgets(page, 'tileGridDesktop');
  const raw = grids.flatMap((grid) => Array.isArray(grid?.items) ? grid.items : []);
  const seen = new Set();
  const items = [];
  for (const item of raw.map(parseSearchItem).filter(Boolean)) {
    if (seen.has(item.sku)) continue;
    seen.add(item.sku);
    items.push(item);
    if (items.length >= MAX_SEARCH_RESULTS) break;
  }
  return items;
}

function richText(nodes) {
  return (nodes || [])
    .map((node) => node?.text || node?.content)
    .filter(Boolean)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseCharacteristics(page) {
  const result = {};
  const short = widget(page, 'webShortCharacteristics');
  for (const characteristic of short?.characteristics || []) {
    const title = richText(characteristic.title?.textRs) || (typeof characteristic.title === 'string' ? characteristic.title : '');
    const value = richText(characteristic.values || characteristic.contentRS || characteristic.valueRs);
    if (title && value) result[title] = value;
  }
  return result;
}

function parseSeller(page) {
  const current = widget(page, 'webCurrentSeller');
  if (!current) return null;
  const name = current.sellerCell?.centerBlock?.title?.text || current.title?.text || null;
  const url = cleanUrl(current.sellerCell?.common?.action?.link);
  return name ? { name, url } : null;
}

function parseDescription(page) {
  const descriptionWidget = widgets(page, 'webDescription').find((item) => item.richAnnotationJson);
  if (!descriptionWidget) return { text: '', images: [] };
  let annotation = descriptionWidget.richAnnotationJson;
  if (typeof annotation === 'string') {
    try {
      annotation = JSON.parse(annotation);
    } catch {
      return { text: '', images: [] };
    }
  }
  const texts = [];
  const images = [];
  const walk = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach(walk);
      return;
    }
    if (typeof node !== 'object') return;
    if (node.type === 'text' && typeof node.content === 'string') texts.push(node.content);
    if (node.img?.src) images.push(absoluteUrl(node.img.src));
    for (const value of Object.values(node)) if (value && typeof value === 'object') walk(value);
  };
  walk(annotation.content || annotation);
  return {
    text: texts.join(' ').replace(/\s+/g, ' ').trim(),
    images: [...new Set(images.filter(Boolean))]
  };
}

function parseDetails(basePage, page2, requestedSku) {
  const heading = widget(basePage, 'webProductHeading');
  const gallery = widget(basePage, 'webGallery');
  const price = widget(basePage, 'webPrice');
  const characteristics = parseCharacteristics(basePage);
  const description = parseDescription(page2);
  const sku = String(gallery?.sku || requestedSku || skuFromUrl(basePage?.seo?.link?.[0]?.href) || '');
  const galleryImages = [];
  if (gallery?.coverImage) galleryImages.push(absoluteUrl(gallery.coverImage));
  for (const image of gallery?.images || []) {
    const source = typeof image === 'string' ? image : image?.src || image?.image;
    if (source) galleryImages.push(absoluteUrl(source));
  }
  const images = [...new Set([...galleryImages, ...description.images].filter(Boolean))].slice(0, 12);
  const countryEntry = Object.entries(characteristics).find(([key]) => /страна.*изготов|страна производства/i.test(key));
  return {
    sku,
    name: heading?.title || basePage?.seo?.title || null,
    url: `https://www.ozon.ru/product/${sku}/`,
    seller: parseSeller(basePage),
    country: countryEntry?.[1] || null,
    price: priceToNumber(price?.cardPrice) ?? priceToNumber(price?.price),
    images,
    characteristics,
    description: description.text
  };
}

async function startBrowser() {
  const browser = await chromium.launch({
    headless: true,
    args: [
      '--disable-blink-features=AutomationControlled',
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--no-first-run',
      '--no-default-browser-check'
    ]
  });
  const context = await browser.newContext({
    viewport: { width: 1920, height: 1080 },
    userAgent: USER_AGENT,
    locale: 'ru-RU',
    timezoneId: 'Europe/Moscow'
  });
  const page = await context.newPage();
  log('Opening Ozon and passing browser challenge');
  await page.goto(HOME, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT });
  await page.waitForTimeout(CHALLENGE_WAIT);
  const title = await page.title();
  log('Ozon title:', title);
  if (/antibot|ограничен|доступ запрещен|access denied/i.test(title)) {
    throw new Error(`Ozon challenge was not passed: ${title}`);
  }
  return { browser, context, page };
}

async function composerFetch(page, path, retries = 2) {
  let lastError;
  for (let attempt = 1; attempt <= retries + 1; attempt += 1) {
    try {
      const response = await page.evaluate(async (url) => {
        const result = await fetch(url, { headers: { accept: 'application/json' }, credentials: 'include' });
        return { status: result.status, body: await result.text() };
      }, COMPOSER + encodeURIComponent(path));
      if (response.status !== 200) throw new Error(`HTTP ${response.status}`);
      await sleep(WAIT_BETWEEN_REQUESTS);
      return JSON.parse(response.body);
    } catch (error) {
      lastError = error;
      log(`Composer request failed (${attempt}/${retries + 1})`, path, String(error));
      await sleep(1500 * attempt);
    }
  }
  throw lastError;
}

async function getDetails(page, sku) {
  const path = `/product/${sku}/`;
  const base = await composerFetch(page, path);
  const page2 = await composerFetch(page, `${path}?layout_container=pdpPage2column&layout_page_index=2`);
  return parseDetails(base, page2, sku);
}

async function searchProducts(page, query) {
  const path = `/search/?text=${encodeURIComponent(query)}&from_global=true`;
  const response = await composerFetch(page, path);
  return parseSearch(response);
}

const fingerprintCache = new Map();

async function fetchImageBuffer(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 20_000);
  try {
    const response = await fetch(url, {
      headers: { 'user-agent': USER_AGENT, accept: 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8' },
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`image HTTP ${response.status}`);
    return Buffer.from(await response.arrayBuffer());
  } finally {
    clearTimeout(timer);
  }
}

async function imageFingerprint(url) {
  if (!url) return null;
  if (fingerprintCache.has(url)) return fingerprintCache.get(url);
  const promise = (async () => {
    try {
      const input = await fetchImageBuffer(url);
      const sha256 = crypto.createHash('sha256').update(input).digest('hex');
      const dhashPixels = await sharp(input)
        .rotate()
        .resize(9, 8, { fit: 'fill' })
        .greyscale()
        .raw()
        .toBuffer();
      const bits = [];
      for (let y = 0; y < 8; y += 1) {
        for (let x = 0; x < 8; x += 1) {
          bits.push(dhashPixels[y * 9 + x] > dhashPixels[y * 9 + x + 1] ? 1 : 0);
        }
      }
      const pixels = await sharp(input)
        .rotate()
        .resize(32, 32, { fit: 'fill' })
        .greyscale()
        .raw()
        .toBuffer();
      return { url, sha256, bits, pixels: Array.from(pixels) };
    } catch (error) {
      log('Image fingerprint failed', url, String(error));
      return null;
    }
  })();
  fingerprintCache.set(url, promise);
  return promise;
}

function hamming(a, b) {
  if (!a || !b || a.length !== b.length) return 64;
  let distance = 0;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) distance += 1;
  return distance;
}

function pixelSimilarity(a, b) {
  if (!a || !b || a.length !== b.length) return 0;
  let difference = 0;
  for (let i = 0; i < a.length; i += 1) difference += Math.abs(a[i] - b[i]);
  return Math.max(0, 1 - difference / (255 * a.length));
}

async function compareImages(originalImages, candidateImages) {
  const originals = (await Promise.all(originalImages.slice(0, MAX_IMAGES_PER_CARD).map(imageFingerprint))).filter(Boolean);
  const candidates = (await Promise.all(candidateImages.slice(0, MAX_IMAGES_PER_CARD).map(imageFingerprint))).filter(Boolean);
  let best = null;
  for (let originalIndex = 0; originalIndex < originals.length; originalIndex += 1) {
    for (let candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
      const original = originals[originalIndex];
      const candidate = candidates[candidateIndex];
      const exact = original.sha256 === candidate.sha256;
      const distance = hamming(original.bits, candidate.bits);
      const pixels = pixelSimilarity(original.pixels, candidate.pixels);
      const rank = exact ? 10_000 : pixels * 100 - distance;
      if (!best || rank > best.rank) {
        best = {
          rank,
          exact,
          dhashDistance: distance,
          pixelSimilarity: Number(pixels.toFixed(4)),
          originalImage: original.url,
          candidateImage: candidate.url,
          originalImageNumber: originalIndex + 1,
          candidateImageNumber: candidateIndex + 1
        };
      }
    }
  }
  if (!best) return null;
  delete best.rank;
  return best;
}

function confidenceFor(match) {
  const image = match.imageMatch;
  const description = match.descriptionSimilarity;
  const title = match.titleSimilarity;
  if (image?.exact) return 'HIGH';
  if (image && (image.pixelSimilarity >= 0.965 || (image.dhashDistance <= 5 && image.pixelSimilarity >= 0.88))) return 'HIGH';
  if (image && (image.pixelSimilarity >= 0.92 || (image.dhashDistance <= 10 && image.pixelSimilarity >= 0.82)) && title >= 0.55) return 'MEDIUM';
  if (description >= 0.9 && match.originalDescriptionLength >= 180 && match.candidateDescriptionLength >= 180) return 'MEDIUM';
  return 'LOW';
}

function csvEscape(value) {
  const text = value == null ? '' : String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function auditOne(page, source) {
  log('Auditing original', source.sku);
  let original;
  try {
    original = await getDetails(page, source.sku);
  } catch (error) {
    return { source, error: `Не удалось открыть исходную карточку: ${String(error)}`, original: null, matches: [] };
  }
  const query = searchQuery(original.name, source.expected);
  log('Search query', source.sku, query);
  let searchItems = [];
  try {
    searchItems = await searchProducts(page, query);
  } catch (error) {
    return { source, query, original, error: `Не удалось выполнить поиск: ${String(error)}`, matches: [] };
  }

  const ownSeller = normalizeText(original.seller?.name);
  const ranked = searchItems
    .filter((item) => item.sku !== original.sku)
    .map((item) => ({ ...item, searchTitleSimilarity: textSimilarity(original.name || source.expected, item.name) }))
    .filter((item) => item.searchTitleSimilarity >= 0.24)
    .sort((a, b) => b.searchTitleSimilarity - a.searchTitleSimilarity)
    .slice(0, MAX_DETAIL_CANDIDATES);

  const matches = [];
  for (const item of ranked) {
    let candidate;
    try {
      candidate = await getDetails(page, item.sku);
    } catch (error) {
      log('Candidate details failed', item.sku, String(error));
      continue;
    }
    if (ownSeller && ownSeller === normalizeText(candidate.seller?.name)) continue;

    const titleSimilarity = textSimilarity(original.name || source.expected, candidate.name || item.name);
    const descriptionSimilarity = textSimilarity(original.description, candidate.description);
    const imageMatch = await compareImages(original.images, candidate.images);
    const record = {
      originalSku: original.sku,
      originalUrl: original.url,
      originalName: original.name || source.expected,
      originalSeller: original.seller?.name || null,
      candidateSku: candidate.sku,
      candidateUrl: candidate.url,
      candidateName: candidate.name || item.name,
      candidateSeller: candidate.seller?.name || null,
      candidateCountry: candidate.country,
      titleSimilarity: Number(titleSimilarity.toFixed(4)),
      descriptionSimilarity: Number(descriptionSimilarity.toFixed(4)),
      originalDescriptionLength: original.description?.length || 0,
      candidateDescriptionLength: candidate.description?.length || 0,
      imageMatch
    };
    record.confidence = confidenceFor(record);
    if (record.confidence !== 'LOW' || titleSimilarity >= 0.78 || descriptionSimilarity >= 0.75) matches.push(record);
  }

  matches.sort((a, b) => {
    const level = { HIGH: 3, MEDIUM: 2, LOW: 1 };
    return level[b.confidence] - level[a.confidence]
      || (b.imageMatch?.pixelSimilarity || 0) - (a.imageMatch?.pixelSimilarity || 0)
      || b.descriptionSimilarity - a.descriptionSimilarity;
  });

  return { source, query, original, searchResultCount: searchItems.length, matches };
}

async function main() {
  await fs.mkdir('output', { recursive: true });
  const startedAt = new Date().toISOString();
  const runtime = await startBrowser();
  const reports = [];
  try {
    for (const source of ORIGINAL_PRODUCTS) reports.push(await auditOne(runtime.page, source));
  } finally {
    await runtime.browser.close();
  }

  const allMatches = reports.flatMap((report) => report.matches || []);
  const high = allMatches.filter((item) => item.confidence === 'HIGH');
  const medium = allMatches.filter((item) => item.confidence === 'MEDIUM');
  const errors = reports.filter((item) => item.error);
  const finishedAt = new Date().toISOString();

  const result = {
    startedAt,
    finishedAt,
    methodology: {
      scope: 'Публичные карточки Ozon, найденные по названиям исходных товаров',
      highConfidence: 'Точное или очень близкое техническое совпадение изображения',
      mediumConfidence: 'Вероятное совпадение изображения либо почти дословный длинный текст',
      warning: 'Совпадение не доказывает авторские права. Жалоба допустима только при наличии исходников или договора на создание контента.'
    },
    totals: {
      originalsRequested: ORIGINAL_PRODUCTS.length,
      originalsOpened: reports.filter((item) => item.original).length,
      errors: errors.length,
      candidatesInspected: allMatches.length,
      highConfidence: high.length,
      mediumConfidence: medium.length
    },
    reports
  };

  await fs.writeFile('output/report.json', JSON.stringify(result, null, 2), 'utf8');

  const headers = [
    'confidence', 'originalSku', 'originalName', 'originalSeller', 'originalUrl',
    'candidateSku', 'candidateName', 'candidateSeller', 'candidateCountry', 'candidateUrl',
    'titleSimilarity', 'descriptionSimilarity', 'imageExact', 'imageDhashDistance',
    'imagePixelSimilarity', 'originalImage', 'candidateImage'
  ];
  const csvRows = [headers.join(',')];
  for (const item of allMatches) {
    const values = [
      item.confidence, item.originalSku, item.originalName, item.originalSeller, item.originalUrl,
      item.candidateSku, item.candidateName, item.candidateSeller, item.candidateCountry, item.candidateUrl,
      item.titleSimilarity, item.descriptionSimilarity, item.imageMatch?.exact,
      item.imageMatch?.dhashDistance, item.imageMatch?.pixelSimilarity,
      item.imageMatch?.originalImage, item.imageMatch?.candidateImage
    ];
    csvRows.push(values.map(csvEscape).join(','));
  }
  await fs.writeFile('output/report.csv', csvRows.join('\n'), 'utf8');

  const lines = [
    '# Проверка возможного копирования карточек Ozon',
    '',
    `Проверено исходных карточек: ${result.totals.originalsOpened} из ${result.totals.originalsRequested}.`,
    `Высокая уверенность: ${high.length}. Средняя уверенность: ${medium.length}. Ошибок открытия: ${errors.length}.`,
    '',
    '> Важно: техническое совпадение контента не подтверждает право собственности. Для жалобы нужны исходники, дата создания или договор с дизайнером.',
    ''
  ];
  for (const item of [...high, ...medium]) {
    lines.push(`## ${item.confidence}: ${item.originalSku} → ${item.candidateSku}`);
    lines.push(`- Исходная карточка: ${item.originalUrl}`);
    lines.push(`- Найденная карточка: ${item.candidateUrl}`);
    lines.push(`- Продавец: ${item.candidateSeller || 'не определён'}; страна: ${item.candidateCountry || 'не определена'}`);
    lines.push(`- Сходство названия: ${item.titleSimilarity}; описания: ${item.descriptionSimilarity}`);
    if (item.imageMatch) {
      lines.push(`- Изображение: exact=${item.imageMatch.exact}, dHash=${item.imageMatch.dhashDistance}, pixel=${item.imageMatch.pixelSimilarity}`);
      lines.push(`- Исходное изображение: ${item.imageMatch.originalImage}`);
      lines.push(`- Изображение кандидата: ${item.imageMatch.candidateImage}`);
    }
    lines.push('');
  }
  await fs.writeFile('output/summary.md', lines.join('\n'), 'utf8');

  console.log(JSON.stringify(result.totals));
}

main().catch(async (error) => {
  console.error(error?.stack || error);
  await fs.mkdir('output', { recursive: true });
  await fs.writeFile('output/fatal-error.txt', String(error?.stack || error), 'utf8');
  process.exitCode = 1;
});
