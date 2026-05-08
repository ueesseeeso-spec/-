/**
 * Ozon Unit Economics for Google Sheets
 * Версия: 1.0
 *
 * Что делает:
 * 1) Создает структуру Google Таблицы.
 * 2) Тянет товары, заказы FBO/FBS, финансовые транзакции и рекламу Ozon.
 * 3) Строит две юнит-экономики:
 *    - Юнит — Факт: по финансовым операциям Ozon.
 *    - Юнит — Прогноз: заказ = потенциальная продажа, с учетом отмен и прогнозных возвратов.
 * 4) Строит сравнение факт/прогноз и дашборд.
 *
 * Как начать:
 * 1) Вставь этот код в Extensions → Apps Script → Code.gs.
 * 2) Запусти setupWorkbook().
 * 3) На листе "Настройки" заполни seller_client_id_tmp и seller_api_key_tmp.
 * 4) Запусти saveCredentialsFromSettings().
 * 5) Заполни лист "Себестоимость".
 * 6) Запусти runFullSync().
 *
 * Важно:
 * - Seller API и Performance API используют разные ключи.
 * - Реклама через Performance API может быть нестабильной в Apps Script из-за асинхронных отчетов и лимитов.
 * - Если реклама через API не нужна, поставь use_performance_api = FALSE и заполняй лист API_Реклама вручную.
 */

const OZON_UNIT_VERSION = '1.0.0';

const SHEETS = Object.freeze({
  SETTINGS: 'Настройки',
  PERIOD: 'Период',
  COSTS: 'Себестоимость',
  PRODUCTS: 'API_Товары',
  ORDERS: 'API_Заказы',
  FINANCE: 'API_Финансы',
  ADS: 'API_Реклама',
  HISTORY: 'История по артикулам',
  UNIT_FACT: 'Юнит — Факт',
  UNIT_FORECAST: 'Юнит — Прогноз',
  COMPARE: 'Сравнение факт / прогноз',
  STOCK_VALUE: 'Остатки и потенциал',
  DASHBOARD: 'Дашборд',
  LOGS: 'Логи'
});

const HEADERS = Object.freeze({
  SETTINGS: ['Параметр', 'Значение', 'Комментарий'],
  PERIOD: ['Параметр', 'Значение', 'Подсказка'],
  COSTS: [
    'offer_id', 'sku', 'Ссылка_Ozon', 'Название', 'Себестоимость_шт', 'Упаковка_шт', 'Доп_расход_шт',
    'Налог_%_override', 'Целевая_маржа_%_override', 'Средний_%_отмен_override',
    'Средний_%_возвратов_override', 'Категория', 'Комментарий'
  ],
  PRODUCTS: [
    'product_id', 'offer_id', 'sku', 'Ссылка_Ozon', 'Название', 'Статус', 'Видимость', 'Цена', 'Маркет_цена',
    'Старая_цена', 'Валюта', 'Остаток_FBO', 'Остаток_FBS', 'Резерв', 'Обновлено'
  ],
  ORDERS: [
    'date', 'scheme', 'posting_number', 'order_id', 'status', 'offer_id', 'sku', 'product_id',
    'Название', 'quantity', 'price', 'total', 'commission_amount', 'payout', 'delivery_charge',
    'cancel_reason_id', 'in_process_at', 'shipment_date', 'raw_status'
  ],
  FINANCE: [
    'operation_date', 'posting_number', 'delivery_schema', 'operation_id', 'operation_type',
    'operation_type_name', 'type', 'sku', 'offer_id', 'Название', 'quantity_estimated',
    'accruals_for_sale', 'sale_commission', 'amount', 'delivery_charge', 'return_delivery_charge',
    'services_total', 'services_json', 'items_count', 'raw_json'
  ],
  ADS: [
    'date', 'campaign_id', 'campaign_name', 'adv_object_type', 'offer_id', 'sku', 'product_id',
    'Показы', 'Клики', 'Расход', 'Заказы_реклама', 'Выручка_реклама', 'Источник'
  ],
  HISTORY: [
    'offer_id', 'sku', 'Название', 'Заказано_история', 'Отменено_история', 'Продано_факт_оценка',
    'Возвраты_оценка', 'Средний_%_отмен', 'Средний_%_возвратов', 'Средняя_цена',
    'Средний_%_расходов_Ozon', 'Средний_%_комиссии', 'Средняя_логистика_шт',
    'Средний_расход_возврата_шт', 'Средний_ДРР', 'Обновлено'
  ],
  UNIT_FACT: [
    'offer_id', 'sku', 'Название', 'Продано_факт_оценка', 'Выручка_факт', 'Расходы_Ozon_факт',
    'Реклама_факт', 'Себестоимость', 'Упаковка', 'Доп_расходы', 'Налог', 'Чистая_прибыль',
    'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%', 'Прибыль_на_шт', 'Статус', 'Рекомендация'
  ],
  UNIT_FORECAST: [
    'offer_id', 'sku', 'Название', 'Заказано', 'Факт_отменено', 'Прогноз_итог_отмен',
    'Заказы_после_отмен', 'Средний_%_возвратов', 'Ожидаемые_возвраты', 'Ожидаемые_продажи',
    'Средняя_цена', 'Выручка_прогноз', 'Расходы_Ozon_прогноз', 'Реклама_факт', 'Себестоимость',
    'Упаковка', 'Доп_расходы', 'Налог', 'Чистая_прибыль_прогноз', 'Маржа_%', 'ДРР_%',
    'Расходы_Ozon_%', 'Прибыль_на_шт', 'Статус', 'Рекомендация'
  ],
  COMPARE: [
    'offer_id', 'sku', 'Название', 'Прибыль_факт', 'Прибыль_прогноз', 'Разница', 'Маржа_факт_%',
    'Маржа_прогноз_%', 'ДРР_факт_%', 'ДРР_прогноз_%', 'Вывод'
  ],
  STOCK_VALUE: [
    'Артикул', 'SKU Ozon', 'Название', 'Остаток FBO', 'Остаток FBS', 'Остаток всего',
    'Себестоимость за шт', 'Остаток по себесу', 'Текущая цена', 'Потенциальный оборот',
    'Прогнозная прибыль с остатка', 'Дней продаж в остатке', 'Статус остатка', 'Комментарий'
  ],
  DASHBOARD: ['Блок', 'Показатель', 'Факт', 'Прогноз', 'Комментарий'],
  LOGS: ['timestamp', 'level', 'scope', 'message']
});

const PROP_KEYS = Object.freeze({
  SELLER_CLIENT_ID: 'OZON_SELLER_CLIENT_ID',
  SELLER_API_KEY: 'OZON_SELLER_API_KEY',
  PERF_CLIENT_ID: 'OZON_PERFORMANCE_CLIENT_ID',
  PERF_CLIENT_SECRET: 'OZON_PERFORMANCE_CLIENT_SECRET',
  PERF_ACCESS_TOKEN: 'OZON_PERFORMANCE_ACCESS_TOKEN',
  PERF_TOKEN_EXPIRES_AT: 'OZON_PERFORMANCE_TOKEN_EXPIRES_AT'
});

const OZON_SELLER_BASE = 'https://api-seller.ozon.ru';
// Новый домен Performance API. Старый performance.ozon.ru может отдавать HTTP 307 Temporary Redirect.
const OZON_PERF_BASE = 'https://api-performance.ozon.ru';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ozon Unit')
    .addItem('1. Создать/обновить структуру таблицы', 'setupWorkbook')
    .addItem('2. Сохранить API-ключи из Настроек', 'saveCredentialsFromSettings')
    .addItem('Открыть выбор периода', 'goToPeriodSettings')
    .addSeparator()
    .addItem('3. Полная синхронизация и пересчет', 'runFullSync')
    .addItem('Только товары', 'syncProducts')
    .addItem('Только заказы', 'syncOrders')
    .addItem('Только финансы', 'syncFinanceTransactions')
    .addItem('Только реклама', 'syncAds')
    .addSeparator()
    .addItem('Пересчитать аналитику', 'rebuildAnalytics')
    .addItem('Создать ежедневное автообновление', 'createDailyTrigger')
    .addSeparator()
    .addItem('Скрыть технические листы', 'hideTechnicalSheets')
    .addItem('Показать все листы', 'showAllSheets')
    .addToUi();
}

function setupWorkbook() {
  ensureSheetWithHeader(SHEETS.SETTINGS, HEADERS.SETTINGS);
  setupPeriodSheet();
  ensureSheetWithHeader(SHEETS.COSTS, HEADERS.COSTS);
  ensureSheetWithHeader(SHEETS.PRODUCTS, HEADERS.PRODUCTS);
  ensureSheetWithHeader(SHEETS.ORDERS, HEADERS.ORDERS);
  ensureSheetWithHeader(SHEETS.FINANCE, HEADERS.FINANCE);
  ensureSheetWithHeader(SHEETS.ADS, HEADERS.ADS);
  ensureSheetWithHeader(SHEETS.HISTORY, HEADERS.HISTORY);
  ensureSheetWithHeader(SHEETS.UNIT_FACT, HEADERS.UNIT_FACT);
  ensureSheetWithHeader(SHEETS.UNIT_FORECAST, HEADERS.UNIT_FORECAST);
  ensureSheetWithHeader(SHEETS.COMPARE, HEADERS.COMPARE);
  ensureSheetWithHeader(SHEETS.STOCK_VALUE, HEADERS.STOCK_VALUE);
  ensureSheetWithHeader(SHEETS.DASHBOARD, HEADERS.DASHBOARD);
  ensureSheetWithHeader(SHEETS.LOGS, HEADERS.LOGS);

  seedSettings();
  formatAllSheets();
  organizeSheetsForUser();
  logInfo('setupWorkbook', 'Структура таблицы создана/обновлена');
}

function goToPeriodSettings() {
  setupPeriodSheet();
  const sh = getSheet(SHEETS.PERIOD);
  SpreadsheetApp.getActive().setActiveSheet(sh);
  sh.getRange('B4').activate();
}

function saveCredentialsFromSettings() {
  const settings = getSettingsMap();
  const props = PropertiesService.getScriptProperties();

  const sellerClientId = String(settings.seller_client_id_tmp || '').trim();
  const sellerApiKey = String(settings.seller_api_key_tmp || '').trim();
  const perfClientId = String(settings.performance_client_id_tmp || '').trim();
  const perfSecret = String(settings.performance_client_secret_tmp || '').trim();

  if (sellerClientId) props.setProperty(PROP_KEYS.SELLER_CLIENT_ID, sellerClientId);
  if (sellerApiKey) props.setProperty(PROP_KEYS.SELLER_API_KEY, sellerApiKey);
  if (perfClientId) props.setProperty(PROP_KEYS.PERF_CLIENT_ID, perfClientId);
  if (perfSecret) props.setProperty(PROP_KEYS.PERF_CLIENT_SECRET, perfSecret);

  clearTemporaryCredentialCells();
  logInfo('saveCredentialsFromSettings', 'Ключи сохранены в Script Properties, временные ячейки очищены');
  SpreadsheetApp.getUi().alert('Ключи сохранены. Временные ячейки в листе Настройки очищены.');
}

function runFullSync() {
  setupWorkbook();
  validateSelectedPeriod();

  syncProducts();
  syncOrders();
  syncFinanceTransactions();
  syncAds();
  rebuildAnalytics();

  setSettingValue('last_sync', new Date());
  logInfo('runFullSync', 'Полная синхронизация и пересчет завершены');
}

function rebuildAnalytics() {
  rebuildHistory();
  rebuildUnitFact();
  rebuildUnitForecast();
  rebuildComparison();
  rebuildStockValuation();
  rebuildDashboard();
  applyConditionalFormattingToUnitSheets();
  logInfo('rebuildAnalytics', 'Аналитика пересчитана');
}

function syncProducts() {
  ensureSheetWithHeader(SHEETS.PRODUCTS, HEADERS.PRODUCTS);
  requireSellerCredentials();

  const listItems = fetchProductList();
  const productIds = unique(listItems.map(x => safeString(x.product_id || x.id)).filter(Boolean));
  const offerIds = unique(listItems.map(x => safeString(x.offer_id)).filter(Boolean));

  const detailsByProductId = fetchProductDetails(productIds, offerIds);
  const priceByProductId = fetchProductPrices(productIds);
  const stockByProductId = fetchProductStocks(productIds);

  const rows = [];
  const now = new Date();

  listItems.forEach(item => {
    const productId = safeString(item.product_id || item.id);
    const detail = detailsByProductId[productId] || {};
    const price = priceByProductId[productId] || {};
    const stock = stockByProductId[productId] || {};

    const offerId = safeString(detail.offer_id || item.offer_id);
    const sku = safeString(detail.sku || item.sku || firstNumber(detail.sources, 'sku'));
    const name = safeString(detail.name || item.name || item.offer_id || offerId);
    const hasPrice = Object.keys(price).length > 0;
    const hasStock = Object.keys(stock).length > 0;

    rows.push({
      product_id: productId,
      offer_id: offerId,
      sku: sku,
      'Ссылка_Ozon': buildOzonProductUrl(sku, productId),
      'Название': name,
      'Статус': safeString(detail.status && (detail.status.state_name || detail.status.name) || item.status || item.state),
      'Видимость': safeString(detail.visibility || item.visibility),
      'Цена': hasPrice ? pickNumOrBlank(price.price, detail.price, item.price) : '',
      'Маркет_цена': hasPrice ? pickNumOrBlank(price.marketing_price, detail.marketing_price) : '',
      'Старая_цена': hasPrice ? pickNumOrBlank(price.old_price, detail.old_price) : '',
      'Валюта': safeString(price.currency_code || detail.currency_code || 'RUB'),
      'Остаток_FBO': hasStock ? num(stock.fbo_stock) : '',
      'Остаток_FBS': hasStock ? num(stock.fbs_stock) : '',
      'Резерв': hasStock ? num(stock.reserved) : '',
      'Обновлено': now
    });
  });

  writeObjects(SHEETS.PRODUCTS, HEADERS.PRODUCTS, rows);
  addMissingProductsToCosts(rows);
  beautifyWorkbook();
  logInfo('syncProducts', 'Загружено товаров: ' + rows.length);
}

function syncOrders() {
  ensureSheetWithHeader(SHEETS.ORDERS, HEADERS.ORDERS);
  requireSellerCredentials();

  const period = getSyncPeriodWithHistory();
  const productMaps = getProductMaps();
  const orderProductMap = getOrderProductMap();
  const rows = [];

  rows.push.apply(rows, fetchFboPostings(period.from, period.to, productMaps));
  rows.push.apply(rows, fetchFbsPostings(period.from, period.to, productMaps));

  writeObjects(SHEETS.ORDERS, HEADERS.ORDERS, rows);
  logInfo('syncOrders', 'Загружено строк заказов: ' + rows.length);
}

function syncFinanceTransactions() {
  ensureSheetWithHeader(SHEETS.FINANCE, HEADERS.FINANCE);
  requireSellerCredentials();

  const period = getSyncPeriodWithHistory();

  // Finance API Ozon ограничивает период одного запроса.
  // Дробим короткими окнами по 28 дней, чтобы не ловить ошибку
  // "too long period, only one month allowed" из-за часовых поясов и месяцев на 31 день.
  const chunks = splitDateRangeByDays(period.from, period.to, 28);

  const productMaps = getProductMaps();

  // Дополнительная связка нужна, потому что в финансовых операциях Ozon не всегда есть offer_id.
  // Тогда определяем артикул через posting_number + sku из листа API_Заказы.
  const orderProductMap = getOrderProductMap();

  const rows = [];

  chunks.forEach(chunk => {
    let page = 1;
    const pageSize = 1000;
    while (true) {
      const payload = {
        filter: {
          date: {
            from: toOzonDateTime(chunk.from, false),
            to: toOzonDateTime(chunk.to, true)
          }
        },
        page: page,
        page_size: pageSize
      };
      const data = ozonSellerPost('/v3/finance/transaction/list', payload, 'finance page ' + page);
      const operations = safeArray(data && data.result && data.result.operations);
      operations.forEach(op => {
        rows.push.apply(rows, normalizeFinanceOperation(op, productMaps, orderProductMap));
      });
      if (operations.length < pageSize) break;
      page += 1;
      Utilities.sleep(250);
    }
  });

  writeObjects(SHEETS.FINANCE, HEADERS.FINANCE, rows);
  logInfo('syncFinanceTransactions', 'Загружено финансовых строк: ' + rows.length);
}

function syncAds() {
  ensureSheetWithHeader(SHEETS.ADS, HEADERS.ADS);
  const settings = getSettingsMap();
  const useApi = parseBool(settings.use_performance_api);

  if (!useApi) {
    logInfo('syncAds', 'Performance API отключен. Лист API_Реклама оставлен для ручного заполнения.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty(PROP_KEYS.PERF_CLIENT_ID);
  const secret = props.getProperty(PROP_KEYS.PERF_CLIENT_SECRET);
  if (!clientId || !secret) {
    logWarn('syncAds', 'Нет ключей Performance API. Заполни performance_client_id_tmp и performance_client_secret_tmp, затем запусти saveCredentialsFromSettings().');
    return;
  }

  const period = getSelectedPeriod();
  const token = getPerformanceToken();
  const campaigns = fetchPerformanceCampaigns(token);
  const productMaps = getProductMaps();
  const rows = [];

  const campaignIds = campaigns.map(c => safeString(c.id)).filter(Boolean);
  const campaignById = {};
  campaigns.forEach(c => campaignById[safeString(c.id)] = c);

  chunkArray(campaignIds, 10).forEach(ids => {
    const uuid = requestPerformanceStatsJson(token, ids, dateToYmd(period.from), dateToYmd(period.to));
    if (!uuid) return;
    const reportJson = downloadPerformanceReport(token, uuid);
    if (!reportJson) return;
    const reportRows = parsePerformanceReport(reportJson, campaignById, productMaps);
    rows.push.apply(rows, reportRows);
    Utilities.sleep(1000);
  });

  if (rows.length) {
    writeObjects(SHEETS.ADS, HEADERS.ADS, rows);
  }
  logInfo('syncAds', 'Загружено строк рекламы: ' + rows.length);
}

function rebuildHistory() {
  ensureSheetWithHeader(SHEETS.HISTORY, HEADERS.HISTORY);

  const period = getSyncPeriodWithHistory();
  const products = getProductMaps();
  const orders = filterRowsByDate(readSheetAsObjects(SHEETS.ORDERS), 'date', period.from, period.to);
  const finance = filterRowsByDate(readSheetAsObjects(SHEETS.FINANCE), 'operation_date', period.from, period.to);
  const ads = filterRowsByDate(readSheetAsObjects(SHEETS.ADS), 'date', period.from, period.to);

  const agg = {};
  const touch = offerId => {
    const key = safeString(offerId || 'NO_OFFER');
    if (!agg[key]) {
      const p = products.byOffer[key] || {};
      agg[key] = {
        offer_id: key,
        sku: safeString(p.sku),
        name: safeString(p.name),
        ordered: 0,
        canceled: 0,
        orderTotal: 0,
        soldQty: 0,
        returnQty: 0,
        revenue: 0,
        commission: 0,
        ozonExpenses: 0,
        logistics: 0,
        returnExpense: 0,
        adsSpend: 0
      };
    }
    return agg[key];
  };

  orders.forEach(r => {
    const a = touch(r.offer_id);
    const qty = num(r.quantity);
    a.ordered += qty;
    a.orderTotal += num(r.total);
    if (isCancelledStatus(r.status)) a.canceled += qty;
  });

  finance.forEach(r => {
    const a = touch(r.offer_id);
    const metrics = financeRowMetrics(r);
    if (metrics.qty > 0) a.soldQty += metrics.qty;
    if (metrics.isReturn) a.returnQty += Math.abs(metrics.qty || 1);
    a.revenue += Math.max(0, metrics.revenue);
    a.commission += metrics.commission;
    a.ozonExpenses += metrics.ozonExpenses;
    a.logistics += metrics.logistics;
    a.returnExpense += metrics.returnExpense;
  });

  ads.forEach(r => {
    const offerId = resolveOfferIdFromRow(r, products);
    const a = touch(offerId);
    a.adsSpend += num(r['Расход']);
  });

  const now = new Date();
  const rows = Object.keys(agg).sort().map(key => {
    const a = agg[key];
    const afterCancel = Math.max(0, a.ordered - a.canceled);
    const avgPrice = a.ordered ? a.orderTotal / a.ordered : safeDiv(a.revenue, Math.max(1, a.soldQty));
    return {
      offer_id: a.offer_id,
      sku: a.sku,
      'Название': a.name,
      'Заказано_история': a.ordered,
      'Отменено_история': a.canceled,
      'Продано_факт_оценка': a.soldQty,
      'Возвраты_оценка': a.returnQty,
      'Средний_%_отмен': safeDiv(a.canceled, a.ordered),
      'Средний_%_возвратов': safeDiv(a.returnQty, afterCancel || a.soldQty),
      'Средняя_цена': avgPrice,
      'Средний_%_расходов_Ozon': safeDiv(a.ozonExpenses, a.revenue),
      'Средний_%_комиссии': safeDiv(a.commission, a.revenue),
      'Средняя_логистика_шт': safeDiv(a.logistics, Math.max(1, a.soldQty)),
      'Средний_расход_возврата_шт': safeDiv(a.returnExpense, Math.max(1, a.returnQty)),
      'Средний_ДРР': safeDiv(a.adsSpend, a.revenue),
      'Обновлено': now
    };
  });

  writeObjects(SHEETS.HISTORY, HEADERS.HISTORY, rows);
  logInfo('rebuildHistory', 'История пересчитана: ' + rows.length + ' артикулов');
}

function rebuildUnitFact() {
  ensureSheetWithHeader(SHEETS.UNIT_FACT, HEADERS.UNIT_FACT);

  const settings = getSettingsMap();
  const period = getSelectedPeriod();
  const products = getProductMaps();
  const costs = getCostsMap();
  const finance = filterRowsByDate(readSheetAsObjects(SHEETS.FINANCE), 'operation_date', period.from, period.to);
  const ads = filterRowsByDate(readSheetAsObjects(SHEETS.ADS), 'date', period.from, period.to);

  const agg = {};
  const touch = offerId => {
    const key = safeString(offerId || 'NO_OFFER');
    if (!agg[key]) {
      const p = products.byOffer[key] || {};
      agg[key] = {
        offer_id: key,
        sku: safeString(p.sku),
        name: safeString(p.name),
        qty: 0,
        revenue: 0,
        ozonExpenses: 0,
        ads: 0
      };
    }
    return agg[key];
  };

  finance.forEach(r => {
    const offerId = resolveOfferIdFromRow(r, products);
    const a = touch(offerId);
    const metrics = financeRowMetrics(r);
    a.qty += metrics.qty;
    a.revenue += metrics.revenue;
    a.ozonExpenses += metrics.ozonExpenses;
  });

  allocateAdsToAggregation(ads, products, agg, touch, a => Math.max(0, num(a.revenue)));


  const rows = Object.keys(agg).sort().map(offerId => {
    const a = agg[offerId];
    const c = costs[offerId] || {};
    const qty = Math.max(0, a.qty);
    const cost = qty * num(c.cost_per_unit);
    const pack = qty * num(c.package_per_unit);
    const additional = qty * num(c.additional_per_unit);
    const taxRate = getTaxRateForOffer(settings, c);
    const tax = calcTax(settings, taxRate, a.revenue, a.ozonExpenses, a.ads, cost, pack, additional);
    const profit = a.revenue - a.ozonExpenses - a.ads - cost - pack - additional - tax;
    const margin = safeDiv(profit, a.revenue);
    const adShare = safeDiv(a.ads, a.revenue);
    const ozonShare = safeDiv(a.ozonExpenses, a.revenue);
    const targetMargin = getTargetMarginForOffer(settings, c);

    return {
      offer_id: a.offer_id,
      sku: a.sku || safeString(c.sku),
      'Название': a.name || safeString(c.name),
      'Продано_факт_оценка': qty,
      'Выручка_факт': a.revenue,
      'Расходы_Ozon_факт': a.ozonExpenses,
      'Реклама_факт': a.ads,
      'Себестоимость': cost,
      'Упаковка': pack,
      'Доп_расходы': additional,
      'Налог': tax,
      'Чистая_прибыль': profit,
      'Маржа_%': margin,
      'ДРР_%': adShare,
      'Расходы_Ozon_%': ozonShare,
      'Прибыль_на_шт': safeDiv(profit, qty),
      'Статус': buildStatus(profit, margin, adShare, ozonShare, settings, targetMargin),
      'Рекомендация': buildRecommendation(profit, margin, adShare, ozonShare, settings, targetMargin, 'fact')
    };
  });

  writeObjects(SHEETS.UNIT_FACT, HEADERS.UNIT_FACT, rows);
  formatPercentColumns(SHEETS.UNIT_FACT, ['Маржа_%', 'ДРР_%', 'Расходы_Ozon_%']);
  logInfo('rebuildUnitFact', 'Факт пересчитан: ' + rows.length + ' артикулов');
}

function rebuildUnitForecast() {
  ensureSheetWithHeader(SHEETS.UNIT_FORECAST, HEADERS.UNIT_FORECAST);

  const settings = getSettingsMap();
  const period = getSelectedPeriod();
  const products = getProductMaps();
  const costs = getCostsMap();
  const history = getHistoryMap();
  const orders = filterRowsByDate(readSheetAsObjects(SHEETS.ORDERS), 'date', period.from, period.to);
  const ads = filterRowsByDate(readSheetAsObjects(SHEETS.ADS), 'date', period.from, period.to);

  const agg = {};
  const touch = offerId => {
    const key = safeString(offerId || 'NO_OFFER');
    if (!agg[key]) {
      const p = products.byOffer[key] || {};
      agg[key] = {
        offer_id: key,
        sku: safeString(p.sku),
        name: safeString(p.name),
        ordered: 0,
        canceled: 0,
        notCanceledTotal: 0,
        notCanceledQty: 0,
        ads: 0
      };
    }
    return agg[key];
  };

  orders.forEach(r => {
    const offerId = resolveOfferIdFromRow(r, products);
    const a = touch(offerId);
    const qty = num(r.quantity);
    a.ordered += qty;
    if (isCancelledStatus(r.status)) {
      a.canceled += qty;
    } else {
      a.notCanceledTotal += num(r.total);
      a.notCanceledQty += qty;
    }
  });

  allocateAdsToAggregation(ads, products, agg, touch, a => Math.max(0, num(a.notCanceledTotal) || num(a.ordered)));


  const rows = Object.keys(agg).sort().map(offerId => {
    const a = agg[offerId];
    const c = costs[offerId] || {};
    const h = history[offerId] || {};

    const avgCancelRate = getRateOverrideOrHistory(c.cancel_rate_override, h.avg_cancel_rate, settings.default_cancel_rate);
    const avgReturnRate = getRateOverrideOrHistory(c.return_rate_override, h.avg_return_rate, settings.default_return_rate);

    const expectedFinalCanceled = Math.max(a.canceled, a.ordered * avgCancelRate);
    const afterCancel = Math.max(0, a.ordered - expectedFinalCanceled);
    const expectedReturns = afterCancel * avgReturnRate;
    const expectedSales = Math.max(0, afterCancel - expectedReturns);

    const avgPrice = safeDiv(a.notCanceledTotal, a.notCanceledQty) || num(h.avg_price) || num(settings.default_avg_price);
    const revenue = expectedSales * avgPrice;

    const ozonExpensePct = num(h.avg_ozon_expense_pct) || num(settings.default_ozon_expense_pct);
    const returnExpensePerItem = num(h.avg_return_expense_per_item) || num(settings.default_return_expense_per_item);
    const ozonExpenses = revenue * ozonExpensePct + expectedReturns * returnExpensePerItem;

    const cost = expectedSales * num(c.cost_per_unit);
    const pack = expectedSales * num(c.package_per_unit);
    const additional = expectedSales * num(c.additional_per_unit);
    const taxRate = getTaxRateForOffer(settings, c);
    const tax = calcTax(settings, taxRate, revenue, ozonExpenses, a.ads, cost, pack, additional);
    const profit = revenue - ozonExpenses - a.ads - cost - pack - additional - tax;
    const margin = safeDiv(profit, revenue);
    const adShare = safeDiv(a.ads, revenue);
    const ozonShare = safeDiv(ozonExpenses, revenue);
    const targetMargin = getTargetMarginForOffer(settings, c);

    return {
      offer_id: a.offer_id,
      sku: a.sku || safeString(c.sku),
      'Название': a.name || safeString(c.name),
      'Заказано': a.ordered,
      'Факт_отменено': a.canceled,
      'Прогноз_итог_отмен': expectedFinalCanceled,
      'Заказы_после_отмен': afterCancel,
      'Средний_%_возвратов': avgReturnRate,
      'Ожидаемые_возвраты': expectedReturns,
      'Ожидаемые_продажи': expectedSales,
      'Средняя_цена': avgPrice,
      'Выручка_прогноз': revenue,
      'Расходы_Ozon_прогноз': ozonExpenses,
      'Реклама_факт': a.ads,
      'Себестоимость': cost,
      'Упаковка': pack,
      'Доп_расходы': additional,
      'Налог': tax,
      'Чистая_прибыль_прогноз': profit,
      'Маржа_%': margin,
      'ДРР_%': adShare,
      'Расходы_Ozon_%': ozonShare,
      'Прибыль_на_шт': safeDiv(profit, expectedSales),
      'Статус': buildStatus(profit, margin, adShare, ozonShare, settings, targetMargin),
      'Рекомендация': buildRecommendation(profit, margin, adShare, ozonShare, settings, targetMargin, 'forecast')
    };
  });

  writeObjects(SHEETS.UNIT_FORECAST, HEADERS.UNIT_FORECAST, rows);
  formatPercentColumns(SHEETS.UNIT_FORECAST, ['Средний_%_возвратов', 'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%']);
  logInfo('rebuildUnitForecast', 'Прогноз пересчитан: ' + rows.length + ' артикулов');
}

function rebuildComparison() {
  ensureSheetWithHeader(SHEETS.COMPARE, HEADERS.COMPARE);

  const fact = readSheetAsObjects(SHEETS.UNIT_FACT);
  const forecast = readSheetAsObjects(SHEETS.UNIT_FORECAST);
  const byOffer = {};

  fact.forEach(r => {
    const key = safeString(r.offer_id || 'NO_OFFER');
    byOffer[key] = byOffer[key] || { offer_id: key };
    byOffer[key].sku = safeString(r.sku);
    byOffer[key].name = safeString(r['Название']);
    byOffer[key].factProfit = num(r['Чистая_прибыль']);
    byOffer[key].factMargin = num(r['Маржа_%']);
    byOffer[key].factDrr = num(r['ДРР_%']);
  });

  forecast.forEach(r => {
    const key = safeString(r.offer_id || 'NO_OFFER');
    byOffer[key] = byOffer[key] || { offer_id: key };
    byOffer[key].sku = byOffer[key].sku || safeString(r.sku);
    byOffer[key].name = byOffer[key].name || safeString(r['Название']);
    byOffer[key].forecastProfit = num(r['Чистая_прибыль_прогноз']);
    byOffer[key].forecastMargin = num(r['Маржа_%']);
    byOffer[key].forecastDrr = num(r['ДРР_%']);
  });

  const rows = Object.keys(byOffer).sort().map(k => {
    const x = byOffer[k];
    const diff = num(x.forecastProfit) - num(x.factProfit);
    return {
      offer_id: x.offer_id,
      sku: x.sku,
      'Название': x.name,
      'Прибыль_факт': num(x.factProfit),
      'Прибыль_прогноз': num(x.forecastProfit),
      'Разница': diff,
      'Маржа_факт_%': num(x.factMargin),
      'Маржа_прогноз_%': num(x.forecastMargin),
      'ДРР_факт_%': num(x.factDrr),
      'ДРР_прогноз_%': num(x.forecastDrr),
      'Вывод': compareConclusion(num(x.factProfit), num(x.forecastProfit))
    };
  });

  writeObjects(SHEETS.COMPARE, HEADERS.COMPARE, rows);
  formatPercentColumns(SHEETS.COMPARE, ['Маржа_факт_%', 'Маржа_прогноз_%', 'ДРР_факт_%', 'ДРР_прогноз_%']);
  logInfo('rebuildComparison', 'Сравнение пересчитано: ' + rows.length + ' артикулов');
}

function rebuildStockValuation() {
  ensureSheetWithHeader(SHEETS.STOCK_VALUE, HEADERS.STOCK_VALUE);

  const settings = getSettingsMap();
  const products = readSheetAsObjects(SHEETS.PRODUCTS);
  const costs = getCostsMap();
  const history = getHistoryMap();

  const rows = products.map(p => {
    const offerId = safeString(p.offer_id);
    const c = costs[offerId] || {};
    const h = history[offerId] || {};
    const fbo = num(p['Остаток_FBO']);
    const fbs = num(p['Остаток_FBS']);
    const stock = fbo + fbs;
    const costPerUnit = num(c.cost_per_unit);
    const pack = num(c.package_per_unit);
    const add = num(c.additional_per_unit);
    const price = normalizeOzonMoney(p['Цена']) || num(h.avg_price);
    const ozonPct = num(h.avg_ozon_expense_pct) || num(settings.default_ozon_expense_pct);
    const taxRate = getTaxRateForOffer(settings, c);
    const revenue = stock * price;
    const costValue = stock * costPerUnit;
    const ozonExpenses = revenue * ozonPct;
    const tax = calcTax(settings, taxRate, revenue, ozonExpenses, 0, costValue, stock * pack, stock * add);
    const profit = revenue - ozonExpenses - costValue - stock * pack - stock * add - tax;
    const soldDaily = estimateDailySales(offerId);
    const daysLeft = soldDaily > 0 ? stock / soldDaily : '';

    return {
      'Артикул': offerId,
      'SKU Ozon': safeString(p.sku),
      'Название': safeString(p['Название']),
      'Остаток FBO': fbo,
      'Остаток FBS': fbs,
      'Остаток всего': stock,
      'Себестоимость за шт': costPerUnit,
      'Остаток по себесу': costValue,
      'Текущая цена': price,
      'Потенциальный оборот': revenue,
      'Прогнозная прибыль с остатка': profit,
      'Дней продаж в остатке': daysLeft,
      'Статус остатка': stockStatus(stock, daysLeft),
      'Комментарий': stockComment(stock, daysLeft, costPerUnit, price)
    };
  }).filter(r => safeString(r['Артикул']) || safeString(r['Название']));

  writeObjects(SHEETS.STOCK_VALUE, HEADERS.STOCK_VALUE, rows);
  styleStockValuationSheet();
  logInfo('rebuildStockValuation', 'Остатки и потенциал пересчитаны: ' + rows.length + ' артикулов');
}

function rebuildDashboard() {
  ensureSheetWithHeader(SHEETS.DASHBOARD, HEADERS.DASHBOARD);

  const fact = readSheetAsObjects(SHEETS.UNIT_FACT);
  const forecast = readSheetAsObjects(SHEETS.UNIT_FORECAST);
  const settings = getSettingsMap();

  const factKpi = calcUnitKpi(fact, 'fact');
  const forecastKpi = calcUnitKpi(forecast, 'forecast');
  const rows = [];

  rows.push({ 'Блок': 'KPI', 'Показатель': 'Оборот', 'Факт': factKpi.revenue, 'Прогноз': forecastKpi.revenue, 'Комментарий': '' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'Чистая прибыль', 'Факт': factKpi.profit, 'Прогноз': forecastKpi.profit, 'Комментарий': '' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'Маржа', 'Факт': factKpi.margin, 'Прогноз': forecastKpi.margin, 'Комментарий': 'Доля чистой прибыли в обороте' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'ДРР', 'Факт': factKpi.adShare, 'Прогноз': forecastKpi.adShare, 'Комментарий': 'Реклама / выручка' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'Расходы Ozon', 'Факт': factKpi.ozonShare, 'Прогноз': forecastKpi.ozonShare, 'Комментарий': 'Комиссии, логистика и услуги Ozon / выручка' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'Товаров в минус', 'Факт': factKpi.negativeCount, 'Прогноз': forecastKpi.negativeCount, 'Комментарий': '' });
  rows.push({ 'Блок': 'KPI', 'Показатель': 'Товаров с высоким ДРР', 'Факт': factKpi.highDrrCount, 'Прогноз': forecastKpi.highDrrCount, 'Комментарий': 'Порог: ' + toPercent(num(settings.critical_ad_share_pct)) });
  rows.push({ 'Блок': '', 'Показатель': '', 'Факт': '', 'Прогноз': '', 'Комментарий': '' });

  topRows(forecast, 'Чистая_прибыль_прогноз', 10).forEach((r, idx) => {
    rows.push({
      'Блок': 'Топ прибыль прогноз',
      'Показатель': (idx + 1) + '. ' + safeString(r.offer_id),
      'Факт': '',
      'Прогноз': num(r['Чистая_прибыль_прогноз']),
      'Комментарий': safeString(r['Название'])
    });
  });
  rows.push({ 'Блок': '', 'Показатель': '', 'Факт': '', 'Прогноз': '', 'Комментарий': '' });

  topRows(forecast, 'Выручка_прогноз', 10).forEach((r, idx) => {
    rows.push({
      'Блок': 'Топ оборот прогноз',
      'Показатель': (idx + 1) + '. ' + safeString(r.offer_id),
      'Факт': '',
      'Прогноз': num(r['Выручка_прогноз']),
      'Комментарий': safeString(r['Название'])
    });
  });
  rows.push({ 'Блок': '', 'Показатель': '', 'Факт': '', 'Прогноз': '', 'Комментарий': '' });

  bottomRows(forecast, 'Чистая_прибыль_прогноз', 10).forEach((r, idx) => {
    rows.push({
      'Блок': 'Проблемные товары',
      'Показатель': (idx + 1) + '. ' + safeString(r.offer_id),
      'Факт': '',
      'Прогноз': num(r['Чистая_прибыль_прогноз']),
      'Комментарий': safeString(r['Рекомендация'])
    });
  });

  writeObjects(SHEETS.DASHBOARD, HEADERS.DASHBOARD, rows);
  formatDashboard();
  logInfo('rebuildDashboard', 'Дашборд пересчитан');
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFullSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFullSync')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  SpreadsheetApp.getUi().alert('Ежедневное автообновление создано: примерно в 06:00 по часовому поясу проекта.');
}

function organizeSheetsForUser() {
  const settings = getSettingsMap();
  const usePerformanceApi = parseBool(settings.use_performance_api);

  const visibleSheets = [
    SHEETS.PERIOD,
    SHEETS.DASHBOARD,
    SHEETS.UNIT_FORECAST,
    SHEETS.UNIT_FACT,
    SHEETS.STOCK_VALUE,
    SHEETS.COSTS
  ];

  if (!usePerformanceApi) visibleSheets.push(SHEETS.ADS);

  const technicalSheets = [
    SHEETS.SETTINGS,
    SHEETS.PRODUCTS,
    SHEETS.ORDERS,
    SHEETS.FINANCE,
    SHEETS.HISTORY,
    SHEETS.COMPARE,
    SHEETS.LOGS
  ];

  const ss = SpreadsheetApp.getActive();
  visibleSheets.forEach(name => {
    const sh = getSheet(name);
    if (sh) sh.showSheet();
  });

  reorderSheets(visibleSheets.concat(technicalSheets));
  const startSheet = getSheet(SHEETS.PERIOD) || getSheet(SHEETS.DASHBOARD) || ss.getSheets()[0];
  if (startSheet) ss.setActiveSheet(startSheet);

  technicalSheets.forEach(name => {
    const sh = getSheet(name);
    if (sh) sh.hideSheet();
  });
}

function hideTechnicalSheets() {
  organizeSheetsForUser();
  SpreadsheetApp.getUi().alert('Технические листы скрыты. Рабочие листы оставлены: Период, Дашборд, Юнит — Прогноз, Юнит — Факт, Себестоимость.');
}

function showAllSheets() {
  Object.keys(SHEETS).forEach(k => {
    const sh = getSheet(SHEETS[k]);
    if (sh) sh.showSheet();
  });
  reorderSheets([
    SHEETS.PERIOD,
    SHEETS.DASHBOARD,
    SHEETS.UNIT_FORECAST,
    SHEETS.UNIT_FACT,
    SHEETS.STOCK_VALUE,
    SHEETS.COMPARE,
    SHEETS.COSTS,
    SHEETS.ADS,
    SHEETS.SETTINGS,
    SHEETS.PRODUCTS,
    SHEETS.ORDERS,
    SHEETS.FINANCE,
    SHEETS.HISTORY,
    SHEETS.LOGS
  ]);
  SpreadsheetApp.getUi().alert('Все листы показаны. Можно смотреть технические данные и логи.');
}

function reorderSheets(names) {
  const ss = SpreadsheetApp.getActive();
  let position = 1;
  names.forEach(name => {
    const sh = getSheet(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(position);
    position += 1;
  });
}

function fetchProductList() {
  let items = [];
  let lastId = '';
  while (true) {
    const payload = {
      filter: { visibility: 'ALL' },
      last_id: lastId,
      limit: 1000
    };
    let data;
    try {
      data = ozonSellerPost('/v3/product/list', payload, 'product list v3');
    } catch (e) {
      logWarn('fetchProductList', 'v3/product/list не сработал, пробую v2/product/list: ' + e.message);
      data = ozonSellerPost('/v2/product/list', payload, 'product list v2');
    }
    const result = data.result || {};
    const part = safeArray(result.items);
    items = items.concat(part);
    lastId = safeString(result.last_id);
    if (!lastId || part.length === 0) break;
    Utilities.sleep(250);
  }
  return items;
}

function fetchProductDetails(productIds, offerIds) {
  const map = {};
  const chunks = chunkArray(productIds, 1000);
  chunks.forEach(ids => {
    let data;
    try {
      data = ozonSellerPost('/v3/product/info/list', { product_id: ids.map(String) }, 'product info v3');
    } catch (e) {
      logWarn('fetchProductDetails', 'v3/product/info/list не сработал, пробую v2/product/info/list: ' + e.message);
      data = ozonSellerPost('/v2/product/info/list', { product_id: ids.map(String) }, 'product info v2');
    }
    const items = safeArray(data && data.result && (data.result.items || data.result));
    items.forEach(it => {
      const id = safeString(it.id || it.product_id);
      if (id) map[id] = it;
    });
    Utilities.sleep(250);
  });

  if (Object.keys(map).length === 0 && offerIds.length) {
    chunkArray(offerIds, 1000).forEach(ids => {
      try {
        const data = ozonSellerPost('/v3/product/info/list', { offer_id: ids }, 'product info by offer_id');
        const items = safeArray(data && data.result && (data.result.items || data.result));
        items.forEach(it => {
          const id = safeString(it.id || it.product_id);
          if (id) map[id] = it;
        });
      } catch (e) {
        logWarn('fetchProductDetails', 'Не удалось получить детали по offer_id: ' + e.message);
      }
      Utilities.sleep(250);
    });
  }
  return map;
}

function fetchProductPrices(productIds) {
  const map = {};
  if (!productIds.length) return map;

  try {
    let cursor = '';
    while (true) {
      const data = ozonSellerPost('/v5/product/info/prices', {
        cursor: cursor,
        filter: { product_id: productIds.map(String), visibility: 'ALL' },
        limit: 1000
      }, 'product prices v5');
      const result = data.result || {};
      const items = safeArray(result.items);
      items.forEach(it => {
        const id = safeString(it.product_id || it.id);
        if (id) map[id] = normalizePriceItem(it);
      });
      cursor = safeString(result.cursor || result.next_cursor);
      if (!cursor || items.length === 0) break;
      Utilities.sleep(250);
    }
  } catch (e) {
    logWarn('fetchProductPrices', 'v5/product/info/prices не сработал, пробую v4: ' + e.message);
  }

  const missingIds = productIds.filter(id => !map[safeString(id)]);
  if (missingIds.length) {
    let lastId = '';
    while (true) {
      try {
        const data = ozonSellerPost('/v4/product/info/prices', {
          filter: { product_id: missingIds.map(String), visibility: 'ALL' },
          last_id: lastId,
          limit: 1000
        }, 'product prices v4');
        const result = data.result || {};
        const items = safeArray(result.items);
        items.forEach(it => {
          const id = safeString(it.product_id || it.id);
          if (id) map[id] = normalizePriceItem(it);
        });
        lastId = safeString(result.last_id);
        if (!lastId || items.length === 0) break;
      } catch (e2) {
        logWarn('fetchProductPrices', 'Цены не загружены: ' + e2.message);
        break;
      }
      Utilities.sleep(250);
    }
  }

  return map;
}

function normalizePriceItem(item) {
  const p = item.price && typeof item.price === 'object' ? item.price : item;
  return {
    price: pickNumOrBlank(p.price, item.price),
    marketing_price: pickNumOrBlank(p.marketing_price, item.marketing_price),
    old_price: pickNumOrBlank(p.old_price, item.old_price),
    currency_code: safeString(p.currency_code || item.currency_code || 'RUB')
  };
}

function fetchProductStocks(productIds) {
  const map = {};
  if (!productIds.length) return map;
  try {
    let cursor = '';
    while (true) {
      const data = ozonSellerPost('/v4/product/info/stocks', {
        cursor: cursor,
        filter: { product_id: productIds.map(String), visibility: 'ALL' },
        limit: 1000
      }, 'product stocks');
      const result = data.result || data || {};
      const items = safeArray(result.items);
      items.forEach(it => {
        const id = safeString(it.product_id || it.id);
        const stocksList = normalizeStocksContainer(it.stocks);
        let fbo = 0, fbs = 0, reserved = 0;
        stocksList.forEach(s => {
          const present = num(s.present || s.free_to_sell_amount || s.valid_stock_count || s.stock || s.available_stock_count);
          reserved += num(s.reserved || s.reserved_stock_count);
          const type = safeString(s.type || s.source || s.delivery_schema || s.name).toLowerCase();
          if (type.indexOf('fbs') >= 0 || type.indexOf('rfbs') >= 0) fbs += present;
          else fbo += present;
        });
        if (id) map[id] = { fbo_stock: fbo, fbs_stock: fbs, reserved: reserved };
      });
      cursor = safeString(result.cursor || result.next_cursor);
      if (!cursor || items.length === 0) break;
      Utilities.sleep(250);
    }
  } catch (e) {
    logWarn('fetchProductStocks', 'Остатки не загружены: ' + e.message);
  }
  return map;
}

function normalizeStocksContainer(stocks) {
  if (!stocks) return [];
  if (Array.isArray(stocks)) return stocks;
  if (typeof stocks === 'object') {
    const out = [];
    Object.keys(stocks).forEach(k => {
      const v = stocks[k];
      if (v && typeof v === 'object') {
        v.name = v.name || k;
        out.push(v);
      }
    });
    return out;
  }
  return [];
}

function fetchFboPostings(from, to, productMaps) {
  const rows = [];
  let offset = 0;
  const limit = 1000;
  while (true) {
    const payload = {
      dir: 'ASC',
      filter: {
        since: toOzonDateTime(from, false),
        to: toOzonDateTime(to, true),
        status: ''
      },
      limit: limit,
      offset: offset,
      translit: true,
      with: { analytics_data: true, financial_data: true }
    };
    const data = ozonSellerPost('/v2/posting/fbo/list', payload, 'fbo postings');
    const postings = normalizePostingsResponse(data);
    postings.forEach(p => rows.push.apply(rows, normalizePosting(p, 'FBO', productMaps)));
    if (postings.length < limit) break;
    offset += limit;
    Utilities.sleep(250);
  }
  return rows;
}

function fetchFbsPostings(from, to, productMaps) {
  const rows = [];
  let offset = 0;
  const limit = 1000;
  while (true) {
    const payload = {
      dir: 'ASC',
      filter: {
        since: toOzonDateTime(from, false),
        to: toOzonDateTime(to, true),
        status: ''
      },
      limit: limit,
      offset: offset,
      translit: true,
      with: { analytics_data: true, financial_data: true }
    };
    const data = ozonSellerPost('/v3/posting/fbs/list', payload, 'fbs postings');
    const postings = normalizePostingsResponse(data);
    postings.forEach(p => rows.push.apply(rows, normalizePosting(p, 'FBS', productMaps)));
    if (postings.length < limit) break;
    offset += limit;
    Utilities.sleep(250);
  }
  return rows;
}

function normalizePostingsResponse(data) {
  if (!data) return [];
  if (Array.isArray(data.result)) return data.result;
  if (data.result && Array.isArray(data.result.postings)) return data.result.postings;
  if (Array.isArray(data.postings)) return data.postings;
  return [];
}

function normalizePosting(posting, scheme, productMaps) {
  const rows = [];
  const postingNumber = safeString(posting.posting_number);
  const orderId = safeString(posting.order_id);
  const status = safeString(posting.status);
  const date = posting.created_at || posting.in_process_at || posting.shipment_date || posting.delivering_date || posting.cancelled_at;
  const financialProducts = posting.financial_data && Array.isArray(posting.financial_data.products) ? posting.financial_data.products : [];
  const products = safeArray(posting.products).length ? safeArray(posting.products) : financialProducts;
  const postingServices = posting.financial_data && posting.financial_data.posting_services ? posting.financial_data.posting_services : {};
  const postingDeliveryCharge = sumObjectNumbers(postingServices);

  products.forEach(product => {
    const sku = safeString(product.sku || product.sku_id || product.ozon_sku);
    const productId = safeString(product.product_id || product.id);
    const offerId = safeString(product.offer_id || (productMaps.bySku[sku] && productMaps.bySku[sku].offer_id) || (productMaps.byProductId[productId] && productMaps.byProductId[productId].offer_id));
    const qty = num(product.quantity || 1) || 1;
    const price = num(product.price || product.client_price || product.old_price || product.payout);
    const itemServices = product.item_services ? sumObjectNumbers(product.item_services) : 0;
    const deliveryCharge = itemServices || safeDiv(postingDeliveryCharge, Math.max(1, products.length));

    rows.push({
      date: date,
      scheme: scheme,
      posting_number: postingNumber,
      order_id: orderId,
      status: status,
      offer_id: offerId,
      sku: sku,
      product_id: productId,
      'Название': safeString(product.name || (productMaps.byOffer[offerId] && productMaps.byOffer[offerId].name)),
      quantity: qty,
      price: price,
      total: price * qty,
      commission_amount: num(product.commission_amount),
      payout: num(product.payout),
      delivery_charge: deliveryCharge,
      cancel_reason_id: safeString(posting.cancel_reason_id),
      in_process_at: posting.in_process_at || '',
      shipment_date: posting.shipment_date || '',
      raw_status: status
    });
  });

  return rows;
}

function normalizeFinanceOperation(op, productMaps, orderProductMap) {
  const rows = [];
  const items = safeArray(op.items);
  const normalizedItems = items.length ? items : [{ sku: '', name: '' }];
  const share = 1 / Math.max(1, normalizedItems.length);
  const services = safeArray(op.services);
  const servicesTotal = services.reduce((s, x) => s + num(x.price), 0);
  const opType = safeString(op.operation_type);
  const opTypeName = safeString(op.operation_type_name);
  const isRet = isReturnOperation(opType, opTypeName);
  const qtyEstimateBase = inferFinanceQuantity(op, isRet);

  normalizedItems.forEach(item => {
    const sku = safeString(item.sku);
    const productName = safeString(item.name);
    const mapped = productMaps.bySku[sku] || {};
    const byPostingSku = orderProductMap && orderProductMap.byPostingSku ? orderProductMap.byPostingSku[safeString(op.posting && op.posting.posting_number) + '|' + sku] : null;
    const byPostingOnly = orderProductMap && orderProductMap.byPosting ? orderProductMap.byPosting[safeString(op.posting && op.posting.posting_number)] : null;
    const offerId = safeString(mapped.offer_id || item.offer_id || (byPostingSku && byPostingSku.offer_id) || (byPostingOnly && byPostingOnly.offer_id));

    rows.push({
      operation_date: op.operation_date || '',
      posting_number: safeString(op.posting && op.posting.posting_number),
      delivery_schema: safeString(op.posting && op.posting.delivery_schema),
      operation_id: safeString(op.operation_id),
      operation_type: opType,
      operation_type_name: opTypeName,
      type: safeString(op.type),
      sku: sku,
      offer_id: offerId,
      'Название': productName || safeString(mapped.name),
      quantity_estimated: qtyEstimateBase * share,
      accruals_for_sale: num(op.accruals_for_sale) * share,
      sale_commission: num(op.sale_commission) * share,
      amount: num(op.amount) * share,
      delivery_charge: num(op.delivery_charge) * share,
      return_delivery_charge: num(op.return_delivery_charge) * share,
      services_total: servicesTotal * share,
      services_json: JSON.stringify(services),
      items_count: normalizedItems.length,
      raw_json: JSON.stringify(op)
    });
  });
  return rows;
}

function inferFinanceQuantity(op, isRet) {
  const accrual = num(op.accruals_for_sale);
  const amount = num(op.amount);
  if (isRet) return -1;
  if (accrual > 0 || amount > 0) return 1;
  return 0;
}

function financeRowMetrics(r) {
  const operationType = safeString(r.operation_type);
  const operationName = safeString(r.operation_type_name);
  const isRet = isReturnOperation(operationType, operationName);
  const qty = num(r.quantity_estimated);
  const revenue = num(r.accruals_for_sale);
  const commission = Math.abs(num(r.sale_commission));
  const delivery = Math.abs(num(r.delivery_charge));
  const returnDelivery = Math.abs(num(r.return_delivery_charge));
  const services = Math.abs(num(r.services_total));
  const logistics = delivery + returnDelivery;
  const returnExpense = isRet ? (returnDelivery + services) : 0;
  const ozonExpenses = commission + delivery + returnDelivery + services;
  return {
    qty: qty,
    revenue: revenue,
    commission: commission,
    logistics: logistics,
    returnExpense: returnExpense,
    ozonExpenses: ozonExpenses,
    isReturn: isRet
  };
}

function getPerformanceToken() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty(PROP_KEYS.PERF_ACCESS_TOKEN);
  const expiresAt = num(props.getProperty(PROP_KEYS.PERF_TOKEN_EXPIRES_AT));
  if (cached && expiresAt && Date.now() < expiresAt - 60000) return cached;

  const clientId = props.getProperty(PROP_KEYS.PERF_CLIENT_ID);
  const secret = props.getProperty(PROP_KEYS.PERF_CLIENT_SECRET);
  const body = {
    client_id: clientId,
    client_secret: secret,
    grant_type: 'client_credentials'
  };

  const data = fetchJson(OZON_PERF_BASE + '/api/client/token', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    headers: { Accept: 'application/json' },
    muteHttpExceptions: true
  }, 'performance token');

  const token = data.access_token;
  if (!token) throw new Error('Performance API не вернул access_token');
  const ttl = num(data.expires_in) || 1800;
  props.setProperty(PROP_KEYS.PERF_ACCESS_TOKEN, token);
  props.setProperty(PROP_KEYS.PERF_TOKEN_EXPIRES_AT, String(Date.now() + ttl * 1000));
  return token;
}

function fetchPerformanceCampaigns(token) {
  const rows = [];
  const types = ['SKU', 'SEARCH_PROMO'];
  types.forEach(type => {
    try {
      const url = OZON_PERF_BASE + '/api/client/campaign?advObjectType=' + encodeURIComponent(type);
      const data = fetchJson(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
        muteHttpExceptions: true
      }, 'performance campaigns ' + type);
      const list = safeArray(data.list || data.campaigns || data.result);
      list.forEach(c => {
        c.advObjectType = c.advObjectType || type;
        rows.push(c);
      });
    } catch (e) {
      logWarn('fetchPerformanceCampaigns', 'Не удалось получить кампании ' + type + ': ' + e.message);
    }
    Utilities.sleep(500);
  });
  return rows;
}

function requestPerformanceStatsJson(token, campaignIds, dateFrom, dateTo) {
  // Для ДРР по каждому товару сначала пробуем получить детализацию SKU.
  // Если конкретный тип кампании не поддерживает такую группировку, уходим на обычный отчет.
  const groups = ['SKU', 'DATE'];
  for (let i = 0; i < groups.length; i++) {
    const payload = {
      campaigns: campaignIds,
      dateFrom: dateFrom,
      dateTo: dateTo,
      groupBy: groups[i]
    };
    try {
      const data = fetchJson(OZON_PERF_BASE + '/api/client/statistics/json', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
        muteHttpExceptions: true
      }, 'performance statistics json ' + groups[i]);
      const uuid = safeString(data.UUID || data.uuid || data.result && data.result.UUID);
      if (uuid) return uuid;
    } catch (e) {
      logWarn('requestPerformanceStatsJson', 'Отчет рекламы groupBy=' + groups[i] + ' не получен: ' + e.message);
    }
  }
  return '';
}

function downloadPerformanceReport(token, uuid) {
  for (let i = 0; i < 8; i++) {
    try {
      const url = OZON_PERF_BASE + '/api/client/statistics/report?UUID=' + encodeURIComponent(uuid);
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      const text = resp.getContentText();
      if (code >= 200 && code < 300 && text) return JSON.parse(text);
      if (code === 202 || code === 204 || code === 404 || code === 429 || code >= 500) {
        Utilities.sleep(3000 + i * 2000);
        continue;
      }
      logWarn('downloadPerformanceReport', 'Ошибка скачивания отчета ' + uuid + ': HTTP ' + code + ' ' + text.slice(0, 300));
      return null;
    } catch (e) {
      logWarn('downloadPerformanceReport', 'Попытка ' + (i + 1) + ' неуспешна: ' + e.message);
      Utilities.sleep(3000 + i * 2000);
    }
  }
  return null;
}

function parsePerformanceReport(reportJson, campaignById, productMaps) {
  const output = [];
  const rows = extractPerformanceRows(reportJson);
  rows.forEach(row => {
    const campaignId = safeString(row.campaignId || row.campaign_id || row.campaign || row.id);
    const campaign = campaignById[campaignId] || {};
    const sku = safeString(row.sku || row.SKU || row.product_sku || row.skuId || row.ozonSku);
    const productId = safeString(row.productId || row.product_id || row.idProduct);
    const offerId = safeString(row.offer_id || row.offerId || (productMaps.bySku[sku] && productMaps.bySku[sku].offer_id) || (productMaps.byProductId[productId] && productMaps.byProductId[productId].offer_id));

    output.push({
      date: row.date || row.day || row.Date || '',
      campaign_id: campaignId,
      campaign_name: safeString(row.campaignName || row.campaign_name || campaign.title || campaign.name),
      adv_object_type: safeString(campaign.advObjectType || row.advObjectType || row.adv_object_type),
      offer_id: offerId,
      sku: sku,
      product_id: productId,
      'Показы': parseLocaleNumber(row.views || row.impressions || row.shows || row['Показы']),
      'Клики': parseLocaleNumber(row.clicks || row['Клики']),
      'Расход': parseLocaleNumber(row.moneySpent || row.expense || row.cost || row.spend || row['Расход']),
      'Заказы_реклама': parseLocaleNumber(row.orders || row.ordersCount || row['Заказы']),
      'Выручка_реклама': parseLocaleNumber(row.revenue || row.sales || row.demand || row['Выручка']),
      'Источник': 'Performance API'
    });
  });
  return output;
}

function extractPerformanceRows(reportJson) {
  if (!reportJson) return [];
  if (Array.isArray(reportJson)) return reportJson;
  if (reportJson.report && Array.isArray(reportJson.report.rows)) return reportJson.report.rows;
  if (Array.isArray(reportJson.rows)) return reportJson.rows;
  if (reportJson.result && Array.isArray(reportJson.result.rows)) return reportJson.result.rows;
  if (reportJson.result && reportJson.result.report && Array.isArray(reportJson.result.report.rows)) return reportJson.result.report.rows;

  const found = [];
  walkJsonForAdRows(reportJson, found);
  return found;
}

function walkJsonForAdRows(node, out) {
  if (!node) return;
  if (Array.isArray(node)) {
    node.forEach(x => walkJsonForAdRows(x, out));
    return;
  }
  if (typeof node !== 'object') return;

  const hasMetrics = node.moneySpent !== undefined || node.expense !== undefined || node.cost !== undefined || node.spend !== undefined || node['Расход'] !== undefined;
  const hasId = node.sku !== undefined || node.SKU !== undefined || node.productId !== undefined || node.product_id !== undefined || node.campaignId !== undefined || node.campaign_id !== undefined;
  if (hasMetrics && hasId) out.push(node);

  Object.keys(node).forEach(k => {
    if (typeof node[k] === 'object') walkJsonForAdRows(node[k], out);
  });
}

function ozonSellerPost(path, payload, scope) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty(PROP_KEYS.SELLER_CLIENT_ID);
  const apiKey = props.getProperty(PROP_KEYS.SELLER_API_KEY);
  if (!clientId || !apiKey) throw new Error('Не заданы Ozon Seller Client-Id и Api-Key');

  return fetchJson(OZON_SELLER_BASE + path, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    headers: {
      'Client-Id': clientId,
      'Api-Key': apiKey,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  }, scope || path);
}

function fetchJson(url, options, scope) {
  const maxAttempts = 5;
  const maxRedirects = 5;
  let currentUrl = url;
  let redirectCount = 0;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const requestOptions = Object.assign({}, options || {}, {
      muteHttpExceptions: true,
      followRedirects: false
    });

    const resp = UrlFetchApp.fetch(currentUrl, requestOptions);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code >= 200 && code < 300) {
      if (!text) return {};
      try {
        return JSON.parse(text);
      } catch (e) {
        throw new Error('Не удалось распарсить JSON для ' + scope + ': ' + text.slice(0, 500));
      }
    }

    if ([301, 302, 303, 307, 308].indexOf(code) >= 0) {
      const headers = resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders();
      const location = headers.Location || headers.location;
      if (!location) {
        throw new Error(scope + ': HTTP ' + code + ', но сервер не вернул Location для редиректа. Ответ: ' + text.slice(0, 500));
      }
      redirectCount += 1;
      if (redirectCount > maxRedirects) {
        throw new Error(scope + ': слишком много редиректов. Последний URL: ' + currentUrl);
      }
      currentUrl = absolutizeUrl(currentUrl, location);
      logWarn('fetchJson', scope + ': HTTP ' + code + ', перехожу на ' + currentUrl);
      continue;
    }

    if (code === 429 || code >= 500) {
      const pause = Math.min(30000, 1000 * Math.pow(2, attempt - 1));
      logWarn('fetchJson', scope + ': HTTP ' + code + ', попытка ' + attempt + ', пауза ' + pause + ' мс');
      Utilities.sleep(pause);
      continue;
    }

    throw new Error(scope + ': HTTP ' + code + ' — ' + text.slice(0, 1000));
  }
  throw new Error(scope + ': превышено число попыток');
}

function absolutizeUrl(baseUrl, location) {
  const loc = safeString(location);
  if (/^https?:\/\//i.test(loc)) return loc;
  const m = String(baseUrl).match(/^(https?:\/\/[^\/]+)/i);
  const origin = m ? m[1] : '';
  if (loc.charAt(0) === '/') return origin + loc;
  return origin + '/' + loc;
}

function requireSellerCredentials() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_KEYS.SELLER_CLIENT_ID) || !props.getProperty(PROP_KEYS.SELLER_API_KEY)) {
    throw new Error('Сначала заполни seller_client_id_tmp и seller_api_key_tmp в Настройках и запусти saveCredentialsFromSettings().');
  }
}

function seedSettings() {
  const sh = getSheet(SHEETS.SETTINGS);
  const existing = getSettingsMap();
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  const defaults = [
    ['date_from', existing.date_from || firstDay, 'Дата начала выбранного периода. Удобнее менять на листе Период'],
    ['date_to', existing.date_to || today, 'Дата окончания выбранного периода. Удобнее менять на листе Период'],
    ['historical_days', existing.historical_days || 90, 'Сколько дней истории брать для средних % отмен и возвратов'],
    ['tax_mode', existing.tax_mode || 'revenue', 'revenue = налог с оборота, profit = налог с прибыли, none = не считать'],
    ['tax_percent', existing.tax_percent || 0.06, 'Например, 0.06 = 6%'],
    ['default_cancel_rate', existing.default_cancel_rate || 0.03, 'Если по товару мало истории'],
    ['default_return_rate', existing.default_return_rate || 0.05, 'Если по товару мало истории'],
    ['default_ozon_expense_pct', existing.default_ozon_expense_pct || 0.25, 'Средняя доля расходов Ozon от выручки для прогноза'],
    ['default_return_expense_per_item', existing.default_return_expense_per_item || 0, 'Средний расход на один возврат, если нет истории'],
    ['default_avg_price', existing.default_avg_price || 0, 'Запасная средняя цена, если нет заказов и истории'],
    ['auto_fix_kopeck_prices', existing.auto_fix_kopeck_prices === '' ? true : existing.auto_fix_kopeck_prices, 'TRUE = автоматически исправлять цены вида 80000 → 800'],
    ['target_margin_pct', existing.target_margin_pct || 0.25, 'Целевая чистая маржа'],
    ['critical_ad_share_pct', existing.critical_ad_share_pct || 0.20, 'Критичный ДРР'],
    ['critical_ozon_expense_share_pct', existing.critical_ozon_expense_share_pct || 0.35, 'Критичная доля расходов Ozon'],
    ['critical_returns_pct', existing.critical_returns_pct || 0.10, 'Критичный % возвратов'],
    ['use_performance_api', existing.use_performance_api || false, 'TRUE = тянуть рекламу через Performance API; FALSE = заполнять API_Реклама вручную'],
    ['seller_client_id_tmp', '', 'Временно вставь Client-Id, затем запусти сохранение ключей'],
    ['seller_api_key_tmp', '', 'Временно вставь Api-Key, затем запусти сохранение ключей'],
    ['performance_client_id_tmp', '', 'Временно вставь client_id Performance API'],
    ['performance_client_secret_tmp', '', 'Временно вставь client_secret Performance API'],
    ['last_sync', existing.last_sync || '', 'Последняя успешная синхронизация'],
    ['version', OZON_UNIT_VERSION, 'Версия скрипта']
  ];

  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.SETTINGS.length).setValues([HEADERS.SETTINGS]);
  sh.getRange(2, 1, defaults.length, 3).setValues(defaults);
  sh.getRange('A1:C1').setFontWeight('bold').setBackground('#e8f0fe');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
}

function setupPeriodSheet() {
  let sh = getSheet(SHEETS.PERIOD);
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet(SHEETS.PERIOD, 1);

  const settings = getSettingsMap();
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  const currentFrom = parseSheetDate(sh.getRange('B4').getValue()) || parseSheetDate(settings.date_from) || firstDay;
  const currentTo = parseSheetDate(sh.getRange('B5').getValue()) || parseSheetDate(settings.date_to) || today;
  const currentTax = sh.getRange('B7').getValue() !== '' ? sh.getRange('B7').getValue() : (settings.tax_percent || 0.06);
  const currentTaxMode = safeString(sh.getRange('B8').getValue()) || safeString(settings.tax_mode || 'revenue');

  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1:D1').merge().setValue('Выбор периода и основных настроек')
    .setBackground('#0b57d0').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);

  const rows = [
    ['', '', '', ''],
    ['Что менять здесь', 'Значение', 'Как выбрать', ''],
    ['Дата начала', currentFrom, 'Нажми на ячейку B4 — Google Sheets откроет календарь', ''],
    ['Дата окончания', currentTo, 'Нажми на ячейку B5 — Google Sheets откроет календарь', ''],
    ['', '', '', ''],
    ['Налог, %', currentTax, 'Например 6% или 0,06', ''],
    ['Режим налога', currentTaxMode, 'revenue = с оборота, profit = с прибыли, none = не считать', ''],
    ['', '', '', ''],
    ['Кнопка запуска', 'Меню Ozon Unit → Полная синхронизация и пересчет', 'Скрипт возьмет даты именно отсюда', '']
  ];
  sh.getRange(2, 1, rows.length, 4).setValues(rows);

  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sh.getRange('B4:B5').setDataValidation(dateValidation).setNumberFormat('dd.mm.yyyy');

  const taxValidation = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(false).build();
  sh.getRange('B7').setDataValidation(taxValidation).setNumberFormat('0.00%');

  const taxModeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['revenue', 'profit', 'none'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B8').setDataValidation(taxModeValidation);

  sh.getRange('A3:D3').setBackground('#d9ead3').setFontWeight('bold');
  sh.getRange('A4:A5').setBackground('#d9ead3').setFontWeight('bold');
  sh.getRange('B4:B5').setBackground('#fff2cc').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A7:A8').setBackground('#d9ead3').setFontWeight('bold');
  sh.getRange('B7:B8').setBackground('#fff2cc').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A10:D10').setBackground('#f3f3f3');
  sh.getRange('A2:D10').setBorder(true, true, true, true, true, true, '#c9daf8', SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 460);
  sh.setColumnWidth(4, 30);
  sh.setFrozenRows(1);
}

function clearTemporaryCredentialCells() {
  const sh = getSheet(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const key = safeString(values[r][0]);
    if (['seller_client_id_tmp', 'seller_api_key_tmp', 'performance_client_id_tmp', 'performance_client_secret_tmp'].indexOf(key) >= 0) {
      sh.getRange(r + 1, 2).clearContent();
    }
  }
}

function setSettingValue(key, value) {
  const sh = getSheet(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (safeString(values[r][0]) === key) {
      sh.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

function getSettingsMap() {
  const sh = getSheet(SHEETS.SETTINGS);
  if (!sh) return {};
  const values = sh.getDataRange().getValues();
  const out = {};
  for (let r = 1; r < values.length; r++) {
    const key = safeString(values[r][0]);
    if (!key) continue;
    out[key] = values[r][1];
  }
  return out;
}

function validateSelectedPeriod() {
  const p = getSelectedPeriod();
  if (p.from > p.to) throw new Error('date_from не может быть позже date_to');
}

function getSelectedPeriod() {
  const visible = getPeriodFromPeriodSheet();
  if (visible && visible.from && visible.to) return visible;

  const settings = getSettingsMap();
  const from = parseSheetDate(settings.date_from);
  const to = parseSheetDate(settings.date_to);
  if (!from || !to) throw new Error('Заполни даты на листе Период или date_from/date_to на листе Настройки');
  return { from: startOfDay(from), to: endOfDay(to) };
}

function getPeriodFromPeriodSheet() {
  const sh = getSheet(SHEETS.PERIOD);
  if (!sh) return null;
  const from = parseSheetDate(sh.getRange('B4').getValue());
  const to = parseSheetDate(sh.getRange('B5').getValue());
  if (!from || !to) return null;

  const settings = getSettingsMap();
  const taxRate = sh.getRange('B7').getValue();
  const taxMode = sh.getRange('B8').getValue();
  if (taxRate !== '') setSettingValue('tax_percent', toRate(taxRate));
  if (safeString(taxMode)) setSettingValue('tax_mode', safeString(taxMode));
  setSettingValue('date_from', from);
  setSettingValue('date_to', to);

  return { from: startOfDay(from), to: endOfDay(to) };
}

function getSyncPeriodWithHistory() {
  const selected = getSelectedPeriod();
  const settings = getSettingsMap();
  const days = Math.max(0, Math.round(num(settings.historical_days) || 90));
  return { from: addDays(selected.from, -days), to: selected.to };
}

function getCostsMap() {
  const rows = readSheetAsObjects(SHEETS.COSTS);
  const out = {};
  rows.forEach(r => {
    const offerId = safeString(r.offer_id);
    if (!offerId) return;
    out[offerId] = {
      offer_id: offerId,
      sku: safeString(r.sku),
      name: safeString(r['Название']),
      cost_per_unit: num(r['Себестоимость_шт']),
      package_per_unit: num(r['Упаковка_шт']),
      additional_per_unit: num(r['Доп_расход_шт']),
      tax_rate_override: toRate(r['Налог_%_override']),
      target_margin_override: toRate(r['Целевая_маржа_%_override']),
      cancel_rate_override: toRate(r['Средний_%_отмен_override']),
      return_rate_override: toRate(r['Средний_%_возвратов_override'])
    };
  });
  return out;
}

function getHistoryMap() {
  const rows = readSheetAsObjects(SHEETS.HISTORY);
  const out = {};
  rows.forEach(r => {
    const offerId = safeString(r.offer_id);
    if (!offerId) return;
    out[offerId] = {
      offer_id: offerId,
      sku: safeString(r.sku),
      name: safeString(r['Название']),
      avg_cancel_rate: num(r['Средний_%_отмен']),
      avg_return_rate: num(r['Средний_%_возвратов']),
      avg_price: num(r['Средняя_цена']),
      avg_ozon_expense_pct: num(r['Средний_%_расходов_Ozon']),
      avg_commission_pct: num(r['Средний_%_комиссии']),
      avg_logistics_per_item: num(r['Средняя_логистика_шт']),
      avg_return_expense_per_item: num(r['Средний_расход_возврата_шт']),
      avg_ad_share: num(r['Средний_ДРР'])
    };
  });
  return out;
}

function getProductMaps() {
  const rows = readSheetAsObjects(SHEETS.PRODUCTS);
  const byOffer = {};
  const bySku = {};
  const byProductId = {};
  rows.forEach(r => {
    const productId = safeString(r.product_id);
    const offerId = safeString(r.offer_id);
    const sku = safeString(r.sku);
    const item = {
      product_id: productId,
      offer_id: offerId,
      sku: sku,
      name: safeString(r['Название'])
    };
    if (offerId) byOffer[offerId] = item;
    if (sku) bySku[sku] = item;
    if (productId) byProductId[productId] = item;
  });
  return { byOffer: byOffer, bySku: bySku, byProductId: byProductId };
}

function getOrderProductMap() {
  const rows = readSheetAsObjects(SHEETS.ORDERS);
  const byPostingSku = {};
  const byPosting = {};
  rows.forEach(r => {
    const posting = safeString(r.posting_number);
    const sku = safeString(r.sku);
    const offerId = safeString(r.offer_id);
    if (!posting || !offerId) return;
    const item = { offer_id: offerId, sku: sku, name: safeString(r['Название']) };
    if (sku) byPostingSku[posting + '|' + sku] = item;
    if (!byPosting[posting]) byPosting[posting] = item;
  });
  return { byPostingSku: byPostingSku, byPosting: byPosting };
}

function addMissingProductsToCosts(productRows) {
  const sh = getSheet(SHEETS.COSTS);
  const current = getCostsMap();
  const newRows = [];
  productRows.forEach(p => {
    const offerId = safeString(p.offer_id);
    if (!offerId || current[offerId]) return;
    newRows.push([
      offerId,
      safeString(p.sku),
      safeString(p['Ссылка_Ozon']),
      safeString(p['Название']),
      '', '', '', '', '', '', '', '', ''
    ]);
  });
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, HEADERS.COSTS.length).setValues(newRows);
    logInfo('addMissingProductsToCosts', 'Добавлено новых товаров в Себестоимость: ' + newRows.length);
  }
}

function resolveOfferIdFromRow(row, productMaps) {
  const offerId = safeString(row.offer_id || row['Артикул']);
  if (offerId) return offerId;
  const sku = safeString(row.sku || row['SKU Ozon']);
  if (sku && productMaps.bySku[sku]) return productMaps.bySku[sku].offer_id;
  const productId = safeString(row.product_id || row['product_id']);
  if (productId && productMaps.byProductId[productId]) return productMaps.byProductId[productId].offer_id;
  return 'NO_OFFER';
}

function allocateAdsToAggregation(adsRows, productMaps, agg, touch, weightFn) {
  let unassigned = 0;
  adsRows.forEach(r => {
    const offerId = resolveOfferIdFromRow(r, productMaps);
    const spend = num(r['Расход']);
    if (!spend) return;
    if (offerId && offerId !== 'NO_OFFER') {
      touch(offerId).ads += spend;
    } else {
      unassigned += spend;
    }
  });

  if (!unassigned) return;
  const keys = Object.keys(agg).filter(k => k !== 'NO_OFFER');
  const weights = keys.map(k => Math.max(0, num(weightFn(agg[k]))));
  const totalWeight = weights.reduce((s, x) => s + x, 0);

  if (totalWeight <= 0) {
    touch('NO_OFFER').ads += unassigned;
    logWarn('allocateAdsToAggregation', 'Рекламу ' + unassigned + ' ₽ не удалось распределить по товарам: нет SKU/offer_id и нет базы для распределения.');
    return;
  }

  keys.forEach((k, i) => {
    agg[k].ads += unassigned * weights[i] / totalWeight;
  });
  logWarn('allocateAdsToAggregation', 'Реклама без SKU/offer_id распределена пропорционально обороту/заказам: ' + unassigned + ' ₽. Для точного ДРР нужен SKU-детализированный отчет Performance API.');
}

function getTaxRateForOffer(settings, costRow) {
  const override = costRow && costRow.tax_rate_override;
  if (override !== '' && override !== null && override !== undefined && !isNaN(override)) return num(override);
  return toRate(settings.tax_percent);
}

function getTargetMarginForOffer(settings, costRow) {
  const override = costRow && costRow.target_margin_override;
  if (override !== '' && override !== null && override !== undefined && !isNaN(override)) return num(override);
  return toRate(settings.target_margin_pct);
}

function calcTax(settings, taxRate, revenue, ozonExpenses, ads, cost, pack, additional) {
  const mode = safeString(settings.tax_mode || 'revenue').toLowerCase();
  if (mode === 'none' || mode === 'off' || taxRate <= 0) return 0;
  if (mode === 'profit') {
    return Math.max(0, revenue - ozonExpenses - ads - cost - pack - additional) * taxRate;
  }
  return Math.max(0, revenue) * taxRate;
}

function getRateOverrideOrHistory(override, historyRate, defaultRate) {
  if (override !== '' && override !== null && override !== undefined && !isNaN(override)) return clamp(num(override), 0, 1);
  if (historyRate !== '' && historyRate !== null && historyRate !== undefined && !isNaN(historyRate) && num(historyRate) > 0) return clamp(num(historyRate), 0, 1);
  return clamp(toRate(defaultRate), 0, 1);
}

function buildStatus(profit, margin, adShare, ozonShare, settings, targetMargin) {
  if (profit < 0) return 'Красный';
  if (margin < targetMargin || adShare > num(settings.critical_ad_share_pct) || ozonShare > num(settings.critical_ozon_expense_share_pct)) return 'Желтый';
  return 'Зеленый';
}

function buildRecommendation(profit, margin, adShare, ozonShare, settings, targetMargin, mode) {
  const criticalAd = num(settings.critical_ad_share_pct);
  const criticalOzon = num(settings.critical_ozon_expense_share_pct);

  if (profit < 0 && adShare > criticalAd) return 'Убыток + высокий ДРР: снизить/отключить рекламу или поднять цену.';
  if (profit < 0 && ozonShare > criticalOzon) return 'Убыток из-за расходов Ozon: проверить логистику, схему поставки, цену и комиссию.';
  if (profit < 0) return mode === 'fact'
    ? 'Факт в минусе: проверь, не висят ли заказы без выкупа; смотри вкладку Прогноз.'
    : 'Прогноз в минусе: менять цену, себестоимость, рекламу или выводить товар.';
  if (margin < targetMargin) return 'Маржа ниже цели: поднять цену, снизить себестоимость или ДРР.';
  if (adShare > criticalAd) return 'ДРР выше нормы: оптимизировать кампании и ставки.';
  if (ozonShare > criticalOzon) return 'Высокая доля расходов Ozon: проверить логистику, комиссии и возвраты.';
  return 'ОК: товар можно удерживать/масштабировать при наличии остатков.';
}

function compareConclusion(factProfit, forecastProfit) {
  if (factProfit < 0 && forecastProfit > 0) return 'Факт в минусе временно: вероятно, заказы еще не выкуплены.';
  if (factProfit > 0 && forecastProfit < 0) return 'Факт хороший, но прогноз плохой: риск возвратов, отмен или рекламы.';
  if (factProfit < 0 && forecastProfit < 0) return 'Реально проблемный товар: факт и прогноз в минусе.';
  if (factProfit > 0 && forecastProfit > 0) return 'Товар выглядит здоровым.';
  return 'Нейтрально: нужно больше данных.';
}

function calcUnitKpi(rows, mode) {
  let revenue = 0, profit = 0, ads = 0, ozon = 0, negativeCount = 0, highDrrCount = 0;
  const settings = getSettingsMap();
  const drrCritical = num(settings.critical_ad_share_pct);

  rows.forEach(r => {
    const rev = mode === 'fact' ? num(r['Выручка_факт']) : num(r['Выручка_прогноз']);
    const pr = mode === 'fact' ? num(r['Чистая_прибыль']) : num(r['Чистая_прибыль_прогноз']);
    const ad = num(r['Реклама_факт']);
    const oz = mode === 'fact' ? num(r['Расходы_Ozon_факт']) : num(r['Расходы_Ozon_прогноз']);
    revenue += rev;
    profit += pr;
    ads += ad;
    ozon += oz;
    if (pr < 0) negativeCount += 1;
    if (safeDiv(ad, rev) > drrCritical) highDrrCount += 1;
  });

  return {
    revenue: revenue,
    profit: profit,
    margin: safeDiv(profit, revenue),
    adShare: safeDiv(ads, revenue),
    ozonShare: safeDiv(ozon, revenue),
    negativeCount: negativeCount,
    highDrrCount: highDrrCount
  };
}

function ensureSheetWithHeader(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const firstRow = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn() || 1)).getValues()[0];
  const hasHeader = firstRow.some(x => safeString(x));
  if (!hasHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0].map(safeString);
    const same = headers.every((h, i) => current[i] === h);
    if (!same) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  return sh;
}

function getSheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function writeObjects(sheetName, headers, objects) {
  const sh = ensureSheetWithHeader(sheetName, headers);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(headers.length, sh.getLastColumn())).clearContent().clearFormat();
  if (!objects || !objects.length) {
    formatSheetBase(sh, headers.length);
    return;
  }
  const values = objects.map(obj => headers.map(h => obj[h] !== undefined ? obj[h] : ''));
  sh.getRange(2, 1, values.length, headers.length).setValues(values);
  formatSheetBase(sh, headers.length);
}

function readSheetAsObjects(sheetName) {
  const sh = getSheet(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(safeString);
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    if (!values[r].some(v => safeString(v) !== '')) continue;
    const obj = {};
    headers.forEach((h, i) => obj[h] = values[r][i]);
    rows.push(obj);
  }
  return rows;
}

function formatAllSheets() {
  Object.keys(SHEETS).forEach(k => {
    const sh = getSheet(SHEETS[k]);
    if (sh) formatSheetBase(sh, sh.getLastColumn());
  });
  formatPercentColumns(SHEETS.HISTORY, ['Средний_%_отмен', 'Средний_%_возвратов', 'Средний_%_расходов_Ozon', 'Средний_%_комиссии', 'Средний_ДРР']);
  formatPercentColumns(SHEETS.UNIT_FACT, ['Маржа_%', 'ДРР_%', 'Расходы_Ozon_%']);
  formatPercentColumns(SHEETS.UNIT_FORECAST, ['Средний_%_возвратов', 'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%']);
  formatPercentColumns(SHEETS.COMPARE, ['Маржа_факт_%', 'Маржа_прогноз_%', 'ДРР_факт_%', 'ДРР_прогноз_%']);
  beautifyWorkbook();
}

function formatSheetBase(sh, columns) {
  if (!columns) return;
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#0b57d0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sh.getDataRange().setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.autoResizeColumns(1, Math.min(columns, 20));

  if (sh.getFilter()) sh.getFilter().remove();

  // На листе "Период" есть объединенная шапка A1:D1.
  // Google Sheets не позволяет создавать фильтр на диапазоне с объединенными ячейками.
  if (sheetHasMergedCells(sh)) return;

  if (sh.getLastRow() > 1) sh.getRange(1, 1, sh.getLastRow(), columns).createFilter();
}

function sheetHasMergedCells(sh) {
  try {
    return sh.getDataRange().getMergedRanges().length > 0;
  } catch (e) {
    return false;
  }
}

function formatPercentColumns(sheetName, headerNames) {
  const sh = getSheet(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeString);
  headerNames.forEach(name => {
    const idx = headers.indexOf(name) + 1;
    if (idx > 0) sh.getRange(2, idx, sh.getLastRow() - 1, 1).setNumberFormat('0.00%');
  });
}

function formatDashboard() {
  const sh = getSheet(SHEETS.DASHBOARD);
  if (!sh) return;
  formatSheetBase(sh, HEADERS.DASHBOARD.length);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 3, sh.getLastRow() - 1, 2).setNumberFormat('#,##0.00');
  }
}

function applyConditionalFormattingToUnitSheets() {
  applyRowColors(SHEETS.UNIT_FACT, 'Чистая_прибыль', 'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%');
  applyRowColors(SHEETS.UNIT_FORECAST, 'Чистая_прибыль_прогноз', 'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%');
  beautifyWorkbook();
}

function applyRowColors(sheetName, profitHeader, marginHeader, drrHeader, ozonHeader) {
  const sh = getSheet(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  const settings = getSettingsMap();
  const targetMargin = num(settings.target_margin_pct);
  const criticalDrr = num(settings.critical_ad_share_pct);
  const criticalOzon = num(settings.critical_ozon_expense_share_pct);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(safeString);
  const pIdx = headers.indexOf(profitHeader);
  const mIdx = headers.indexOf(marginHeader);
  const dIdx = headers.indexOf(drrHeader);
  const oIdx = headers.indexOf(ozonHeader);
  const backgrounds = [];

  for (let r = 1; r < values.length; r++) {
    const profit = num(values[r][pIdx]);
    const margin = num(values[r][mIdx]);
    const drr = num(values[r][dIdx]);
    const ozon = num(values[r][oIdx]);
    let color = '#ffffff';
    if (profit < 0) color = '#fce8e6';
    else if (drr > criticalDrr) color = '#fef7e0';
    else if (ozon > criticalOzon) color = '#fff4e5';
    else if (margin < targetMargin) color = '#fffde7';
    else color = '#e6f4ea';
    backgrounds.push(new Array(values[0].length).fill(color));
  }
  sh.getRange(2, 1, backgrounds.length, values[0].length).setBackgrounds(backgrounds);
}

function filterRowsByDate(rows, field, from, to) {
  return rows.filter(r => {
    const d = parseSheetDate(r[field]);
    if (!d) return false;
    return d >= startOfDay(from) && d <= endOfDay(to);
  });
}

function splitDateRangeByMonth(from, to) {
  const chunks = [];
  let cursor = startOfDay(from);
  const end = endOfDay(to);
  while (cursor <= end) {
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0, 23, 59, 59, 999);
    chunks.push({ from: new Date(cursor), to: monthEnd < end ? monthEnd : new Date(end) });
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }
  return chunks;
}

function splitDateRangeByDays(from, to, maxDays) {
  const chunks = [];
  let cursor = startOfDay(from);
  const end = endOfDay(to);
  const days = Math.max(1, Math.min(28, Math.round(num(maxDays) || 28)));

  while (cursor <= end) {
    const chunkTo = endOfDay(addDays(cursor, days - 1));
    chunks.push({
      from: new Date(cursor),
      to: chunkTo < end ? chunkTo : new Date(end)
    });
    cursor = startOfDay(addDays(chunkTo, 1));
  }
  return chunks;
}

function toOzonDateTime(date, end) {
  // Важно: не форматируем локальную дату через GMT, иначе 01.03 может стать 28.02T21:00Z
  // и Ozon воспринимает период как длиннее месяца.
  const d = new Date(date);
  const y = d.getFullYear();
  const m = d.getMonth();
  const day = d.getDate();
  const utc = end
    ? new Date(Date.UTC(y, m, day, 23, 59, 59, 999))
    : new Date(Date.UTC(y, m, day, 0, 0, 0, 0));
  return Utilities.formatDate(utc, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
}

function dateToYmd(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseSheetDate(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400 * 1000));
  const s = safeString(value).trim();
  if (!s) return null;
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) {
    d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function startOfDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function endOfDay(date) {
  const d = new Date(date);
  d.setHours(23, 59, 59, 999);
  return d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function isCancelledStatus(status) {
  const s = safeString(status).toLowerCase();
  return s.indexOf('cancel') >= 0 || s.indexOf('отмен') >= 0 || s.indexOf('canceled') >= 0 || s.indexOf('cancelled') >= 0;
}

function isReturnOperation(operationType, operationName) {
  const s = (safeString(operationType) + ' ' + safeString(operationName)).toLowerCase();
  return s.indexOf('return') >= 0 || s.indexOf('возврат') >= 0 || s.indexOf('returned') >= 0;
}

function topRows(rows, field, n) {
  return rows.slice().sort((a, b) => num(b[field]) - num(a[field])).slice(0, n);
}

function bottomRows(rows, field, n) {
  return rows.slice().sort((a, b) => num(a[field]) - num(b[field])).slice(0, n);
}

function beautifyWorkbook() {
  stylePeriodSheet();
  styleSettingsSheet();
  styleCostsSheet();
  styleApiSheets();
  styleUnitSheet(SHEETS.UNIT_FACT);
  styleUnitSheet(SHEETS.UNIT_FORECAST);
  styleDashboard();
}

function stylePeriodSheet() {
  const sh = getSheet(SHEETS.PERIOD);
  if (!sh) return;
  sh.setHiddenGridlines(true);
  sh.getRange('A1:D1').setBackground('#0b57d0').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('B4:B5').setBackground('#fff2cc').setNumberFormat('dd.mm.yyyy');
  sh.getRange('B7').setBackground('#fff2cc').setNumberFormat('0.00%');
  sh.getRange('B8').setBackground('#fff2cc');
}

function styleSettingsSheet() {
  const sh = getSheet(SHEETS.SETTINGS);
  if (!sh) return;
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 1, 240);
  sh.setColumnWidths(2, 1, 180);
  sh.setColumnWidths(3, 1, 520);
  sh.getRange('A1:C1').setBackground('#0f9d58').setFontColor('#ffffff').setFontWeight('bold');
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 2, sh.getLastRow() - 1, 1).setBackground('#fff2cc');
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).setBorder(true, true, true, true, true, true, '#d9ead3', SpreadsheetApp.BorderStyle.SOLID);
  }
}

function styleCostsSheet() {
  const sh = getSheet(SHEETS.COSTS);
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeString);
  const manualCols = ['Себестоимость_шт', 'Упаковка_шт', 'Доп_расход_шт', 'Налог_%_override', 'Целевая_маржа_%_override', 'Средний_%_отмен_override', 'Средний_%_возвратов_override', 'Категория', 'Комментарий'];
  const autoCols = ['offer_id', 'sku', 'Ссылка_Ozon', 'Название'];
  autoCols.forEach(name => colorColumnByHeader(sh, headers, name, '#d9ead3'));
  manualCols.forEach(name => colorColumnByHeader(sh, headers, name, '#fff2cc'));
  sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground('#38761d').setFontColor('#ffffff');
  setColumnWidthsByHeader(sh, headers, {
    'offer_id': 170,
    'sku': 120,
    'Ссылка_Ozon': 220,
    'Название': 320,
    'Себестоимость_шт': 130,
    'Упаковка_шт': 120,
    'Доп_расход_шт': 130,
    'Комментарий': 260
  });
}

function styleApiSheets() {
  [SHEETS.PRODUCTS, SHEETS.ORDERS, SHEETS.FINANCE, SHEETS.ADS, SHEETS.HISTORY, SHEETS.LOGS].forEach(name => {
    const sh = getSheet(name);
    if (!sh) return;
    sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground('#6d9eeb').setFontColor('#ffffff');
  });
}

function styleUnitSheet(sheetName) {
  const sh = getSheet(sheetName);
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeString);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground('#0b57d0').setFontColor('#ffffff');
  ['Чистая_прибыль', 'Чистая_прибыль_прогноз', 'Маржа_%', 'ДРР_%', 'Расходы_Ozon_%', 'Статус', 'Рекомендация'].forEach(name => colorColumnByHeader(sh, headers, name, '#e8f0fe'));
  setColumnWidthsByHeader(sh, headers, {
    'offer_id': 170,
    'sku': 120,
    'Название': 320,
    'Рекомендация': 420,
    'Статус': 100
  });
}

function styleDashboard() {
  const sh = getSheet(SHEETS.DASHBOARD);
  if (!sh) return;
  sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground('#674ea7').setFontColor('#ffffff').setFontWeight('bold');
  sh.setColumnWidths(1, 1, 180);
  sh.setColumnWidths(2, 1, 260);
  sh.setColumnWidths(3, 2, 150);
  sh.setColumnWidths(5, 1, 420);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).setBorder(true, true, true, true, true, true, '#d9d2e9', SpreadsheetApp.BorderStyle.SOLID);
  }
}

function styleStockValuationSheet() {
  const sh = getSheet(SHEETS.STOCK_VALUE);
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(safeString);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground('#134f5c').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  ['Остаток по себесу', 'Потенциальный оборот', 'Прогнозная прибыль с остатка'].forEach(name => colorColumnByHeader(sh, headers, name, '#d9ead3'));
  ['Себестоимость за шт', 'Текущая цена'].forEach(name => colorColumnByHeader(sh, headers, name, '#fff2cc'));
  setColumnWidthsByHeader(sh, headers, {
    'Артикул': 170,
    'SKU Ozon': 120,
    'Название': 320,
    'Комментарий': 380
  });
}

function colorColumnByHeader(sh, headers, header, color) {
  const idx = headers.indexOf(header) + 1;
  if (idx <= 0 || sh.getLastRow() < 2) return;
  sh.getRange(2, idx, sh.getLastRow() - 1, 1).setBackground(color);
}

function setColumnWidthsByHeader(sh, headers, widthByHeader) {
  Object.keys(widthByHeader).forEach(name => {
    const idx = headers.indexOf(name) + 1;
    if (idx > 0) sh.setColumnWidth(idx, widthByHeader[name]);
  });
}

function buildOzonProductUrl(sku, productId) {
  const id = safeString(productId || sku);
  return id ? 'https://www.ozon.ru/product/' + encodeURIComponent(id) + '/' : '';
}

function pickNumOrBlank() {
  for (let i = 0; i < arguments.length; i++) {
    const v = arguments[i];
    if (v !== null && v !== undefined && v !== '') return normalizeOzonMoney(v);
  }
  return '';
}

function normalizeOzonMoney(v) {
  const n = num(v);
  if (!n) return 0;
  const settings = getSettingsMap ? getSettingsMap() : {};
  const autoFix = settings.auto_fix_kopeck_prices === undefined ? true : parseBool(settings.auto_fix_kopeck_prices);
  // Иногда из API/отчетов цена может прийти в копейках: 80000 вместо 800.
  // Для недорогих товаров это типичная причина ошибки цены x100.
  if (autoFix && n >= 10000 && Math.round(n) === n && n % 100 === 0) return n / 100;
  return n;
}

function estimateDailySales(offerId) {
  const rows = readSheetAsObjects(SHEETS.ORDERS);
  const period = getSyncPeriodWithHistory();
  let qty = 0;
  rows.forEach(r => {
    if (safeString(r.offer_id) !== safeString(offerId)) return;
    if (isCancelledStatus(r.status)) return;
    const d = parseSheetDate(r.date);
    if (!d || d < period.from || d > period.to) return;
    qty += num(r.quantity);
  });
  const days = Math.max(1, Math.ceil((period.to - period.from) / 86400000));
  return qty / days;
}

function stockStatus(stock, daysLeft) {
  if (stock <= 0) return 'Нет остатка';
  if (daysLeft !== '' && daysLeft < 14) return 'Скоро закончится';
  if (daysLeft !== '' && daysLeft > 120) return 'Избыток';
  return 'ОК';
}

function stockComment(stock, daysLeft, costPerUnit, price) {
  if (!costPerUnit) return 'Заполни себестоимость — стоимость остатка считается от нее.';
  if (!price) return 'Не найдена цена — проверь API_Товары или наличие продаж за период.';
  if (stock <= 0) return 'Товар закончился или остаток не подтянулся из API.';
  if (daysLeft !== '' && daysLeft < 14) return 'Остатка меньше чем на 14 дней продаж — стоит планировать поставку.';
  if (daysLeft !== '' && daysLeft > 120) return 'Остаток больше чем на 120 дней — деньги заморожены в товаре.';
  return 'Остаток выглядит нормально.';
}

function safeString(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function safeArray(v) {
  return Array.isArray(v) ? v : [];
}

function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseLocaleNumber(v);
  return isFinite(n) ? n : 0;
}

function parseLocaleNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v)
    .replace(/\s/g, '')
    .replace(/%/g, '')
    .replace(',', '.')
    .replace(/[₽руб.]/gi, '');
  const n = Number(s);
  if (!isFinite(n)) return 0;
  if (String(v).indexOf('%') >= 0) return n / 100;
  return n;
}

function toRate(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseLocaleNumber(v);
  if (Math.abs(n) > 1) return n / 100;
  return n;
}

function parseBool(v) {
  if (v === true) return true;
  const s = safeString(v).toLowerCase();
  return ['true', 'yes', 'y', '1', 'да', 'истина'].indexOf(s) >= 0;
}

function safeDiv(a, b) {
  a = num(a);
  b = num(b);
  if (!b) return 0;
  return a / b;
}

function clamp(v, min, max) {
  return Math.max(min, Math.min(max, num(v)));
}

function toPercent(v) {
  return (num(v) * 100).toFixed(2) + '%';
}

function unique(arr) {
  const seen = {};
  const out = [];
  arr.forEach(x => {
    const key = safeString(x);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
}

function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function sumObjectNumbers(obj) {
  if (!obj || typeof obj !== 'object') return 0;
  return Object.keys(obj).reduce((s, k) => s + num(obj[k]), 0);
}

function firstNumber(arr, key) {
  if (!Array.isArray(arr)) return '';
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] && arr[i][key]) return arr[i][key];
  }
  return '';
}

function logInfo(scope, message) {
  writeLog('INFO', scope, message);
}

function logWarn(scope, message) {
  writeLog('WARN', scope, message);
}

function writeLog(level, scope, message) {
  try {
    const sh = ensureSheetWithHeader(SHEETS.LOGS, HEADERS.LOGS);
    sh.appendRow([new Date(), level, scope, message]);
  } catch (e) {
    Logger.log(level + ' ' + scope + ': ' + message);
  }
}
