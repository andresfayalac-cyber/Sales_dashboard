const CONFIG = {
  SPREADSHEET_ID: '1_r6A8nSJrr7c_CLtg8CdR8Fi3OQbbsS7SzDjIVrEoig', 
  SPREADSHEET_NAME: 'Spreadsheet_AppSheet_Ready',
  CACHE_EXPIRATION: 300 
};

const DATA_TIME_LIMITS = {
  CUTOFF_DATE: new Date(2026, 1, 28, 23, 59, 59, 999),
  CUTOFF_MONTH_START: new Date(2026, 1, 1),
  TARGETS_START_MONTH: new Date(2024, 0, 1)
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Data Science Portfolio Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    try { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); } catch (e) { Logger.log("Fallo al abrir por ID"); }
  }
  var files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.openById(file.getId());
  }
  throw new Error("No se encontró el spreadsheet: " + CONFIG.SPREADSHEET_NAME);
}

function getSheetDataAsObjects_(sheetName) {
  var ss = openSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("No se encontró la hoja: " + sheetName);

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var header = safeString_(headers[j]);
      if (header) obj[header] = row[j];
    }
    result.push(obj);
  }
  return result;
}

function safeString_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function safeNumber_(value) {
  if (typeof value === 'number') return value;
  var parsed = parseFloat(value);
  return isNaN(parsed) ? 0 : parsed;
}

function parseCurrencyEs_(value) {
  if (typeof value === 'number') return value;
  var str = safeString_(value);
  if (!str) return 0;
  str = str.replace(/\$/g, '').replace(/\./g, '').replace(/,/g, '.').trim();
  return safeNumber_(str);
}

function parsePercentEs_(value) {
  if (typeof value === 'number') return value > 1 ? value / 100 : value; 
  var str = safeString_(value);
  if (!str) return 0;
  str = str.replace(/%/g, '').replace(/\./g, '').replace(/,/g, '.').trim();
  return safeNumber_(str) / 100;
}

function parseDateEs_(value) {
  if (value instanceof Date) return value.toISOString();
  var str = safeString_(value);
  if (!str) return "";

  var direct = new Date(str);
  if (!isNaN(direct.getTime())) return direct.toISOString();

  var strLower = str.toLowerCase();
  var months = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };
  var parts = strLower.split('-');
  if (parts.length === 3 && months[parts[1]] !== undefined) {
    var d = parseInt(parts[0], 10);
    var y = parseInt(parts[2], 10);
    if (!isNaN(d) && !isNaN(y)) {
      return new Date(y, months[parts[1]], d).toISOString();
    }
  }
  return str; 
}

function parseMonthEs_(value) {
  if (value instanceof Date) return value.toISOString();
  var str = safeString_(value).toLowerCase();
  var months = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };
  var parts = str.split('-');
  if (parts.length === 2) {
    var m = months[parts[0]] !== undefined ? months[parts[0]] : 0;
    var y = parseInt(parts[1], 10);
    return new Date(y, m, 1).toISOString();
  }
  return str; 
}

function toDateOrNull_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }

  var str = safeString_(value);
  if (!str) return null;

  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;

  var ym = str.match(/^(\d{4})-(\d{2})$/);
  if (ym) {
    var year = parseInt(ym[1], 10);
    var month = parseInt(ym[2], 10) - 1;
    if (!isNaN(year) && month >= 0 && month <= 11) {
      return new Date(year, month, 1);
    }
  }

  return null;
}

function toMonthStartOrNull_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return new Date(value.getFullYear(), value.getMonth(), 1);
  }

  var str = safeString_(value);
  if (!str) return null;

  var strLower = str.toLowerCase();
  var monthsEs = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };
  var esMatch = strLower.match(/^([a-z]{3})-(\d{4})$/);
  if (esMatch && monthsEs.hasOwnProperty(esMatch[1])) {
    return new Date(parseInt(esMatch[2], 10), monthsEs[esMatch[1]], 1);
  }

  var isoMonthMatch = str.match(/^(\d{4})-(\d{2})(?:-(\d{2}))?(?:[tT ].*)?$/);
  if (isoMonthMatch) {
    var yearIso = parseInt(isoMonthMatch[1], 10);
    var monthIso = parseInt(isoMonthMatch[2], 10) - 1;
    if (!isNaN(yearIso) && monthIso >= 0 && monthIso <= 11) {
      return new Date(yearIso, monthIso, 1);
    }
  }

  var parsed = toDateOrNull_(str);
  if (!parsed) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), 1);
}

function isWithinCutoffMonth_(value) {
  var monthDate = value instanceof Date ? new Date(value.getFullYear(), value.getMonth(), 1) : toMonthStartOrNull_(value);
  if (!monthDate) return false;
  return monthDate.getTime() <= DATA_TIME_LIMITS.CUTOFF_MONTH_START.getTime();
}

function isWithinTargetWindow_(value) {
  var monthDate = value instanceof Date ? new Date(value.getFullYear(), value.getMonth(), 1) : toMonthStartOrNull_(value);
  if (!monthDate) return false;
  return monthDate.getTime() >= DATA_TIME_LIMITS.TARGETS_START_MONTH.getTime() &&
    monthDate.getTime() <= DATA_TIME_LIMITS.CUTOFF_MONTH_START.getTime();
}

function monthKey_(monthDate) {
  return monthDate.getFullYear() + '-' + ('0' + (monthDate.getMonth() + 1)).slice(-2);
}

function monthStamp_(monthDate) {
  return monthKey_(monthDate) + '-01T12:00:00';
}

function buildRecentMonthKeyMap_(latestMonthDate, monthsCount) {
  var safeLatest = latestMonthDate instanceof Date && !isNaN(latestMonthDate.getTime())
    ? new Date(latestMonthDate.getFullYear(), latestMonthDate.getMonth(), 1)
    : new Date(DATA_TIME_LIMITS.CUTOFF_MONTH_START.getFullYear(), DATA_TIME_LIMITS.CUTOFF_MONTH_START.getMonth(), 1);

  var count = Math.max(1, safeNumber_(monthsCount) || 6);
  var monthMap = {};

  for (var i = 0; i < count; i++) {
    var d = new Date(safeLatest.getFullYear(), safeLatest.getMonth() - i, 1);
    monthMap[monthKey_(d)] = true;
  }

  return monthMap;
}

function dayKey_(dayDate) {
  return dayDate.getFullYear() + '-' + ('0' + (dayDate.getMonth() + 1)).slice(-2) + '-' + ('0' + dayDate.getDate()).slice(-2);
}

function dayStamp_(dayDate) {
  return dayKey_(dayDate) + 'T12:00:00';
}

function toDayDateOrNull_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  var str = safeString_(value);
  if (!str) return null;

  var direct = toDateOrNull_(str);
  if (direct) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }

  var esDate = parseDateEs_(str);
  var parsedEs = toDateOrNull_(esDate);
  if (parsedEs) {
    return new Date(parsedEs.getFullYear(), parsedEs.getMonth(), parsedEs.getDate());
  }

  var monthDate = toMonthStartOrNull_(str);
  if (monthDate) {
    return new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  }

  return null;
}

function monthSortValue_(value) {
  var monthDate = value instanceof Date ? new Date(value.getFullYear(), value.getMonth(), 1) : toMonthStartOrNull_(value);
  return monthDate ? monthDate.getTime() : -Infinity;
}

function parseOptionalCurrency_(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string' && value.trim() === '') return null;
  return parseCurrencyEs_(value);
}

function detectMonthlyTargetColumns_(keys) {
  var list = Array.isArray(keys) ? keys : [];
  var cols = {
    sales: null,
    profit: null,
    generic: null
  };

  cols.sales = list.find(function(k) {
    if (!k) return false;
    return /((target|objetivo|meta).*(sales|revenue|venta|ventas|ingreso))|((sales|revenue|venta|ventas|ingreso).*(target|objetivo|meta))/i.test(k);
  }) || null;

  cols.profit = list.find(function(k) {
    if (!k) return false;
    return /((target|objetivo|meta).*(profit|utilidad|net|ganancia))|((profit|utilidad|net|ganancia).*(target|objetivo|meta))/i.test(k);
  }) || null;

  cols.generic = list.find(function(k) {
    if (!k) return false;
    if (!/target|objetivo|meta/i.test(k)) return false;
    if (/attainment|cumpl|ratio|pct|percent|porcentaje/i.test(k)) return false;
    if (/profit|utilidad|net|ganancia|margin|margen|cost|costo|expense|gasto/i.test(k)) return false;
    return true;
  }) || null;

  if (!cols.sales) cols.sales = cols.generic;

  return cols;
}

function withCache_(key, fetchFunction) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  try {
    var data = fetchFunction();
    try {
      cache.put(key, JSON.stringify(data), CONFIG.CACHE_EXPIRATION);
    } catch (cacheEx) {
      Logger.log("Cache error (likely exceeded 100KB limit): " + cacheEx.message);
    }
    return data;
  } catch (e) {
    Logger.log("Error en " + key + ": " + e.message);
    return { error: e.message };
  }
}

function getOverviewData() {
  return withCache_('overview_data', function() {
    var raw = getSheetDataAsObjects_("kpi_summary");
    if (raw.length === 0) throw new Error("La hoja kpi_summary está vacía.");

    var kpis = {};
    if (raw[0].hasOwnProperty('total_sales')) {
      kpis = raw[0];
    } else {
      var keys = Object.keys(raw[0]);
      raw.forEach(function(row) {
        var k = safeString_(row[keys[0]]);
        var v = row[keys[1]];
        if (k) kpis[k] = v;
      });
    }
    
    // Traemos la data mensual para poder hacer el filtro de tiempo en el frontend
    var monthlyRaw = getSheetDataAsObjects_("agg_monthly");
    var monthlyTargetCols = monthlyRaw.length > 0
      ? detectMonthlyTargetColumns_(Object.keys(monthlyRaw[0]))
      : { sales: null, profit: null };

    var monthlyData = monthlyRaw.map(function(row) {
      var monthDate = toMonthStartOrNull_(row.month_start || row.year_month);
      if (!monthDate || !isWithinCutoffMonth_(monthDate)) return null;

      var targetSalesValue = null;
      var targetProfitValue = null;
      if (isWithinTargetWindow_(monthDate)) {
        if (monthlyTargetCols.sales) {
          targetSalesValue = parseOptionalCurrency_(row[monthlyTargetCols.sales]);
        }
        if (monthlyTargetCols.profit) {
          targetProfitValue = parseOptionalCurrency_(row[monthlyTargetCols.profit]);
        }
      }

      return {
        month_start: monthStamp_(monthDate),
        sales: parseCurrencyEs_(row.sales),
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: parseCurrencyEs_(row.profit),
        tx_count: safeNumber_(row.tx_count),
        target_sales: targetSalesValue,
        target_profit: targetProfitValue,
        has_target_data: targetSalesValue !== null,
        has_profit_target_data: targetProfitValue !== null
      };
    }).filter(function(row) {
      return !!row;
    });

    monthlyData.sort(function(a, b) {
      return monthSortValue_(a.month_start) - monthSortValue_(b.month_start);
    });
    var latestAvailableMonth = monthlyData.length > 0 ? monthlyData[monthlyData.length - 1].month_start : null;

    return {
      total_sales: parseCurrencyEs_(kpis.total_sales),
      total_cost: parseCurrencyEs_(kpis.total_cost),
      total_expenses: parseCurrencyEs_(kpis.total_expenses),
      total_margin: parseCurrencyEs_(kpis.total_margin),
      total_profit: parseCurrencyEs_(kpis.total_profit),
      margin_pct: parsePercentEs_(kpis.margin_pct),
      profit_pct: parsePercentEs_(kpis.profit_pct),
      total_transactions: safeNumber_(kpis.total_transactions),
      average_ticket: parseCurrencyEs_(kpis.average_ticket),
      top_location: safeString_(kpis.top_location),
      top_product: safeString_(kpis.top_product),
      top_entity: safeString_(kpis.top_entity),
      monthly_data: monthlyData, // Enviado para el filtro
      latest_available_month: latestAvailableMonth
    };
  });
}

function getMonthlyData() {
  return withCache_('monthly_data', function() {
    var raw = getSheetDataAsObjects_("agg_monthly");
    var monthlyTargetCols = raw.length > 0
      ? detectMonthlyTargetColumns_(Object.keys(raw[0]))
      : { sales: null, profit: null };
    
    var processed = raw.map(function(row) {
      var monthDate = toMonthStartOrNull_(row.month_start || row.year_month);
      if (!monthDate || !isWithinCutoffMonth_(monthDate)) return null;

      var sales = parseCurrencyEs_(row.sales);
      var profit = parseCurrencyEs_(row.profit);
      var tx_count = safeNumber_(row.tx_count);
      var targetSalesValue = null;
      var targetProfitValue = null;
      if (isWithinTargetWindow_(monthDate)) {
        if (monthlyTargetCols.sales) {
          targetSalesValue = parseOptionalCurrency_(row[monthlyTargetCols.sales]);
        }
        if (monthlyTargetCols.profit) {
          targetProfitValue = parseOptionalCurrency_(row[monthlyTargetCols.profit]);
        }
      }
      
      return {
        month_start: monthStamp_(monthDate),
        year_month: monthKey_(monthDate),
        sales: sales,
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: profit,
        tx_count: tx_count,
        avg_margin_pct: parsePercentEs_(row.avg_margin_pct),
        avg_profit_pct: parsePercentEs_(row.avg_profit_pct),
        avg_ticket: tx_count > 0 ? sales / tx_count : 0,
        target_sales: targetSalesValue,
        target_profit: targetProfitValue,
        has_target_data: targetSalesValue !== null,
        has_profit_target_data: targetProfitValue !== null
      };
    }).filter(function(row) {
      return !!row;
    });

    processed.sort(function(a, b) { return monthSortValue_(a.month_start) - monthSortValue_(b.month_start); });

    var best_month_by_sales = null;
    var best_month_by_profit = null;
    var max_sales = -Infinity;
    var max_profit = -Infinity;
    var cumulative_sales = 0;
    var cumulative_profit = 0;

    processed.forEach(function(item) {
      cumulative_sales += item.sales;
      cumulative_profit += item.profit;
      item.cumulative_sales = cumulative_sales;
      item.cumulative_profit = cumulative_profit;

      if (item.sales > max_sales) { max_sales = item.sales; best_month_by_sales = item.month_start; }
      if (item.profit > max_profit) { max_profit = item.profit; best_month_by_profit = item.month_start; }
    });

    return {
      table_data: processed,
      categories: processed.map(function(d) { return d.month_start; }),
      series_sales: processed.map(function(d) { return d.sales; }),
      series_profit: processed.map(function(d) { return d.profit; }),
      series_cumulative_sales: processed.map(function(d) { return d.cumulative_sales; }),
      kpis: {
        best_month_by_sales: best_month_by_sales,
        best_month_by_profit: best_month_by_profit,
        total_cumulative_sales: cumulative_sales,
        total_cumulative_profit: cumulative_profit
      }
    };
  });
}

function getTimeSeriesData() {
  return withCache_('time_series_data', function() {
    var raw = getSheetDataAsObjects_("fact_sales_clean");
    if (!raw || raw.length === 0) {
      return {
        daily_data: [],
        has_target_data: false,
        cutoff_date: DATA_TIME_LIMITS.CUTOFF_DATE.toISOString(),
        target_window_start: monthStamp_(DATA_TIME_LIMITS.TARGETS_START_MONTH),
        target_window_end: monthStamp_(DATA_TIME_LIMITS.CUTOFF_MONTH_START)
      };
    }

    var keys = Object.keys(raw[0]);
    var colDate = keys.find(function(k) { return k.match(/date|fecha|day|month|mes|period/i); }) || keys[0];
    var colSales = keys.find(function(k) { return k.match(/sales|revenue|ventas|amount/i); }) || "sales";
    var colCost = keys.find(function(k) { return k.match(/cost|costo/i); }) || "cost";
    var colExpenses = keys.find(function(k) { return k.match(/expenses|expense|gasto/i); }) || null;
    var colMargin = keys.find(function(k) { return k.match(/margin|margen/i); }) || "margin";
    var colProfit = keys.find(function(k) { return k.match(/profit|utilidad|net/i); }) || "profit";
    var colTarget = keys.find(function(k) { return k.match(/target|objetivo|meta/i); }) || null;

    var dayMap = {};

    raw.forEach(function(r) {
      if (!r[colDate]) return;

      var dayDate = toDayDateOrNull_(r[colDate]);
      if (!dayDate) return;
      if (dayDate.getTime() > DATA_TIME_LIMITS.CUTOFF_DATE.getTime()) return;

      var monthDate = new Date(dayDate.getFullYear(), dayDate.getMonth(), 1);
      var dayKey = dayKey_(dayDate);

      if (!dayMap[dayKey]) {
        dayMap[dayKey] = {
          date: dayStamp_(dayDate),
          sales: 0,
          cost: 0,
          expenses: 0,
          margin: 0,
          profit: 0,
          tx_count: 0,
          target_sales: null,
          has_target_data: false
        };
      }

      var item = dayMap[dayKey];
      item.sales += parseCurrencyEs_(r[colSales]);
      item.cost += parseCurrencyEs_(r[colCost]);
      item.expenses += colExpenses ? parseCurrencyEs_(r[colExpenses]) : 0;
      item.margin += parseCurrencyEs_(r[colMargin]);
      item.profit += parseCurrencyEs_(r[colProfit]);
      item.tx_count += 1;

      var targetValue = null;
      if (colTarget && isWithinTargetWindow_(monthDate)) {
        targetValue = parseOptionalCurrency_(r[colTarget]);
      }

      if (targetValue !== null) {
        if (item.target_sales === null) item.target_sales = 0;
        item.target_sales += targetValue;
        item.has_target_data = true;
      }
    });

    var dailyData = Object.keys(dayMap).sort(function(a, b) {
      return a.localeCompare(b);
    }).map(function(key) {
      return dayMap[key];
    });

    return {
      daily_data: dailyData,
      has_target_data: dailyData.some(function(row) { return row.has_target_data; }),
      cutoff_date: DATA_TIME_LIMITS.CUTOFF_DATE.toISOString(),
      target_window_start: monthStamp_(DATA_TIME_LIMITS.TARGETS_START_MONTH),
      target_window_end: monthStamp_(DATA_TIME_LIMITS.CUTOFF_MONTH_START)
    };
  });
}

function getEntityData() {
  return withCache_('entity_data', function() {
    var raw = getSheetDataAsObjects_("agg_entity");

    var rawRows = raw.map(function(row) {
      var monthDate = row.month_start ? toMonthStartOrNull_(row.month_start) : null;
      if (monthDate && !isWithinCutoffMonth_(monthDate)) return null;

      var lastSaleDate = parseDateEs_(row.last_sale_date);
      var lastSaleParsed = toDateOrNull_(lastSaleDate);
      if (!monthDate && lastSaleParsed && lastSaleParsed.getTime() > DATA_TIME_LIMITS.CUTOFF_DATE.getTime()) {
        return null;
      }
      if (lastSaleParsed && lastSaleParsed.getTime() > DATA_TIME_LIMITS.CUTOFF_DATE.getTime()) {
        lastSaleDate = DATA_TIME_LIMITS.CUTOFF_DATE.toISOString();
      }

      return {
        entity_name: safeString_(row.entity_name || row.entity || row.seller || row.seller_name),
        month_start: monthDate ? monthStamp_(monthDate) : '',
        sales: parseCurrencyEs_(row.sales),
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: parseCurrencyEs_(row.profit),
        tx_count: safeNumber_(row.tx_count),
        avg_margin_pct: parsePercentEs_(row.avg_margin_pct),
        avg_profit_pct: parsePercentEs_(row.avg_profit_pct),
        last_sale_date: lastSaleDate
      };
    }).filter(function(row) {
      return !!(row && row.entity_name);
    });

    var entityMap = {};
    rawRows.forEach(function(row) {
      if (!entityMap[row.entity_name]) {
        entityMap[row.entity_name] = {
          entity_name: row.entity_name,
          sales: 0,
          cost: 0,
          expenses: 0,
          margin: 0,
          profit: 0,
          tx_count: 0,
          avg_margin_pct: 0,
          avg_profit_pct: 0,
          last_sale_date: row.last_sale_date
        };
      }

      var item = entityMap[row.entity_name];
      item.sales += row.sales;
      item.cost += row.cost;
      item.expenses += row.expenses;
      item.margin += row.margin;
      item.profit += row.profit;
      item.tx_count += row.tx_count;

      var existingDate = new Date(item.last_sale_date);
      var rowDate = new Date(row.last_sale_date);
      if (!isNaN(rowDate.getTime()) && (isNaN(existingDate.getTime()) || rowDate > existingDate)) {
        item.last_sale_date = row.last_sale_date;
      }
    });

    var aggregatedRows = Object.keys(entityMap).map(function(key) {
      var item = entityMap[key];
      item.avg_margin_pct = item.sales > 0 ? item.margin / item.sales : 0;
      item.avg_profit_pct = item.sales > 0 ? item.profit / item.sales : 0;
      return item;
    });

    return {
      table_data: aggregatedRows,
      raw_rows: rawRows,
      has_time_dimension: rawRows.some(function(row) { return !!row.month_start; })
    };
  });
}

function getProductLocationData() {
  return withCache_('product_location_data', function() {
    var raw = getSheetDataAsObjects_("fact_sales_clean");
    if (!raw || raw.length === 0) return { error: "No data in fact_sales_clean" };
    
    var keys = Object.keys(raw[0]);
    var colDate = keys.find(function(k) { return k.match(/date|fecha|month|mes|period/i); }) || keys[0];
    var colProduct = keys.find(function(k) { return k.match(/product|producto|item/i); }) || "product";
    var colCategory = keys.find(function(k) { return k.match(/category|categoria|line/i); }) || "category";
    var colLocation = keys.find(function(k) { return k.match(/location|region|tienda|store|ubicacion|ciudad/i); }) || "location";
    var colEntity = keys.find(function(k) { return k.match(/entity|seller|vendedor|rep|person/i); }) || "seller";
    
    var colSales = keys.find(function(k) { return k.match(/sales|revenue|ventas|amount/i); }) || "sales";
    var colCost = keys.find(function(k) { return k.match(/cost|costo/i); }) || "cost";
    var colMargin = keys.find(function(k) { return k.match(/margin|margen/i); }) || "margin";
    var colProfit = keys.find(function(k) { return k.match(/profit|utilidad|net/i); }) || "profit";
    var colTarget = keys.find(function(k) { return k.match(/target|objetivo|meta/i); }) || null;

    var prodRawMap = {};
    var locRawMap = {};
    var productMap = {};
    var locationMap = {};

    raw.forEach(function(r) {
      if (!r[colDate]) return;
      var monthDate = toMonthStartOrNull_(r[colDate]);
      if (!monthDate || !isWithinCutoffMonth_(monthDate)) return;

      var dStr = monthStamp_(monthDate);
      var inTargetWindow = isWithinTargetWindow_(monthDate);
      var targetValue = inTargetWindow && colTarget ? parseOptionalCurrency_(r[colTarget]) : null;
      
      var product = safeString_(r[colProduct]) || "Unknown";
      var category = safeString_(r[colCategory]) || "Uncategorized";
      var entity = safeString_(r[colEntity]) || "Unknown";
      var location = safeString_(r[colLocation]) || "Unknown";
      
      var sales = parseCurrencyEs_(r[colSales]);
      var cost = parseCurrencyEs_(r[colCost]);
      var margin = parseCurrencyEs_(r[colMargin]);
      var profit = parseCurrencyEs_(r[colProfit]);
      var tx_count = 1;

      // Agg for raw_rows
      var pKey = dStr + "|" + product + "|" + category + "|" + location + "|" + entity;
      if (!prodRawMap[pKey]) {
        prodRawMap[pKey] = { month_start: dStr, product: product, category: category, location: location, seller: entity, sales: 0, cost: 0, margin: 0, profit: 0, tx_count: 0 };
      }
      prodRawMap[pKey].sales += sales; prodRawMap[pKey].cost += cost; prodRawMap[pKey].margin += margin; prodRawMap[pKey].profit += profit; prodRawMap[pKey].tx_count += tx_count;

      var lKey = dStr + "|" + location;
      if (!locRawMap[lKey]) {
        locRawMap[lKey] = { month_start: dStr, location: location, sales: 0, cost: 0, margin: 0, profit: 0, tx_count: 0, sales_target_scope: 0, target_sales: null, target_attainment: null, has_target_data: false };
      }
      locRawMap[lKey].sales += sales; locRawMap[lKey].cost += cost; locRawMap[lKey].margin += margin; locRawMap[lKey].profit += profit; locRawMap[lKey].tx_count += tx_count;
      if (inTargetWindow) {
        locRawMap[lKey].sales_target_scope += sales;
        if (targetValue !== null) {
          if (locRawMap[lKey].target_sales === null) locRawMap[lKey].target_sales = 0;
          locRawMap[lKey].target_sales += targetValue;
          locRawMap[lKey].has_target_data = true;
        }
      }

      // Agg for list (all time)
      if (!productMap[product]) {
        productMap[product] = { product: product, category: category, sales: 0, cost: 0, margin: 0, profit: 0, tx_count: 0 };
      }
      productMap[product].sales += sales; productMap[product].cost += cost; productMap[product].margin += margin; productMap[product].profit += profit; productMap[product].tx_count += tx_count;
      
      if (!locationMap[location]) {
        locationMap[location] = { location: location, sales: 0, cost: 0, expenses: 0, margin: 0, profit: 0, sales_target_scope: 0, target_sales: null, target_attainment: null, target_gap: null, has_target_data: false };
      }
      locationMap[location].sales += sales; locationMap[location].cost += cost; locationMap[location].margin += margin; locationMap[location].profit += profit;
      if (inTargetWindow) {
        locationMap[location].sales_target_scope += sales;
        if (targetValue !== null) {
          if (locationMap[location].target_sales === null) locationMap[location].target_sales = 0;
          locationMap[location].target_sales += targetValue;
          locationMap[location].has_target_data = true;
        }
      }
    });

    var productRows = Object.keys(prodRawMap).map(function(k) { return prodRawMap[k]; });
    var locationRows = Object.keys(locRawMap).map(function(k) {
      var row = locRawMap[k];
      if (row.target_sales !== null && row.target_sales > 0) {
        row.target_attainment = row.sales_target_scope / row.target_sales;
      } else {
        row.target_attainment = null;
      }
      return row;
    });
    
    var products = Object.keys(productMap).map(function(k) {
      var item = productMap[k];
      item.avg_margin_pct = item.sales > 0 ? item.margin / item.sales : 0;
      item.avg_profit_pct = item.sales > 0 ? item.profit / item.sales : 0;
      return item;
    });
    
    var locations = Object.keys(locationMap).map(function(k) {
      var item = locationMap[k];
      if (item.target_sales !== null && item.target_sales > 0) {
        item.target_attainment = item.sales_target_scope / item.target_sales;
        item.target_gap = item.target_sales - item.sales_target_scope;
        item.has_target_data = true;
      } else {
        item.target_attainment = null;
        item.target_gap = null;
        item.has_target_data = false;
      }
      return item;
    });

    var categoryMap = {};
    products.forEach(function(item) {
      var c = item.category || 'Uncategorized';
      if (!categoryMap[c]) categoryMap[c] = { category: c, sales: 0, margin: 0, profit: 0 };
      categoryMap[c].sales += item.sales; categoryMap[c].margin += item.margin; categoryMap[c].profit += item.profit;
    });
    var category_distribution = Object.keys(categoryMap).map(function(k) { return categoryMap[k]; }).sort(function(a, b) { return b.sales - a.sales; });

    products.sort(function(a, b) { return b.sales - a.sales; });
    locations.sort(function(a, b) { return b.sales - a.sales; });

    var productsByProfit = products.slice().sort(function(a, b) { return b.profit - a.profit; });
    var locationsByAttainment = locations
      .filter(function(item) { return item.target_attainment !== null; })
      .sort(function(a, b) { return b.target_attainment - a.target_attainment; });
    
    var totalSalesLocTargetScope = 0;
    var totalTargetLoc = 0;
    locations.forEach(function(item) {
      if (item.target_sales !== null && item.target_sales > 0) {
        totalSalesLocTargetScope += item.sales_target_scope;
        totalTargetLoc += item.target_sales;
      }
    });

    var hasTargetData = totalTargetLoc > 0;

    return {
      products: {
        list: products,
        raw_rows: productRows,
        has_time_dimension: true,
        category_distribution: category_distribution,
        top_product_by_sales: products.length > 0 ? products[0].product : "N/A",
        top_product_by_profit: productsByProfit.length > 0 ? productsByProfit[0].product : "N/A",
        top_category_by_sales: category_distribution.length > 0 ? category_distribution[0].category : "N/A"
      },
      locations: {
        list: locations,
        raw_rows: locationRows,
        has_time_dimension: true,
        top_location_by_sales: locations.length > 0 ? locations[0].location : "N/A",
        top_location_by_target_attainment: locationsByAttainment.length > 0 ? locationsByAttainment[0].location : "N/A",
        total_target_gap: hasTargetData ? totalTargetLoc - totalSalesLocTargetScope : null,
        has_target_data: hasTargetData,
        target_window_start: monthStamp_(DATA_TIME_LIMITS.TARGETS_START_MONTH),
        target_window_end: monthStamp_(DATA_TIME_LIMITS.CUTOFF_MONTH_START)
      }
    };
  });
}

function getProductLocationStrategyData() {
  return withCache_('product_location_strategy_data', function() {
    var raw = getSheetDataAsObjects_("fact_sales_clean");
    if (!raw || raw.length === 0) return { error: "No data in fact_sales_clean" };

    var overview = getOverviewData();
    var latestMonthDate = overview && !overview.error ? toMonthStartOrNull_(overview.latest_available_month) : null;
    if (!latestMonthDate || !isWithinCutoffMonth_(latestMonthDate)) {
      latestMonthDate = new Date(DATA_TIME_LIMITS.CUTOFF_MONTH_START.getFullYear(), DATA_TIME_LIMITS.CUTOFF_MONTH_START.getMonth(), 1);
    }

    var scopeMonths = 6;
    var allowedMonthMap = buildRecentMonthKeyMap_(latestMonthDate, scopeMonths);

    var keys = Object.keys(raw[0]);
    var colDate = keys.find(function(k) { return k.match(/date|fecha|month|mes|period/i); }) || keys[0];
    var colProduct = keys.find(function(k) { return k.match(/product|producto|item/i); }) || "product";
    var colCategory = keys.find(function(k) { return k.match(/category|categoria|line/i); }) || "category";
    var colLocation = keys.find(function(k) { return k.match(/location|region|tienda|store|ubicacion|ciudad/i); }) || "location";

    var colSales = keys.find(function(k) { return k.match(/sales|revenue|ventas|amount/i); }) || "sales";
    var colCost = keys.find(function(k) { return k.match(/cost|costo/i); }) || "cost";
    var colMargin = keys.find(function(k) { return k.match(/margin|margen/i); }) || "margin";
    var colProfit = keys.find(function(k) { return k.match(/profit|utilidad|net/i); }) || "profit";

    var prodRawMap = {};
    var locRawMap = {};
    var productMap = {};
    var locationMap = {};

    raw.forEach(function(r) {
      if (!r[colDate]) return;
      var monthDate = toMonthStartOrNull_(r[colDate]);
      if (!monthDate || !isWithinCutoffMonth_(monthDate)) return;

      var monthKey = monthKey_(monthDate);
      if (!allowedMonthMap[monthKey]) return;

      var monthStamp = monthStamp_(monthDate);
      var product = safeString_(r[colProduct]) || "Unknown";
      var category = safeString_(r[colCategory]) || "Uncategorized";
      var location = safeString_(r[colLocation]) || "Unknown";

      var sales = parseCurrencyEs_(r[colSales]);
      var cost = parseCurrencyEs_(r[colCost]);
      var margin = parseCurrencyEs_(r[colMargin]);
      var profit = parseCurrencyEs_(r[colProfit]);

      var productRawKey = monthStamp + "|" + product + "|" + category + "|" + location;
      if (!prodRawMap[productRawKey]) {
        prodRawMap[productRawKey] = {
          month_start: monthStamp,
          product: product,
          category: category,
          location: location,
          sales: 0,
          cost: 0,
          margin: 0,
          profit: 0,
          tx_count: 0
        };
      }
      prodRawMap[productRawKey].sales += sales;
      prodRawMap[productRawKey].cost += cost;
      prodRawMap[productRawKey].margin += margin;
      prodRawMap[productRawKey].profit += profit;
      prodRawMap[productRawKey].tx_count += 1;

      var locationRawKey = monthStamp + "|" + location;
      if (!locRawMap[locationRawKey]) {
        locRawMap[locationRawKey] = {
          month_start: monthStamp,
          location: location,
          sales: 0,
          cost: 0,
          margin: 0,
          profit: 0,
          tx_count: 0
        };
      }
      locRawMap[locationRawKey].sales += sales;
      locRawMap[locationRawKey].cost += cost;
      locRawMap[locationRawKey].margin += margin;
      locRawMap[locationRawKey].profit += profit;
      locRawMap[locationRawKey].tx_count += 1;

      if (!productMap[product]) {
        productMap[product] = {
          product: product,
          category: category,
          sales: 0,
          cost: 0,
          margin: 0,
          profit: 0,
          tx_count: 0
        };
      }
      productMap[product].sales += sales;
      productMap[product].cost += cost;
      productMap[product].margin += margin;
      productMap[product].profit += profit;
      productMap[product].tx_count += 1;

      if (!locationMap[location]) {
        locationMap[location] = {
          location: location,
          sales: 0,
          cost: 0,
          margin: 0,
          profit: 0,
          tx_count: 0
        };
      }
      locationMap[location].sales += sales;
      locationMap[location].cost += cost;
      locationMap[location].margin += margin;
      locationMap[location].profit += profit;
      locationMap[location].tx_count += 1;
    });

    var productRows = Object.keys(prodRawMap).map(function(k) { return prodRawMap[k]; });
    var locationRows = Object.keys(locRawMap).map(function(k) { return locRawMap[k]; });

    var products = Object.keys(productMap).map(function(k) {
      var item = productMap[k];
      item.avg_margin_pct = item.sales > 0 ? item.margin / item.sales : 0;
      item.avg_profit_pct = item.sales > 0 ? item.profit / item.sales : 0;
      return item;
    });

    var locations = Object.keys(locationMap).map(function(k) {
      var item = locationMap[k];
      item.avg_margin_pct = item.sales > 0 ? item.margin / item.sales : 0;
      item.avg_profit_pct = item.sales > 0 ? item.profit / item.sales : 0;
      item.avg_ticket = item.tx_count > 0 ? item.sales / item.tx_count : 0;
      return item;
    });

    var categoryMap = {};
    products.forEach(function(item) {
      var c = item.category || 'Uncategorized';
      if (!categoryMap[c]) categoryMap[c] = { category: c, sales: 0, margin: 0, profit: 0 };
      categoryMap[c].sales += item.sales;
      categoryMap[c].margin += item.margin;
      categoryMap[c].profit += item.profit;
    });

    var categoryDistribution = Object.keys(categoryMap)
      .map(function(k) { return categoryMap[k]; })
      .sort(function(a, b) { return b.sales - a.sales; });

    products.sort(function(a, b) { return b.sales - a.sales; });
    locations.sort(function(a, b) { return b.sales - a.sales; });

    var productsByProfit = products.slice().sort(function(a, b) { return b.profit - a.profit; });

    return {
      products: {
        list: products,
        raw_rows: productRows,
        has_time_dimension: true,
        category_distribution: categoryDistribution,
        top_product_by_sales: products.length > 0 ? products[0].product : "N/A",
        top_product_by_profit: productsByProfit.length > 0 ? productsByProfit[0].product : "N/A",
        top_category_by_sales: categoryDistribution.length > 0 ? categoryDistribution[0].category : "N/A"
      },
      locations: {
        list: locations,
        raw_rows: locationRows,
        has_time_dimension: true,
        top_location_by_sales: locations.length > 0 ? locations[0].location : "N/A"
      },
      meta: {
        latest_month: monthStamp_(latestMonthDate),
        scope_months: scopeMonths
      }
    };
  });
}

function getAppBootstrapData() {
  return {
    overview: getOverviewData()
  };
}

function refreshAllDashboardData() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['overview_data', 'monthly_data', 'time_series_data', 'entity_data', 'product_location_data', 'product_location_strategy_data']);
  var res = {
    overview: getOverviewData(),
    monthly: getMonthlyData(),
    time_series: getTimeSeriesData(),
    entity: getEntityData(),
    product_location_strategy: getProductLocationStrategyData()
  };
  if (res.overview.error || res.monthly.error || res.time_series.error || res.entity.error || res.product_location_strategy.error) {
    return { success: false, message: "Hubo errores al refrescar algunos datos. Revisa el formato en Sheets." };
  }
  return { success: true, message: "Datos actualizados correctamente." };
}

function diagnosticarProblema() {
  var ss = openSpreadsheet_();
  var hojasEsperadas = ["agg_monthly", "agg_entity", "agg_product", "agg_location"];
  
  hojasEsperadas.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) {
      Logger.log("❌ ERROR: No se encontró la hoja '" + nombre + "'. Revisa el nombre de la pestaña.");
    } else {
      var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      Logger.log("✅ Hoja '" + nombre + "' encontrada. Columnas: " + headers.join(", "));
    }
  });
}
