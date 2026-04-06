const CONFIG = {
  SPREADSHEET_ID: '1_r6A8nSJrr7c_CLtg8CdR8Fi3OQbbsS7SzDjIVrEoig', 
  SPREADSHEET_NAME: 'Spreadsheet_AppSheet_Ready',
  CACHE_EXPIRATION: 300 
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
  var str = safeString_(value).toLowerCase();
  var months = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };
  var parts = str.split('-');
  if (parts.length === 3) {
    var d = parseInt(parts[0], 10);
    var m = months[parts[1]] !== undefined ? months[parts[1]] : 0;
    var y = parseInt(parts[2], 10);
    return new Date(y, m, d).toISOString();
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

function withCache_(key, fetchFunction) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  try {
    var data = fetchFunction();
    cache.put(key, JSON.stringify(data), CONFIG.CACHE_EXPIRATION);
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
    var monthlyData = monthlyRaw.map(function(row) {
      return {
        month_start: parseMonthEs_(row.month_start),
        sales: parseCurrencyEs_(row.sales),
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: parseCurrencyEs_(row.profit),
        tx_count: safeNumber_(row.tx_count)
      };
    });

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
      monthly_data: monthlyData // Enviado para el filtro
    };
  });
}

function getMonthlyData() {
  return withCache_('monthly_data', function() {
    var raw = getSheetDataAsObjects_("agg_monthly");
    
    var processed = raw.map(function(row) {
      var sales = parseCurrencyEs_(row.sales);
      var profit = parseCurrencyEs_(row.profit);
      var tx_count = safeNumber_(row.tx_count);
      
      return {
        month_start: safeString_(row.month_start),
        year_month: safeString_(row.year_month),
        sales: sales,
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: profit,
        tx_count: tx_count,
        avg_margin_pct: parsePercentEs_(row.avg_margin_pct),
        avg_profit_pct: parsePercentEs_(row.avg_profit_pct),
        avg_ticket: tx_count > 0 ? sales / tx_count : 0
      };
    });

    processed.sort(function(a, b) { return a.year_month.localeCompare(b.year_month); });

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

function getEntityData() {
  return withCache_('entity_data', function() {
    var raw = getSheetDataAsObjects_("agg_entity");
    
    var processed = raw.map(function(row) {
      return {
        entity_name: safeString_(row.entity_name),
        sales: parseCurrencyEs_(row.sales),
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: parseCurrencyEs_(row.profit),
        tx_count: safeNumber_(row.tx_count),
        avg_margin_pct: parsePercentEs_(row.avg_margin_pct),
        avg_profit_pct: parsePercentEs_(row.avg_profit_pct),
        last_sale_date: parseDateEs_(row.last_sale_date)
      };
    });

    return { table_data: processed };
  });
}

function getProductLocationData() {
  return withCache_('product_location_data', function() {
    var rawProducts = getSheetDataAsObjects_("agg_product");
    var rawLocations = getSheetDataAsObjects_("agg_location");

    var products = rawProducts.map(function(row) {
      return {
        product: safeString_(row.product),
        category: safeString_(row.category),
        sales: parseCurrencyEs_(row.sales),
        cost: parseCurrencyEs_(row.cost), // FIX: Añadido cost
        margin: parseCurrencyEs_(row.margin), // FIX: Añadido margin
        profit: parseCurrencyEs_(row.profit),
        avg_margin_pct: parsePercentEs_(row.avg_margin_pct),
        avg_profit_pct: parsePercentEs_(row.avg_profit_pct)
      };
    });

    var locations = rawLocations.map(function(row) {
      var sales = parseCurrencyEs_(row.sales);
      var target = parseCurrencyEs_(row.target_sales);
      return {
        location: safeString_(row.location),
        sales: sales,
        cost: parseCurrencyEs_(row.cost),
        expenses: parseCurrencyEs_(row.expenses),
        margin: parseCurrencyEs_(row.margin),
        profit: parseCurrencyEs_(row.profit),
        target_sales: target,
        target_attainment: target > 0 ? sales / target : 0
      };
    });

    var categoryMap = {};
    products.forEach(function(p) {
      if (!categoryMap[p.category]) categoryMap[p.category] = 0;
      categoryMap[p.category] += p.sales;
    });
    var category_distribution = Object.keys(categoryMap).map(function(k) {
      return { category: k, sales: categoryMap[k] };
    }).sort(function(a, b) { return b.sales - a.sales; });

    products.sort(function(a, b) { return b.sales - a.sales; });
    locations.sort(function(a, b) { return b.sales - a.sales; });
    var locationsByAttainment = locations.slice().sort(function(a, b) { return b.target_attainment - a.target_attainment; });

    var total_sales_loc = 0;
    var total_target_loc = 0;
    locations.forEach(function(l) {
      total_sales_loc += l.sales;
      total_target_loc += l.target_sales;
    });

    return {
      products: {
        list: products,
        category_distribution: category_distribution,
        top_product_by_sales: products.length > 0 ? products[0].product : "N/A",
        top_product_by_profit: products.slice().sort(function(a, b) { return b.profit - a.profit; })[0]?.product || "N/A",
        top_category_by_sales: category_distribution.length > 0 ? category_distribution[0].category : "N/A"
      },
      locations: {
        list: locations,
        top_location_by_sales: locations.length > 0 ? locations[0].location : "N/A",
        top_location_by_target_attainment: locationsByAttainment.length > 0 ? locationsByAttainment[0].location : "N/A",
        total_target_gap: total_target_loc - total_sales_loc
      }
    };
  });
}

function getAppBootstrapData() { return { overview: getOverviewData() }; }

function refreshAllDashboardData() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['overview_data', 'monthly_data', 'entity_data', 'product_location_data']);
  var res = { overview: getOverviewData(), monthly: getMonthlyData(), entity: getEntityData(), product_location: getProductLocationData() };
  if (res.overview.error || res.monthly.error || res.entity.error || res.product_location.error) {
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