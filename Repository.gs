/**
 * Создаёт/обновляет структуру всех листов.
 */
function ensureSchema_() {
  const ss = SpreadsheetApp.getActive();

  const tripsSheet = getOrCreateSheet_(ss, SHEETS.TRIPS);
  ensureHeader_(tripsSheet, TRIPS_HEADERS);

  const routesSheet = getOrCreateSheet_(ss, SHEETS.ROUTES);
  ensureHeader_(routesSheet, ROUTES_HEADERS);

  const settingsSheet = getOrCreateSheet_(ss, SHEETS.SETTINGS);
  ensureHeader_(settingsSheet, ['Параметр', 'Значение']);
  ensureSettingsRows_(settingsSheet);

  const dictSheet = getOrCreateSheet_(ss, SHEETS.DICTS);
  ensureHeader_(dictSheet, ['Справочник', 'Значение']);
  ensurePaymentTypesRows_(dictSheet);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeader_(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const current = range.getValues()[0];
  const mismatch = headers.some(function (h, i) { return current[i] !== h; });
  if (mismatch) {
    range.setValues([headers]).setFontWeight('bold').setBackground('#eceff1');
  }
}

function ensureSettingsRows_(sheet) {
  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  const keysMap = {};
  existing.forEach(function (r) { if (r[0]) keysMap[r[0]] = r[1]; });

  const rows = SETTINGS_KEYS.map(function (key) {
    const value = Object.prototype.hasOwnProperty.call(keysMap, key) ? keysMap[key] : DEFAULT_SETTINGS[key];
    return [key, value];
  });
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function ensurePaymentTypesRows_(sheet) {
  const all = sheet.getDataRange().getValues();
  const existing = all.slice(1).filter(function (r) { return r[0] === 'Виды оплаты'; }).map(function (r) { return r[1]; });
  if (existing.length === PAYMENT_TYPES.length && PAYMENT_TYPES.every(function (x) { return existing.indexOf(x) !== -1; })) {
    return;
  }

  const withoutPayment = all.slice(1).filter(function (r) { return r[0] !== 'Виды оплаты'; });
  const rows = withoutPayment.concat(PAYMENT_TYPES.map(function (x) { return ['Виды оплаты', x]; }));
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function mapRows_(headers, rows) {
  return rows.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function getSettingsMap_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  const result = {};
  values.forEach(function (row) {
    if (row[0]) {
      result[row[0]] = Number(row[1]) || 0;
    }
  });
  return result;
}

function getPaymentTypes() {
  ensureSchema_();
  return PAYMENT_TYPES.slice();
}

function getTripsIndex_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRIPS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || TRIPS_HEADERS;
  const rows = values.slice(1).filter(function (r) { return r[0] !== ''; });
  return mapRows_(headers, rows);
}

function getRoutesIndex_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.ROUTES);
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || ROUTES_HEADERS;
  const rows = values.slice(1).filter(function (r) { return r[0] !== ''; });
  return mapRows_(headers, rows);
}

function nextTripId_() {
  const trips = getTripsIndex_();
  const maxId = trips.reduce(function (acc, row) {
    return Math.max(acc, Number(row.ID) || 0);
  }, 0);
  return maxId + 1;
}

function upsertTrip(payload) {
  ensureSchema_();
  validateTripPayload_(payload);

  const settings = getSettingsMap_();
  const calc = calculateTripEconomy_(payload, settings);

  const ss = SpreadsheetApp.getActive();
  const tripsSheet = ss.getSheetByName(SHEETS.TRIPS);
  const routesSheet = ss.getSheetByName(SHEETS.ROUTES);

  let tripId = Number(payload.id) || 0;
  if (!tripId) {
    tripId = nextTripId_();
  }

  const tripRow = [
    tripId,
    payload.date,
    payload.mainRoute,
    Number(payload.km) || 0,
    Number(payload.emptyKm) || 0,
    calc.totalRevenue,
    calc.fuel,
    calc.net,
    calc.leasing,
    calc.repair,
    calc.driver,
    calc.tax,
    calc.companyTotal,
    calc.driverTotal,
    payload.comment || ''
  ];

  const existingRows = tripsSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < existingRows.length; i += 1) {
    if (Number(existingRows[i][0]) === tripId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    tripsSheet.appendRow(tripRow);
  } else {
    tripsSheet.getRange(targetRow, 1, 1, tripRow.length).setValues([tripRow]);
  }

  deleteRoutesByTripId_(tripId);
  if (payload.routes && payload.routes.length) {
    const rows = payload.routes.map(function (route) {
      return [tripId, route.route, Number(route.price) || 0, route.paymentType];
    });
    routesSheet.getRange(routesSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }

  return {
    success: true,
    tripId: tripId,
    economics: calc
  };
}

function deleteRoutesByTripId_(tripId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.ROUTES);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i -= 1) {
    if (Number(data[i][0]) === Number(tripId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function getTripById(tripId) {
  const trips = getTripsIndex_();
  const routes = getRoutesIndex_();
  const t = trips.find(function (row) { return Number(row.ID) === Number(tripId); });
  if (!t) return null;

  const tripRoutes = routes
    .filter(function (row) { return Number(row['ID поездки']) === Number(tripId); })
    .map(function (r) {
      return {
        route: r['Маршрут'],
        price: Number(r['Цена']) || 0,
        paymentType: r['Вид оплаты']
      };
    });

  return {
    id: Number(t.ID),
    date: t['Дата'],
    mainRoute: t['Основной маршрут'],
    km: Number(t['Общий километраж']) || 0,
    emptyKm: Number(t['Холостой пробег (км)']) || 0,
    comment: t['Комментарий'] || '',
    totals: {
      totalRevenue: Number(t['Общая выручка']) || 0,
      fuel: Number(t['Топливо']) || 0,
      net: Number(t['Чистый остаток']) || 0,
      leasing: Number(t['Лизинг']) || 0,
      repair: Number(t['Ремонт']) || 0,
      driver: Number(t['Водитель']) || 0,
      tax: Number(t['Налог']) || 0,
      companyTotal: Number(t['Итог компании']) || 0,
      driverTotal: Number(t['Итог водителя']) || 0
    },
    routes: tripRoutes
  };
}

function listTrips(filters) {
  const routes = getRoutesIndex_();
  const trips = getTripsIndex_();

  const routesByTrip = {};
  routes.forEach(function (r) {
    const id = Number(r['ID поездки']);
    routesByTrip[id] = routesByTrip[id] || [];
    routesByTrip[id].push({
      route: r['Маршрут'],
      price: Number(r['Цена']) || 0,
      paymentType: r['Вид оплаты']
    });
  });

  const from = filters && filters.from ? new Date(filters.from) : null;
  const to = filters && filters.to ? new Date(filters.to) : null;
  const q = filters && filters.query ? String(filters.query).toLowerCase() : '';
  const paymentType = filters && filters.paymentType ? String(filters.paymentType) : '';

  return trips
    .map(function (t) {
      const id = Number(t.ID);
      const tripRoutes = routesByTrip[id] || [];
      return {
        id: id,
        date: t['Дата'],
        mainRoute: t['Основной маршрут'],
        km: Number(t['Общий километраж']) || 0,
        amount: Number(t['Общая выручка']) || 0,
        companyTotal: Number(t['Итог компании']) || 0,
        driverTotal: Number(t['Итог водителя']) || 0,
        routes: tripRoutes
      };
    })
    .filter(function (trip) {
      const dt = new Date(trip.date);
      if (from && dt < from) return false;
      if (to && dt > to) return false;
      if (q) {
        const hay = (trip.mainRoute + ' ' + trip.routes.map(function (r) { return r.route; }).join(' ')).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      if (paymentType) {
        const hasType = trip.routes.some(function (r) { return r.paymentType === paymentType; });
        if (!hasType) return false;
      }
      return true;
    })
    .sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
}

function deleteTrip(tripId) {
  const ss = SpreadsheetApp.getActive();
  const tripsSheet = ss.getSheetByName(SHEETS.TRIPS);
  const rows = tripsSheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (Number(rows[i][0]) === Number(tripId)) {
      tripsSheet.deleteRow(i + 1);
      break;
    }
  }
  deleteRoutesByTripId_(tripId);

  return { success: true };
}

function getSettings() {
  return getSettingsMap_();
}

function saveSettings(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  SETTINGS_KEYS.forEach(function (key, i) {
    const value = Number(payload[key]);
    sheet.getRange(i + 2, 2).setValue(isNaN(value) ? DEFAULT_SETTINGS[key] : value);
  });
  return { success: true };
}

function getDashboardData(filters) {
  const trips = listTrips(filters || {});
  const result = {
    tripsCount: trips.length,
    totalRevenue: 0,
    companyProfit: 0,
    driverPayout: 0,
    fuelCosts: 0,
    leasingCosts: 0,
    repairCosts: 0,
    taxTotal: 0
  };

  trips.forEach(function (trip) {
    const full = getTripById(trip.id);
    result.totalRevenue += full.totals.totalRevenue;
    result.companyProfit += full.totals.companyTotal;
    result.driverPayout += full.totals.driverTotal;
    result.fuelCosts += full.totals.fuel;
    result.leasingCosts += full.totals.leasing;
    result.repairCosts += full.totals.repair;
    result.taxTotal += full.totals.tax;
  });

  Object.keys(result).forEach(function (k) {
    if (typeof result[k] === 'number') {
      result[k] = round2_(result[k]);
    }
  });

  return result;
}
