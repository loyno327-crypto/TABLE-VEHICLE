/**
 * Бизнес-расчеты рейса.
 *
 * @param {Object} payload
 * @param {Object} settings
 * @return {Object}
 */
function calculateTripEconomy_(payload, settings) {
  const km = Number(payload.km) || 0;
  const emptyKm = Number(payload.emptyKm) || 0;
  const routes = payload.routes || [];

  const dieselPrice = Number(settings['Цена дизеля (руб/литр)']) || 0;
  const consumption = Number(settings['Расход (л/100 км)']) || 0;

  const leasingPct = (Number(settings['Процент лизинга']) || 0) / 100;
  const repairPct = (Number(settings['Процент ремонта']) || 0) / 100;
  const driverPct = (Number(settings['Процент водителю']) || 0) / 100;
  const taxPct = (Number(settings['Процент налога (для НДС)']) || 0) / 100;

  const emptySplitPct = (Number(settings['Деление холостого пробега']) || 50) / 100;

  const fuelPerKm = (consumption * dieselPrice) / 100;
  const fuel = km * fuelPerKm;
  const emptyFuel = emptyKm * fuelPerKm;

  let totalRevenue = 0;
  routes.forEach(function (r) {
    totalRevenue += Number(r.price) || 0;
  });

  const net = totalRevenue - fuel;
  const leasing = net * leasingPct;
  const repair = net * repairPct;
  const driverGross = net * driverPct;

  let tax = 0;
  routes.forEach(function (r) {
    if (r.paymentType === 'С НДС') {
      const part = Number(r.price) || 0;
      const ratio = totalRevenue > 0 ? part / totalRevenue : 0;
      tax += (driverGross * ratio) * taxPct;
    }
  });

  const baseCompany = net - leasing - repair - driverGross;
  const baseDriver = driverGross - tax;

  const companyEmptyAdjust = emptyFuel * emptySplitPct;
  const driverEmptyAdjust = emptyFuel * (1 - emptySplitPct);

  const companyTotal = baseCompany - companyEmptyAdjust;
  const driverTotal = baseDriver - driverEmptyAdjust;

  return {
    fuelPerKm: round2_(fuelPerKm),
    totalRevenue: round2_(totalRevenue),
    fuel: round2_(fuel),
    net: round2_(net),
    leasing: round2_(leasing),
    repair: round2_(repair),
    driver: round2_(driverGross),
    tax: round2_(tax),
    emptyFuel: round2_(emptyFuel),
    companyEmptyAdjust: round2_(companyEmptyAdjust),
    driverEmptyAdjust: round2_(driverEmptyAdjust),
    companyTotal: round2_(companyTotal),
    driverTotal: round2_(driverTotal)
  };
}

function validateTripPayload_(payload) {
  if (!payload) throw new Error('Пустой payload.');
  if (!payload.date) throw new Error('Укажите дату поездки.');
  if (!payload.mainRoute) throw new Error('Укажите основной маршрут.');
  if (!payload.routes || !payload.routes.length) throw new Error('Добавьте хотя бы один маршрут груза.');

  payload.routes.forEach(function (r, idx) {
    if (!r.route) throw new Error('Не заполнен маршрут в строке #' + (idx + 1));
    if (!(Number(r.price) > 0)) throw new Error('Цена должна быть > 0 в строке #' + (idx + 1));
    if (PAYMENT_TYPES.indexOf(r.paymentType) === -1) {
      throw new Error('Некорректный вид оплаты в строке #' + (idx + 1));
    }
  });
}

function round2_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}
