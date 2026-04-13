/**
 * Точка входа: создаёт меню и базовую структуру листов.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Заявки')
    .addItem('Список заявок', 'showTripsDialog')
    .addItem('Внести поездку', 'showTripFormDialog')
    .addItem('Дашборд', 'showDashboardDialog')
    .addItem('Настройки', 'showSettingsDialog')
    .addToUi();

  ensureSchema_();
}

/**
 * Инициализация таблиц при установке скрипта.
 */
function onInstall() {
  onOpen();
}

/** Показывает общее окно со списком заявок. */
function showTripsDialog() {
  const template = HtmlService.createTemplateFromFile('MainApp');
  template.initialView = 'list';
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(760)
    .setTitle('Список заявок');

  SpreadsheetApp.getUi().showModalDialog(html, 'Список заявок');
}

/** Показывает окно формы создания/редактирования поездки. */
function showTripFormDialog(tripId) {
  const template = HtmlService.createTemplateFromFile('MainApp');
  template.initialView = 'form';
  template.editTripId = tripId || '';
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(760)
    .setTitle(tripId ? ('Редактирование #' + tripId) : 'Внести поездку');

  SpreadsheetApp.getUi().showModalDialog(html, tripId ? ('Редактирование #' + tripId) : 'Внести поездку');
}

/** Показывает окно дашборда. */
function showDashboardDialog() {
  const template = HtmlService.createTemplateFromFile('MainApp');
  template.initialView = 'dashboard';
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(760)
    .setTitle('Дашборд');

  SpreadsheetApp.getUi().showModalDialog(html, 'Дашборд');
}

/** Показывает окно настроек. */
function showSettingsDialog() {
  const template = HtmlService.createTemplateFromFile('MainApp');
  template.initialView = 'settings';
  const html = template.evaluate()
    .setWidth(900)
    .setHeight(680)
    .setTitle('Настройки');

  SpreadsheetApp.getUi().showModalDialog(html, 'Настройки');
}

/** Отдельное окно карточки поездки. */
function showTripCardDialog(tripId) {
  const template = HtmlService.createTemplateFromFile('TripCard');
  template.tripId = String(tripId);
  const html = template.evaluate()
    .setWidth(900)
    .setHeight(720)
    .setTitle('Карточка поездки #' + tripId);

  SpreadsheetApp.getUi().showModalDialog(html, 'Карточка поездки #' + tripId);
}

/**
 * Обёртка для вызова карточки из HtmlService.
 * @param {string|number} tripId
 */
function openTripCardFromClient(tripId) {
  showTripCardDialog(tripId);
}

/**
 * Подключение html-партиалов.
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
