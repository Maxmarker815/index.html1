# Синхронизация расчётов Юнита между устройствами — настройка

> Этот файл — **памятка/резерв**. Сам код ниже работает не отсюда, а только
> после того, как ты вставишь его в `script.google.com` и развернёшь.
> Это ОТДЕЛЬНЫЙ скрипт, он НЕ связан со скриптом заказов и не может его сломать.

## Шаг 1. Создать скрипт

1. Открой <https://script.google.com>
2. Нажми **«Новый проект»** (New project).
3. Удали весь код по умолчанию и вставь код из блока ниже.
4. Нажми **сохранить** (значок дискеты), назови проект, например `Unit Sync`.

## Шаг 2. Код скрипта

```javascript
/**
 * Unit — синхронизация расчётов между устройствами.
 * Отдельный веб-апп. Хранит JSON расчётов в ячейке A1 собственной таблицы.
 */
function getSyncSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('UNIT_SHEET_ID');
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Unit — данные расчётов');
    props.setProperty('UNIT_SHEET_ID', ss.getId());
  }
  return ss.getSheets()[0];
}

function doGet(e) {
  var value = getSyncSheet_().getRange('A1').getValue();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: value || '' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = (e && e.postData) ? e.postData.contents : '';
  getSyncSheet_().getRange('A1').setValue(body);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## Шаг 3. Развернуть как веб-приложение

1. Справа сверху: **«Развернуть»** → **«Новое развёртывание»** (Deploy → New deployment).
2. Тип (шестерёнка) → **«Веб-приложение»** (Web app).
3. Настройки:
   - **Запуск от имени:** Я (Me)
   - **Доступ:** Все (Anyone)
4. **«Развернуть»** → пройди авторизацию (выбери свой аккаунт →
   «Дополнительно» → «Перейти к проекту» → «Разрешить»).
5. Скопируй **URL веб-приложения** (заканчивается на `/exec`).
6. Пришли этот URL в чат — я впишу его в Юнит, и синхронизация заработает.

## Как это работает

- Юнит при открытии загружает расчёты из этой таблицы (облако).
- При изменении — сохраняет обратно (с задержкой ~1 сек).
- Нет интернета → Юнит работает локально, как раньше.
- Принцип «кто последним сохранил — тот и прав»: не редактируй на двух
  устройствах одновременно.

> Ограничение: ячейка A1 вмещает до 50 000 символов — для расчётов закупа
> этого с огромным запасом хватит.
