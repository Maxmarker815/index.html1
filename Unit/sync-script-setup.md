# Синхронизация расчётов Юнита между устройствами — настройка

> Этот файл — **памятка/резерв**. Сам код ниже работает не отсюда, а только
> после того, как ты вставишь его в `script.google.com` и развернёшь.
> Это ОТДЕЛЬНЫЙ скрипт («Unit Synk»), он НЕ связан со скриптом заказов.

## Версия 2 — с защитой от перезаписи устаревшими данными

Раньше любое устройство могло залить свою (старую) копию и «воскресить»
удалённое. Теперь сервер принимает запись, **только если она сделана поверх
актуальной версии** (`base` совпадает с тем, что лежит в облаке). Старые/
устаревшие устройства получают `stale` и не могут перезатереть свежие данные.

## Код скрипта (вставить ВЕСЬ, заменив старый)

```javascript
function getSyncSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('UNIT_SHEET_ID');
  var ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
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
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'busy' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var sheet = getSyncSheet_();
    var incoming;
    try { incoming = JSON.parse(e.postData.contents); } catch (err) { incoming = null; }
    if (!incoming || !incoming.batches) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var currentRaw = sheet.getRange('A1').getValue();
    var currentUpdatedAt = 0;
    if (currentRaw) {
      try { currentUpdatedAt = JSON.parse(currentRaw).updatedAt || 0; } catch (err) {}
    }
    // Защита: писать можно только поверх актуальной версии.
    if (currentUpdatedAt !== 0 && incoming.base !== currentUpdatedAt) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'stale', current: currentUpdatedAt }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var toStore = JSON.stringify({ batches: incoming.batches, updatedAt: incoming.updatedAt || (new Date()).getTime() });
    sheet.getRange('A1').setValue(toStore);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
```

## Как обновить (НЕ создавать новое развёртывание!)

1. `script.google.com` → проект **«Unit Synk»**.
2. Выдели весь код (`Ctrl+A`) → удали → вставь код выше → сохрани (`Ctrl+S`).
3. **«Начать развертывание»** → **«Управление развертываниями»**.
4. У активного развёртывания нажми **карандаш ✏️** → поле **«Версия»** →
   **«Новая версия»** → **«Развернуть»**.
5. URL останется прежним — менять ничего в Юните не нужно.

> Ограничение: ячейка A1 вмещает до 50 000 символов — для расчётов с запасом хватит.
