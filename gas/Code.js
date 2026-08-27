/**
 * 事務ツール工房 相談受付フォーム 受け口
 * POST(text/plain JSON) → スプレッドシート記録 + info@ へメール通知
 */
const PAGE_KEY = 'jimu2026';
const MAIL_TO = 'info@hayazai.com';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== PAGE_KEY) return json_({ ok: false, err: 'key' });

    const category = String(d.category || '').slice(0, 100);
    const company  = String(d.company  || '').slice(0, 100);
    const name     = String(d.name     || '').slice(0, 100);
    const tel      = String(d.tel      || '').slice(0, 50);
    const memo     = String(d.memo     || '').slice(0, 1000);
    if (!company && !tel) return json_({ ok: false, err: 'empty' });

    getSheet_().appendRow([new Date(), category, company, name, tel, memo]);

    MailApp.sendEmail({
      to: MAIL_TO,
      subject: '【要返信】事務ツール工房：新規相談 ' + company,
      body: [
        '特設サイトから新しい相談が届きました。',
        '',
        '困りごと: ' + category,
        '会社名: ' + company,
        'お名前: ' + name,
        '電話: ' + tel,
        'メモ: ' + memo,
        '',
        '受付一覧: ' + getSheetUrl_()
      ].join('\n')
    });

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, err: String(err) });
  }
}

function getSheet_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('SSID');
  if (!id) {
    const ss = SpreadsheetApp.create('事務ツール工房_相談受付');
    ss.getSheets()[0].appendRow(['受付日時', '困りごと', '会社名', 'お名前', '電話', 'メモ']);
    id = ss.getId();
    p.setProperty('SSID', id);
  }
  return SpreadsheetApp.openById(id).getSheets()[0];
}

function getSheetUrl_() {
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  return id ? 'https://docs.google.com/spreadsheets/d/' + id : '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 初回に1回だけエディタから実行して権限承認する */
function setup() {
  getSheet_();
  MailApp.sendEmail(MAIL_TO, '事務ツール工房フォーム 設定完了', '受付一覧: ' + getSheetUrl_());
}
