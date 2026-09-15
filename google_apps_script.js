/**
 * 安全衛生推進者管理システム - Google Apps Script (GAS) バックエンド
 * 【LINE公式・二重防御セキュリティ対応版】
 * 
 * 機能：
 * 1. 本社スプレッドシートへの記録データ自動蓄積（パトロール・リスク・ヒヤリハット・教育・健康）
 * 2. LINE友だちリスト自動管理（友だち追加時に自動登録）
 * 3. アプリ起動時のリアルタイム認証照会（許可 / 不許可判定）
 * 4. スプレッドシートのステータス変更連動（「不許可」でLINEリッチメニュー即時削除、「許可」で再リンク）
 */

// ════════════════════════════════════════════════════════════════
// 設定・プロパティ取得
// ════════════════════════════════════════════════════════════════
function getLineConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    channelAccessToken: props.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '',
    richMenuId: props.getProperty('LINE_RICH_MENU_ID') || ''
  };
}

/**
 * スプレッドシートを開いた時に専用メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ LINE連携設定')
    .addItem('🔑 LINEアクセストークン・メニューID設定', 'promptLineConfig')
    .addItem('📋 友だちリスト初期シート作成', 'setupFriendListSheet')
    .addToUi();
}

function promptLineConfig() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const tokenResp = ui.prompt('LINE Messaging API 設定', 'チャネルアクセストークン（長期）を入力してください：\n（未設定の場合は空欄でOK）', ui.ButtonSet.OK_CANCEL);
  if (tokenResp.getSelectedButton() === ui.Button.OK) {
    props.setProperty('LINE_CHANNEL_ACCESS_TOKEN', tokenResp.getResponseText().trim());
  }

  const menuResp = ui.prompt('LINE リッチメニュー設定', 'リッチメニューID（richmenu-...）を入力してください：\n（LINE画面からメニュー消去連動を行う場合に入力）', ui.ButtonSet.OK_CANCEL);
  if (menuResp.getSelectedButton() === ui.Button.OK) {
    props.setProperty('LINE_RICH_MENU_ID', menuResp.getResponseText().trim());
  }

  ui.alert('設定を保存しました。');
}

// ════════════════════════════════════════════════════════════════
// HTTP POST ハンドラ (データ同期 & Webhook)
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const raw = e.postData.contents;
    const payload = JSON.parse(raw);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. LINE Webhook イベント処理（友だち追加など）
    if (payload.events && Array.isArray(payload.events)) {
      handleLineWebhook(ss, payload.events);
      return jsonResponse({ status: 'ok' });
    }

    // 2. 接続テスト (Ping)
    if (payload.action === 'ping') {
      return jsonResponse({ status: 'success', message: 'Google Apps Script接続に成功しました！' });
    }

    // 3. アプリ起動時の認証チェック
    if (payload.action === 'check_auth') {
      const authResult = checkUserAuth(ss, payload.userId, payload.userName);
      return jsonResponse(authResult);
    }

    // 4. 単一レコード保存
    if (payload.action === 'save_record') {
      handleSaveRecord(ss, payload);
      return jsonResponse({ status: 'success', message: '記録が正常にスプレッドシートへ保存されました。' });
    }

    // 5. 一括同期 (Batch Sync)
    if (payload.action === 'batch_sync') {
      handleBatchSync(ss, payload);
      return jsonResponse({ status: 'success', message: '全データの一括同期が完了しました。' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + payload.action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════
// HTTP GET ハンドラ (認証照会 & 疎通確認)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter || {};
  const action = p.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // アプリ起動時の認証チェック (GETリクエスト対応)
  if (action === 'check_auth') {
    const userId = p.userId;
    const userName = p.userName || '';
    const authResult = checkUserAuth(ss, userId, userName);
    return jsonResponse(authResult);
  }

  // 全社ヒヤリハット一覧取得
  if (action === 'get_nearmisses') {
    const sheet = ss.getSheetByName('ヒヤリハット報告');
    if (!sheet) return jsonResponse({ status: 'success', data: [] });
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse({ status: 'success', data: [] });
    
    const headers = values[0];
    const rows = values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      return obj;
    });
    return jsonResponse({ status: 'success', data: rows });
  }

  return jsonResponse({
    status: 'success',
    message: '安全衛生推進者管理システム GAS Web API（LINE二重防御対応）稼働中'
  });
}

// ════════════════════════════════════════════════════════════════
// 認証・友だちリスト管理ロジック
// ════════════════════════════════════════════════════════════════

/**
 * ユーザーの利用権限をチェック（友だちリスト照会・未登録時は自動追加）
 */
function checkUserAuth(ss, userId, userName) {
  if (!userId) {
    return { status: 'error', allowed: false, reason: 'no_user_id' };
  }

  const sheet = getOrCreateFriendSheet(ss);
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // ヘッダー行除外して検索
  for (let i = 1; i < data.length; i++) {
    const rowUserId = String(data[i][1]).trim();
    if (rowUserId === String(userId).trim()) {
      const status = String(data[i][3]).trim();
      const isAllowed = (status === '許可');
      // 最終アクセス日時を更新
      sheet.getRange(i + 1, 6).setValue(nowStr);
      if (userName && !data[i][0]) sheet.getRange(i + 1, 1).setValue(userName);
      
      return {
        status: 'success',
        allowed: isAllowed,
        userStatus: status,
        name: data[i][2] || data[i][0] || '社員'
      };
    }
  }

  // 未登録の場合：新規追加（初期ステータス: 許可）
  sheet.appendRow([
    userName || 'LINEユーザー',
    userId,
    userName || '',
    '許可',
    nowStr,
    nowStr
  ]);

  return {
    status: 'success',
    allowed: true,
    userStatus: '許可',
    name: userName || '新規社員',
    isNew: true
  };
}

/**
 * LINE Webhook (follow イベント等) の処理
 */
function handleLineWebhook(ss, events) {
  const sheet = getOrCreateFriendSheet(ss);
  const config = getLineConfig();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  events.forEach(ev => {
    if (ev.type === 'follow' && ev.source && ev.source.userId) {
      const userId = ev.source.userId;
      let displayName = 'LINE友だち';

      // LINE プロファイルから表示名取得
      if (config.channelAccessToken) {
        try {
          const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
            headers: { 'Authorization': 'Bearer ' + config.channelAccessToken },
            muteHttpExceptions: true
          });
          if (res.getResponseCode() === 200) {
            const prof = JSON.parse(res.getContentText());
            displayName = prof.displayName || displayName;
          }
        } catch(e) {
          console.warn('Profile fetch failed', e);
        }
      }

      // シートに登録（存在しなければ）
      checkUserAuth(ss, userId, displayName);

      // リッチメニューが設定されていれば自動リンク
      if (config.channelAccessToken && config.richMenuId) {
        linkRichMenu(config.channelAccessToken, userId, config.richMenuId);
      }
    }
  });
}

/**
 * スプレッドシート編集検知トリガー（ステータスが「不許可」になったらメニュー削除）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== '友だちリスト') return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    // 4列目（ステータス列）が編集された場合
    if (col === 4 && row > 1) {
      const newStatus = String(e.range.getValue()).trim();
      const userId = String(sheet.getRange(row, 2).getValue()).trim();
      const config = getLineConfig();

      if (!userId || !config.channelAccessToken) return;

      if (newStatus === '不許可') {
        // LINEからリッチメニューを即座に削除
        unlinkRichMenu(config.channelAccessToken, userId);
      } else if (newStatus === '許可' && config.richMenuId) {
        // LINEにリッチメニューを再リンク
        linkRichMenu(config.channelAccessToken, userId, config.richMenuId);
      }
    }
  } catch(err) {
    console.error('onEdit error', err);
  }
}

/**
 * LINE Messaging API: リッチメニューのリンク（表示）
 */
function linkRichMenu(token, userId, richMenuId) {
  try {
    UrlFetchApp.fetch(`https://api.line.me/v2/bot/user/${userId}/richmenu/${richMenuId}`, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  } catch(e) {
    console.warn('linkRichMenu failed', e);
  }
}

/**
 * LINE Messaging API: リッチメニューのリンク解除（非表示・消去）
 */
function unlinkRichMenu(token, userId) {
  try {
    UrlFetchApp.fetch(`https://api.line.me/v2/bot/user/${userId}/richmenu`, {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  } catch(e) {
    console.warn('unlinkRichMenu failed', e);
  }
}

/**
 * 友だちリストシートの取得・自動生成
 */
function getOrCreateFriendSheet(ss) {
  let sheet = ss.getSheetByName('友だちリスト');
  if (!sheet) {
    sheet = ss.insertSheet('友だちリスト');
    const headers = ['LINE表示名', 'LINEユーザーID', '所属・氏名', 'ステータス', '登録日時', '最終アクセス日時'];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0e3d25'); // 濃いダークグリーン
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);

    // 列幅
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 150);

    // ステータス列（D列）に「許可 / 不許可」の入力規則（プルダウン）を設定
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['許可', '不許可'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('D2:D500').setDataValidation(rule);
  }
  return sheet;
}

function setupFriendListSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateFriendSheet(ss);
  SpreadsheetApp.getUi().alert('「友だちリスト」シートを作成しました。');
}

// ════════════════════════════════════════════════════════════════
// 業務記録データの保存処理（パトロール・リスク・ヒヤリハット・教育・健康）
// ════════════════════════════════════════════════════════════════
function handleSaveRecord(ss, payload) {
  const type = payload.type;
  const meta = payload.meta || {};
  const data = payload.data || {};
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const company = meta.company || '未設定事業場';
  const person = meta.person || '未設定推進者';

  if (type === 'patrol') {
    const sheet = getOrCreateRecordSheet(ss, '職場巡視・パトロール', [
      '受信日時', '事業場名', '安全衛生推進者', '実施日', '実施時刻', '点検エリア', '点検担当者', '合格件数', '問題件数', '要是正項目', '所見・サマリー'
    ]);
    const ngList = (data.items || []).filter(i => i.status === 'ng').map(i => i.label).join(' / ');
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.time || '', data.area || '', data.person || '',
      data.okCount || 0, data.ngCount || 0, ngList, data.summary || ''
    ]);
  } else if (type === 'risk') {
    const sheet = getOrCreateRecordSheet(ss, 'リスクアセスメント', [
      '受信日時', '事業場名', '安全衛生推進者', '登録日', '作業名・工程', '危険源・状態', '重篤度(1-4)', '可能性(1-4)', 'リスクスコア', 'リスク判定', '是正措置・対策', '担当者', '対策期限'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.task || '', data.hazard || '',
      data.sev || '', data.prob || '', data.score || '', data.levelLabel || '',
      data.measure || '', data.owner || '', data.deadline || ''
    ]);
  } else if (type === 'nearmiss') {
    const sheet = getOrCreateRecordSheet(ss, 'ヒヤリハット報告', [
      '受信日時', '事業場名', '安全衛生推進者', '発生日時', '発生場所', '報告者', '切迫度・重篤度', '状況・経緯', '推定原因', '再発防止対策'
    ]);
    const sevLabel = { 's1': '軽微（ひやりとした）', 's2': '中度（怪我の一歩手前）', 's3': '重大（重大災害の可能性）' }[data.severity] || data.severity || '';
    sheet.appendRow([
      nowStr, company, person, data.datetime || '', data.location || '', data.reporter || '',
      sevLabel, data.situation || '', data.cause || '', data.measure || ''
    ]);
  } else if (type === 'education') {
    const sheet = getOrCreateRecordSheet(ss, '安全衛生教育', [
      '受信日時', '事業場名', '安全衛生推進者', '実施年月日', '教育区分', 'テーマ・研修名', '対象者・参加人数', '講師・担当者', '教育内容・所見'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.category || '', data.title || '',
      data.participants || '', data.instructor || '', data.content || ''
    ]);
  } else if (type === 'health') {
    const sheet = getOrCreateRecordSheet(ss, '健康管理記録', [
      '受信日時', '事業場名', '安全衛生推進者', '実施日', '区分', '対象者・件数', '判定・結果概要', '事後措置・配慮事項', '備考'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.category || '', data.target || '',
      data.result || '', data.action || '', data.note || ''
    ]);
  }
}

function handleBatchSync(ss, payload) {
  const meta = payload.meta || {};
  const state = payload.state || {};

  if (Array.isArray(state.patrols)) state.patrols.forEach(item => handleSaveRecord(ss, { type: 'patrol', meta: meta, data: item }));
  if (Array.isArray(state.risks)) state.risks.forEach(item => handleSaveRecord(ss, { type: 'risk', meta: meta, data: item }));
  if (Array.isArray(state.nearmisses)) state.nearmisses.forEach(item => handleSaveRecord(ss, { type: 'nearmiss', meta: meta, data: item }));
  if (Array.isArray(state.educations)) state.educations.forEach(item => handleSaveRecord(ss, { type: 'education', meta: meta, data: item }));
  if (Array.isArray(state.healthRecords)) state.healthRecords.forEach(item => handleSaveRecord(ss, { type: 'health', meta: meta, data: item }));
}

function getOrCreateRecordSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a5c38');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    headers.forEach((_, i) => sheet.setColumnWidth(i + 1, 140));
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
