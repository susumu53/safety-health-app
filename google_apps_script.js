/**
 * 安全衛生推進者管理システム - Google Apps Script (GAS) バックエンド
 * 
 * 【使い方】
 * 1. Google ドライブで「新規」>「Google スプレッドシート」を作成します（ファイル名は任意：例「全社安全衛生推進統合データベース」）。
 * 2. メニューバーの「拡張機能」>「Apps Script」をクリックします。
 * 3. 既存のコードを全て消去し、このファイルの内容を丸ごと貼り付けて保存（フロッピーアイコン）します。
 * 4. 画面右上の「デプロイ」>「新しいデプロイ」をクリックします。
 * 5. 歯車アイコンをクリックし「ウェブアプリ」を選択します。
 * 6. 設定：
 *    - 次のユーザーとして実行: 「自分」
 *    - アクセスできるユーザー: 「全員」（※社外・現場端末からのデータ送信を受け付けるため）
 * 7. 「デプロイ」ボタンを押し、アクセスを承認します。
 * 8. 発行された「ウェブアプリの URL」（https://script.google.com/macros/s/.../exec）をコピーして、
 *    安全衛生推進者アプリの「基本設定」>「本社スプレッドシート連携URL」に貼り付けます。
 */

function doPost(e) {
  try {
    const raw = e.postData.contents;
    const payload = JSON.parse(raw);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 接続テスト (Ping)
    if (payload.action === 'ping') {
      return jsonResponse({ status: 'success', message: 'Google Apps Script接続に成功しました！' });
    }

    // 単一レコード保存
    if (payload.action === 'save_record') {
      handleSaveRecord(ss, payload);
      return jsonResponse({ status: 'success', message: '記録が正常にスプレッドシートへ保存されました。' });
    }

    // 一括同期 (Batch Sync)
    if (payload.action === 'batch_sync') {
      handleBatchSync(ss, payload);
      return jsonResponse({ status: 'success', message: '全データの一括同期が完了しました。' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + payload.action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  // ブラウザで直接開いた場合や疎通テスト用
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'get_nearmisses') {
    // 全社ヒヤリハット一覧の取得
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
    message: '安全衛生推進者管理システム GAS Web API は正常に稼働しています。'
  });
}

/**
 * 単一レコードの振り分け保存
 */
function handleSaveRecord(ss, payload) {
  const type = payload.type;
  const meta = payload.meta || {};
  const data = payload.data || {};
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const company = meta.company || '未設定事業場';
  const person = meta.person || '未設定推進者';

  if (type === 'patrol') {
    const sheet = getOrCreateSheet(ss, '職場巡視・パトロール', [
      '受信日時', '事業場名', '安全衛生推進者', '実施日', '実施時刻', '点検エリア', '点検担当者', '合格件数', '問題件数', '要是正項目', '所見・サマリー'
    ]);
    const ngList = (data.items || []).filter(i => i.status === 'ng').map(i => i.label).join(' / ');
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.time || '', data.area || '', data.person || '',
      data.okCount || 0, data.ngCount || 0, ngList, data.summary || ''
    ]);
  } else if (type === 'risk') {
    const sheet = getOrCreateSheet(ss, 'リスクアセスメント', [
      '受信日時', '事業場名', '安全衛生推進者', '登録日', '作業名・工程', '危険源・状態', '重篤度(1-4)', '可能性(1-4)', 'リスクスコア', 'リスク判定', '是正措置・対策', '担当者', '対策期限'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.task || '', data.hazard || '',
      data.sev || '', data.prob || '', data.score || '', data.levelLabel || '',
      data.measure || '', data.owner || '', data.deadline || ''
    ]);
  } else if (type === 'nearmiss') {
    const sheet = getOrCreateSheet(ss, 'ヒヤリハット報告', [
      '受信日時', '事業場名', '安全衛生推進者', '発生日時', '発生場所', '報告者', '切迫度・重篤度', '状況・経緯', '推定原因', '再発防止対策'
    ]);
    const sevLabel = { 's1': '軽微（ひやりとした）', 's2': '中度（怪我の一歩手前）', 's3': '重大（重大災害の可能性）' }[data.severity] || data.severity || '';
    sheet.appendRow([
      nowStr, company, person, data.datetime || '', data.location || '', data.reporter || '',
      sevLabel, data.situation || '', data.cause || '', data.measure || ''
    ]);
  } else if (type === 'education') {
    const sheet = getOrCreateSheet(ss, '安全衛生教育', [
      '受信日時', '事業場名', '安全衛生推進者', '実施年月日', '教育区分', 'テーマ・研修名', '対象者・参加人数', '講師・担当者', '教育内容・所見'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.category || '', data.title || '',
      data.participants || '', data.instructor || '', data.content || ''
    ]);
  } else if (type === 'health') {
    const sheet = getOrCreateSheet(ss, '健康管理記録', [
      '受信日時', '事業場名', '安全衛生推進者', '実施日', '区分', '対象者・件数', '判定・結果概要', '事後措置・配慮事項', '備考'
    ]);
    sheet.appendRow([
      nowStr, company, person, data.date || '', data.category || '', data.target || '',
      data.result || '', data.action || '', data.note || ''
    ]);
  }
}

/**
 * 一括同期処理
 */
function handleBatchSync(ss, payload) {
  const meta = payload.meta || {};
  const state = payload.state || {};

  if (Array.isArray(state.patrols)) {
    state.patrols.forEach(item => handleSaveRecord(ss, { type: 'patrol', meta: meta, data: item }));
  }
  if (Array.isArray(state.risks)) {
    state.risks.forEach(item => handleSaveRecord(ss, { type: 'risk', meta: meta, data: item }));
  }
  if (Array.isArray(state.nearmisses)) {
    state.nearmisses.forEach(item => handleSaveRecord(ss, { type: 'nearmiss', meta: meta, data: item }));
  }
  if (Array.isArray(state.educations)) {
    state.educations.forEach(item => handleSaveRecord(ss, { type: 'education', meta: meta, data: item }));
  }
  if (Array.isArray(state.healthRecords)) {
    state.healthRecords.forEach(item => handleSaveRecord(ss, { type: 'health', meta: meta, data: item }));
  }
}

/**
 * 指定名のシートを取得（存在しなければ作成しヘッダー行を設定＆スタイル適用）
 */
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // ヘッダー追加
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a5c38'); // アプリ基調色グリーン
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 列幅自動調整
    headers.forEach((_, i) => sheet.setColumnWidth(i + 1, 140));
  }
  return sheet;
}

/**
 * JSON レスポンス生成ヘルパー
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
