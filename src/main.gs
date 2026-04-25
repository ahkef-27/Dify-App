/**
 * 🚀 メイン処理：自動リサーチ・分析フロー（プロ仕様・重複回避＆件数確保版）
 * [概要] 設定シートのキーワードに基づき、新規データが10件たまるまで検索を継続。
 * 常にボリュームのある分析結果をDifyへ送信します。
 */

function main() { 
  try { 
    // 1. 設定の読み込み
    const ss = SpreadsheetApp.getActive();
    const configSheet = ss.getSheetByName("設定");
    const keyword = configSheet.getRange("B1").getValue();

    if (!keyword || String(keyword).trim() === "") {
      sendSlack("⚠️ 【リサーチ中断】設定シートのキーワードが空欄です。");
      writeLog("main", "info", "キーワード未設定のためスキップ");
      return; 
    }

    // --- 検索・重複排除ループ開始 ---
    let allNewData = [];
    let startIndex = 1;
    const maxAttempts = 3;  // 最大3ページ（30件分）まで調査
    const targetCount = 10; // 確保したい新規データの件数

    for (let i = 0; i < maxAttempts; i++) {
      console.log(`${startIndex}件目から検索を実行中...`);
      writeLog("main", "info", `${startIndex}件目からの検索を開始`);

      // Google検索実行
      const googleData = fetchGoogleResults(keyword, startIndex);
      if (googleData.length === 0) break; 

      // 重複チェック（既存のURLを排除）
      const newData = filterNewResults(googleData);
      
      // 新規分をストックに追加
      allNewData = allNewData.concat(newData);

      console.log(`現在の新規取得累計: ${allNewData.length}件`);

      // 目標件数に達したらループ終了
      if (allNewData.length >= targetCount) {
        allNewData = allNewData.slice(0, targetCount); // ちょうど10件にする
        break;
      }

      // 次の10件へ（ページめくり）
      startIndex += 10;
    }
    // --- ループ終了 ---

    if (allNewData.length === 0) {
      writeLog("main", "info", "新規データが見つかりませんでした。");
      console.log("新規データがないため終了。");
      return;
    }

    // 2. 新規データのみを保存
    saveRawData(allNewData); 

    // 3. Dify用入力テキスト作成
    let inputText = buildDifyInputFromData(allNewData); 

    // 英字判定（翻訳モード）
    const isEnglishContent = /[a-zA-Z]{20,}/.test(inputText); 
    if (isEnglishContent) {
      inputText = "【重要：分析結果はすべて日本語で出力してください】\n\n" + inputText;
    }

    // 4. Dify AI分析実行
    const summary = callDifyAI(inputText); 
    
    // 5. 分析結果を保存
    saveAIAnalysis(summary); 

    // 6. 完了通知
    formatSheets();
    const successMsg = `✅ *【リサーチ完了】*\n` +
                       `キーワード: *${keyword}*\n` +
                       `新規取得数: *${allNewData.length}件*\n` +
                       `Difyによる分析が完了しました。ダッシュボードを確認してください。`;
    sendSlack(successMsg);
    writeLog("main", "success", `${allNewData.length}件の処理を完遂`); 

  } catch (e) { 
    const errorAlert = `🚨 *【システムエラー発生！】*\n内容: \`${e.message}\`\n場所: \`${e.stack}\``;
    sendSlack(errorAlert);
    writeLog("main", "error", e.toString());
  } 
}

// 重複URLを弾くためのフィルター関数
function filterNewResults(googleData) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("生データ");
  const lastRow = sheet.getLastRow();
  
  // シートが空（ヘッダーのみ）の場合は、すべての検索結果を新規として扱う
  if (lastRow < 2) return googleData; 

  /* --- 旧ロジック（URL列の指定ミスと空白判定を修正するためコメントアウト） ---
  const existingUrls = lastRow < 2 ? [] : sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  return googleData.filter(item => !existingUrls.includes(item.url));
  -------------------------------------------------------------------------- */

  // 新ロジック：URL列(4列目)を正確に取得し、空白を除去して判定
  const existingUrls = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat(); 
  
  return googleData.filter(item => {
    return !existingUrls.some(existingUrl => 
      String(existingUrl).trim() === String(item.url).trim()
    );
  });
}

// 新しいデータ配列からDify用の入力テキストを作る関数
function buildDifyInputFromData(dataList) {
  return dataList
    .map(row => `タイトル: ${row.title}\n本文: ${row.body}`)
    .join("\n\n");
}

/**
 * Google Custom Search APIを使用して検索結果を取得する関数（startIndex対応版）
 * @param {string} keyword 検索キーワード
 * @param {number} startIndex 取得開始位置（1, 11, 21...）
 * @returns {Array} 取得した検索結果の配列
 */
function fetchGoogleResults(keyword, startIndex = 1) { 
  // キーワードが空なら空の配列を返して安全に終わる
  if (!keyword || String(keyword).trim() === "") {
    return [];
  }
  
  const ss = SpreadsheetApp.getActive(); 
  const sheet = ss.getSheetByName("設定"); 

  // 設定シートの列幅を調整（UIの視認性向上のため）
  if (sheet) {
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    SpreadsheetApp.flush(); 
  }

  // スクリプトプロパティからAPIキーと検索エンジンIDを取得
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GOOGLE_API_KEY');
  const cx = props.getProperty('GOOGLE_CX');

  // URLに &start=${startIndex} を追加してページめくりを可能にする
  const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cx}&q=${encodeURIComponent(keyword)}&start=${startIndex}`;

  // APIリクエスト実行
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); 
  const json = JSON.parse(response.getContentText()); 
  const items = json.items || [];

  // 取得した直後に「エラーURL」を物理的に消す処理
  const cleanItems = items.filter(item => item.link && !item.link.includes("google.com/sorry"));

  // 必要な情報を整理して返す
  return cleanItems.map(item => ({ 
    source: "google", 
    title: item.title, 
    url: item.link, 
    body: item.snippet, 
    date: new Date() 
  })); 
}

/**
 * DifyのWorkflow APIを呼び出してテキスト解析を行う関数
 * @param {string} inputText 解析対象のテキスト
 * @returns {string} AIからの回答テキスト
 */

function callDifyAI(inputText) {
  const url = "https://api.dify.ai/v1/workflows/run";

  const apiKey = PropertiesService.getScriptProperties().getProperty('DIFY_API_KEY');

  const payload = {
    inputs: { text: inputText },
    user: "gas-user",
    response_mode: "blocking" // 同期処理（結果が出るまで待機）
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // エラー時もレスポンスを取得
  };

  const response = UrlFetchApp.fetch(url, options);
  const resText = response.getContentText();
  const json = JSON.parse(resText);

  // Difyのレスポンス構造に合わせてデータを抽出
  if (json.data && json.data.outputs) {
  return json.data.outputs.text || json.data.outputs.result || JSON.stringify(json.data.outputs);
  }
  
  return "解析エラー: " + resText;
}

/**
 * 取得した生の検索データを「生データ」シートに保存する関数
 * @param {Array} data 保存するデータの配列
 */

function saveRawData(data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("生データ");
  const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

  data.forEach(row => {
    sheet.appendRow([now,row.source,row.title,row.url,row.body]);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, 5);
    range.setHorizontalAlignment("left");
    range.setVerticalAlignment("middle");
    range.setWrap(true);

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 600);
  }

  sheet.getRange(2, 4, lastRow - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, 5, lastRow - 1, 1).setWrap(true);
  
  SpreadsheetApp.flush(); // ← これを入れると、即座に画面が変わります！
}

/**
 * Difyから返ってきたJSON形式の解析結果を「AI分析」シートに保存する関数
 * @param {string} summaryText Difyからのレスポンス（JSON文字列を想定）
 */

function saveAIAnalysis(summaryText) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("AI分析");
  
  try {
    // Markdownのコードブロック等を除去して純粋なJSONにする
    let cleanJson = summaryText.replace(/```json|```/g, "").trim();
    let rawData = JSON.parse(cleanJson);
    let finalData;

    // ネストされた構造（rawData.textの中身がJSON文字列など）に対応
    if (typeof rawData === 'object' && rawData !== null) {
      // もし既にパース済みのオブジェクトの中に、文字列のJSONが隠れていたら再度パース
      let nestedText = rawData.text || rawData.summary || rawData.result || "";
      if (typeof nestedText === "string" && nestedText.includes("{")) {
        finalData = JSON.parse(nestedText);
      } else {
        finalData = rawData;
      }
    }

    const now = new Date();
    
    // シートの各列（A〜G）に対応するデータを配列化
    const rowData = [
      now, // A: 日時
      finalData.sentiment || "不明", // B: ポジネガ
      finalData.emotion_score || 0, // C: スコア
      Array.isArray(finalData.trend_words) ? finalData.trend_words.join(", ") : (finalData.trend_words || ""), // D: トレンド
      `強み: ${finalData.competitor_analysis?.strengths || ""}\n弱み: ${finalData.competitor_analysis?.weaknesses || ""}`, // E: 競合
      finalData.summary || "", // F: サマリー
      Array.isArray(finalData.suggestions) ? finalData.suggestions.join("\n") : (finalData.suggestions || "") // G: 改善提案
    ];

    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, 7);
    range.setVerticalAlignment("middle").setWrap(true);
    const widths = [150, 100, 80, 200, 300, 450, 450];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    console.log("成功！各列にデータが整列しました。");

  } catch (e) {
    console.error("解析失敗:", e);
    sheet.appendRow([new Date(), "ERROR", 0, "", "", "解析失敗: " + e.message, summaryText]);

    // ★ここでもSlack通知（AIの回答形式が不正だった場合など）
    sendSlack("⚠️ *【AI解析エラー】*\n分析結果の保存に失敗しました。\n内容: " + e.message);
  }
  // 列幅の最終調整
  const widths = [130, 100, 80, 170, 300, 300, 300];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  SpreadsheetApp.flush(); // ← これを入れると、即座に画面が変わります！
}

function formatSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("AI分析");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    range.setVerticalAlignment("middle"); // 上下真ん中
    range.setHorizontalAlignment("center"); // 左右真ん中
    range.setWrap(true); // 文字が長い時に折り返す（これがないと横に突き抜けます）
  }
}

/**
 * 実行状況を「ログ」シートに記録する関数
 */

function writeLog(func, status, message) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("ログ");
  // 現在時刻を「2026/03/04 15:30:45」のような日本時間形式にする
  const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 250);

  sheet.appendRow([
    now,
    func,
    status,
    message
  ]);
}

/**
 * GASをWebアプリとして公開した際に、画面（HTML）を表示するための関数
 */
function doGet() {
  // index.html というファイルを表示
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * Webダッシュボード（フロントエンド）から呼び出され、最新の分析結果を返す関数
 * @returns {Object} 最新の分析結果データ
 */
function getLatestAnalysis() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("AI分析");
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return { summary: "データがまだありません。" };

  // 1. まずは全データを取得
  const rowData = sheet.getRange(lastRow, 1, 1, 7).getValues()[0];
  
  // 2. 🚨【ここが重要】日付を「ただの文字列」として、GAS側でガチガチに固定して作る
  // Utilities.formatDate を使い、ブラウザに「計算」させる隙を与えません。
  const fixedDateStr = Utilities.formatDate(new Date(rowData[0]), "JST", "yyyy/MM/dd HH:mm:ss");
  
  return {
    date: fixedDateStr, // 「2026/04/16 18:49:40」という「文字」として送る
    sentiment: rowData[1],
    score: parseFloat(rowData[2]) || 0,
    trend_words: rowData[3],
    summary: rowData[5]
  };
}

/**
 * SlackのIncoming Webhookを使用してメッセージを通知する関数
 * @param {string} message 送信するテキスト
 */

function sendSlack(message) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK'); 
  
  const payload = {
    "text": message
  };
  
  const options = {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    console.error("Slack送信失敗: " + e.toString());
  }
}

/**
 * 初回設定用：各種APIキーをスクリプトプロパティに安全に保存する関数
 * ※一度実行したら、コード内の生キーは削除してOK
 */
function setFinalSecrets() {
  const props = PropertiesService.getScriptProperties();

  // ここに実際のキーを入力して一度だけ実行する
  props.setProperty('GOOGLE_API_KEY', 'GOOGLE_API_KEY');
  props.setProperty('GOOGLE_CX', 'GOOGLE_CX');
  props.setProperty('DIFY_API_KEY', 'DIFY_API_KEY');
  props.setProperty('SLACK_WEBHOOK', 'https://hooks.slack.com/SLACK_WEBHOOK');
  
  console.log("✅ セキュリティ設定完了！これでコードから生キーを消せます。");
}
