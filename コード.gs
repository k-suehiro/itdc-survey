// ===============================================================
// ★★★ 統合履歴管理型アンケートシステム V1.3 (高速応答版) ★★★
// ===============================================================

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const HISTORY_SHEET_NAME = '履歴';

function doGet(e) {
  // (この部分は変更ありません)
  if (e.parameter.surveyId) {
    const template = HtmlService.createTemplateFromFile('form');
    template.surveyId = e.parameter.surveyId;
    return template.evaluate().setTitle('ITDCフィードバックアンケート');
  } else {
    return HtmlService.createHtmlOutputFromFile('index').setTitle('アンケート送付ツール');
  }
}

/**
 * 【ステップ1：高速処理】
 * 履歴シートに記録し、メール送信に必要な情報を返す（すぐに完了する）
 */
function logAndPrepareEmail(recipientEmail, recipientName, subject) {
  try {
    if (!recipientEmail || !subject) {
      throw new Error('宛先と案件名を入力してください。');
    }
    const operatorEmail = Session.getActiveUser().getEmail();
    const uniqueId = Math.random().toString(36).slice(-4);
    const surveyUrl = `${ScriptApp.getService().getUrl()}?surveyId=${uniqueId}`;

    // 履歴シートに記録
    const historySheet = SPREADSHEET.getSheetByName(HISTORY_SHEET_NAME);
    historySheet.appendRow([
      uniqueId, '処理開始', new Date(), operatorEmail, recipientEmail, subject
    ]);

    // 次のステップ（メール送信）に必要な情報をまとめて返す
    return {
      recipientEmail: recipientEmail,
      recipientName: recipientName,
      subject: subject,
      surveyUrl: surveyUrl
    };

  } catch (e) {
    console.error(e);
    // エラーオブジェクトを返すことで、より詳細なエラーをクライアントに伝える
    return { error: e.message };
  }
}

/**
 * 【ステップ2：低速処理】
 * メールをバックグラウンドで送信する（時間がかかる）
 */
function sendEmailInBackground(emailDetails) {
  try {
    const displayName = emailDetails.recipientName || emailDetails.recipientEmail.split('@')[0];
    const emailSubject = 'ITDCの対応に関するフィードバックのお願い';
    const htmlBody = `
      ${displayName}様<br><br>
      この度は、ツールのご利用や各種ご相談など、誠にありがとうございました。<br><br>
      今後の開発やサポート品質の向上のため、星を選択するだけ（5問・30秒程度）の簡単なアンケートにご協力いただけますと幸いです。<br><br>
      ▼アンケートページ<br>
      <a href="${emailDetails.surveyUrl}">${emailDetails.surveyUrl}</a><br><br>
      本アンケートは、ご回答者様が特定される仕組みにはなっておりません。どうぞご安心の上、ご回答ください。<br><br>
      ----------------------------------------------------<br>
      ※このメールは送信専用です。ご返信いただいても対応はできかねますのでご了承ください。<br>
      ----------------------------------------------------<br>
      ITDC
    `;

    GmailApp.sendEmail(emailDetails.recipientEmail, emailSubject, "", {
      from: 'itdc@crestec.co.jp',
      name: 'ITDC',
      htmlBody: htmlBody
    });
    
    // バックグラウンドでの成功をログに残す
    console.log(`メール送信成功: ${emailDetails.recipientEmail}`);

  } catch (e) {
    // バックグラウンドでのエラーをログに残す
    console.error(`メール送信失敗: ${e.message}`);
  }
}

/**
 * 回答ページから呼び出され、回答をスプレッドシートに保存する
 */
function saveResponse(responseObject) {
  // （この関数は変更ありません）
  try {
    const surveyId = responseObject.surveyId;
    if (!surveyId) { throw new Error('ユニークIDが見つかりません。'); }

    const historySheet = SPREADSHEET.getSheetByName(HISTORY_SHEET_NAME);
    const data = historySheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === surveyId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) { throw new Error('元の送信記録が見つかりませんでした。'); }

    if (data[targetRow - 1][1] === '回答済み') { return { duplicate: true }; }

    const rowData = data[targetRow - 1];
    const operatorEmail = rowData[3];
    const recipientEmail = rowData[4];
    const subject = rowData[5];

    historySheet.getRange(targetRow, 2).setValue('回答済み');
    historySheet.getRange(targetRow, 7).setValue(new Date());
    historySheet.getRange(targetRow, 8).setValue(responseObject.q1);
    historySheet.getRange(targetRow, 9).setValue(responseObject.q2);
    historySheet.getRange(targetRow, 10).setValue(responseObject.q3);
    historySheet.getRange(targetRow, 11).setValue(responseObject.q4);
    historySheet.getRange(targetRow, 12).setValue(responseObject.q5);
    historySheet.getRange(targetRow, 13).setValue(responseObject.comment || '');

    sendNotificationToOperator(operatorEmail, recipientEmail, subject, responseObject);

    return { duplicate: false };
  } catch (e) {
    console.error(e);
    return { error: e.message };
  }
}

/**
 * 回答があったことをオペレーター宛に通知する
 */
function sendNotificationToOperator(operatorEmail, recipientEmail, subject, responseObject) {
  try {
    const stars = (n) => '★'.repeat(n) + '☆'.repeat(5 - n);
    const avg = ((+responseObject.q1 + +responseObject.q2 + +responseObject.q3 + +responseObject.q4 + +responseObject.q5) / 5).toFixed(1);
    const now = new Date().toLocaleString('ja-JP');
    const commentRow = responseObject.comment
      ? `<tr><td colspan="2" style="padding:8px;"><strong>コメント:</strong><br>${responseObject.comment.replace(/\n/g, '<br>')}</td></tr>`
      : '';

    const htmlBody = `
      アンケートへの回答がありました。<br><br>
      <strong>案件名:</strong> ${subject}<br>
      <strong>回答者:</strong> ${recipientEmail}<br>
      <strong>回答日時:</strong> ${now}<br><br>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse; font-size:14px;">
        <tr style="background:#f0f0f0;"><th>質問</th><th>評価</th></tr>
        <tr><td>1. 総合満足度</td><td>${stars(+responseObject.q1)} (${responseObject.q1})</td></tr>
        <tr><td>2. 業務への役立ち度</td><td>${stars(+responseObject.q2)} (${responseObject.q2})</td></tr>
        <tr><td>3. 説明の分かりやすさ</td><td>${stars(+responseObject.q3)} (${responseObject.q3})</td></tr>
        <tr><td>4. スピード感</td><td>${stars(+responseObject.q4)} (${responseObject.q4})</td></tr>
        <tr><td>5. 今後の依頼意向</td><td>${stars(+responseObject.q5)} (${responseObject.q5})</td></tr>
        <tr style="background:#f0f0f0;"><td><strong>平均スコア</strong></td><td><strong>${avg} / 5.0</strong></td></tr>
        ${commentRow}
      </table>
    `;

    GmailApp.sendEmail(operatorEmail, `【アンケート回答】${subject}`, '', {
      from: 'itdc@crestec.co.jp',
      name: 'ITDC アンケートシステム',
      htmlBody: htmlBody
    });
  } catch (e) {
    console.error(`通知メール送信失敗: ${e.message}`);
  }
}

/**
 * 未回答のアンケートに対してリマインダーメールを送る（毎日トリガーで実行）
 * ステータスが「処理開始」かつ送信から3日以上経過した行が対象
 */
function sendReminders() {
  const historySheet = SPREADSHEET.getSheetByName(HISTORY_SHEET_NAME);
  const data = historySheet.getDataRange().getValues();
  const threeDaysAgo = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
  const baseUrl = ScriptApp.getService().getUrl();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] !== '処理開始') continue;
    if (new Date(data[i][2]) > threeDaysAgo) continue;

    const uniqueId = data[i][0];
    const recipientEmail = data[i][4];
    const subject = data[i][5];
    const surveyUrl = `${baseUrl}?surveyId=${uniqueId}`;

    try {
      const htmlBody = `
        先日お送りしたアンケートへのご回答がまだ届いておりません。<br>
        お忙しいところ恐れ入りますが、ご協力いただけますと幸いです。<br><br>
        ▼アンケートページ<br>
        <a href="${surveyUrl}">${surveyUrl}</a><br><br>
        本アンケートは、ご回答者様が特定される仕組みにはなっておりません。どうぞご安心の上、ご回答ください。<br><br>
        ----------------------------------------------------<br>
        ※このメールは送信専用です。ご返信いただいても対応はできかねますのでご了承ください。<br>
        ----------------------------------------------------<br>
        ITDC
      `;
      GmailApp.sendEmail(recipientEmail, 'ITDCの対応に関するフィードバックのお願い（リマインダー）', '', {
        from: 'itdc@crestec.co.jp',
        name: 'ITDC',
        htmlBody: htmlBody
      });
      historySheet.getRange(i + 1, 2).setValue('リマインダー送信済み');
      console.log(`リマインダー送信: ${recipientEmail}`);
    } catch (e) {
      console.error(`リマインダー送信失敗: ${recipientEmail} - ${e.message}`);
    }
  }
}

/**
 * 宛先リスト用のスプレッドシートから宛先のリストを取得する
 */
function getRecipientData() {
  // （この関数は変更ありません）
  try {
    // ★★★【重要】★★★
    // 「Googleアカウント管理DB」のIDを指定
    const sheetId = '1Q9Qdk7K1t_L0KcI0I_J7W62fFHO8i4SW9IZ1jb6L4-k'; 
    const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) { return []; }
    
    const range = sheet.getRange('A2:B' + lastRow);
    const values = range.getValues();
    
    const data = values
      .filter(row => row[0] && row[1])
      .map(row => ({ email: row[0].trim(), name: row[1].trim() }));
      
    return data;
  } catch (e) {
    console.error('getRecipientDataでエラー:', e);
    return [];
  }
}