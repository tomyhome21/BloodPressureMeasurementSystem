// --- [ 定数設定 ] ----------------------------------------------------------------------------------
const MAIN_SHEET_NAME = '血圧測定データ';
const LINK_SHEET_NAME = '血圧管理情報';
const NAME_CELL = 'B2';
const GRAPH_ID_CELL = 'A1'; // グラフIDを格納（削除対象1）
const GRAPH_LINK_CELL = 'A3'; 
const PARENT_LINK_CELL = 'A6'; // 子シート判定用
const PARENT_LINK_HEADER = '親シートに戻る (テンプレート)';
const DEFAULT_GRAPH_SPREADSHEET_NAME = '血圧測定管理 - 長期推移グラフ'; 
const HEADERS = ['日付', '時刻', '最高血圧', '最低血圧', '脈拍'];
const DEFAULT_TITLE_ROW = '血圧の記録';
const DEFAULT_SPREADSHEET_TITLE = '血圧測定管理'; 
const DATA_START_ROW = 3; 
const DATE_COL = 1; 
const TIME_COL = 2; 

// --- [ カスタムメニューの作成 ] ----------------------------------------------------------------------

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚙️ 血圧データ処理')
      .addItem('データ処理を実行（クリップボード入力）', 'checkUserNameAndOpenDialog')
      .addSeparator()
      .addItem('使用者名を変更する', 'openNameInputDialog')
      .addSeparator()
      .addItem('⚠️ スプレッドシートを初期化', 'initializeSpreadsheet') 
      .addItem('📄 新たな個人データ管理を作成', 'createPersonalCopy') 
      .addItem('🗑️ **個人シートを削除**', 'deletePersonalCopy') 
      .addToUi();
}

/**
 * 使用者名が設定されているかチェックし、設定されていれば入力ダイアログを開く
 */
function checkUserNameAndOpenDialog() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const existingLinkSheet = getExistingSheet(spreadsheet, LINK_SHEET_NAME);
    let userName = "";

    if (existingLinkSheet) {
        userName = existingLinkSheet.getRange(NAME_CELL).getValue();
    }
    
    if (!userName || userName === '') {
        ui.alert('エラー', '使用者名が設定されていません。\n\n「⚙️ 血圧データ処理」メニューから「使用者名を変更する」を選択し、名前を設定してからデータ処理を実行してください。', ui.ButtonSet.OK);
        return;
    }
    
    openInputDialog();
}


/**
 * 名前入力用HTMLダイアログを表示する
 */
function openNameInputDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINK_SHEET_NAME);
  const currentName = sheet ? sheet.getRange(NAME_CELL).getValue() : '';
  
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; }
          input[type="text"] { width: 100%; padding: 8px; box-sizing: border-box; }
          .button-container { text-align: right; margin-top: 15px; }
        </style>
      </head>
      <body>
        <p>この血圧測定管理簿の**使用者名**を入力してください。</p>
        <label for="userName">使用者名:</label>
        <input type="text" id="userName" value="${currentName || ''}">
        <div class="button-container">
          <input type="button" value="設定/変更" onclick="setName();">
        </div>
        <script>
          function setName() {
            const userName = document.getElementById('userName').value;
            if (userName.trim() === '') {
              alert('名前を入力してください。');
              return;
            }
            google.script.run
              .withSuccessHandler(function(){
                alert('名前が設定されました。シート名やグラフシート名に反映されます。');
                google.script.host.close();
              })
              .withFailureHandler(function(e){ alert('エラーが発生しました: ' + e); google.script.host.close(); })
              .setUserNameAndTitles(userName);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '使用者名の設定');
}

/**
 * ユーザー名を設定し、関連するシートのタイトルを更新する
 */
function setUserNameAndTitles(userName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = getSheet(spreadsheet, LINK_SHEET_NAME, 1);
    
    linkSheet.getRange(GRAPH_ID_CELL).clearContent(); 
    linkSheet.getRange(GRAPH_LINK_CELL).clearContent();
    
    linkSheet.getRange(GRAPH_ID_CELL).setValue('グラフID (非表示)').setFontColor('white');
    linkSheet.getRange('A2').setValue('使用者名:').setFontWeight('bold').setBackground('#fce5cd');
    linkSheet.setColumnWidth(1, 150);
    linkSheet.setColumnWidth(2, 250);

    linkSheet.getRange(NAME_CELL).setValue(userName).setFontWeight('bold').setFontSize(12).setBackground('#fff2cc');
    
    if (linkSheet.getLastRow() >= 5) {
        linkSheet.getRange(5, 1, linkSheet.getLastRow() - 4, 3).clearContent();
    }

    const newTitle = `${userName}さんの血圧測定管理`;
    spreadsheet.rename(newTitle);
    
    const graphSpreadsheet = getGraphSpreadsheetIfExist(linkSheet);
    if (graphSpreadsheet) {
        const newGraphTitle = `${userName}さんの${DEFAULT_GRAPH_SPREADSHEET_NAME}`;
        graphSpreadsheet.rename(newGraphTitle);
        setGraphHyperlink(linkSheet, graphSpreadsheet.getUrl(), userName);
    }
}

/**
 * カスタムHTMLダイアログを表示する (クリップボード貼り付け用)
 */
function openInputDialog() {
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; padding: 15px; }
          textarea { width: 100%; box-sizing: border-box; resize: none; border: 1px solid #ccc; }
          .button-container { text-align: right; margin-top: 10px; }

          /* --- アニメーションの設定 --- */
          
          /* 1. クルクル回転させる設定 */
          @keyframes rotate-icon {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .spinning {
            display: inline-block;
            animation: rotate-icon 2s linear infinite; /* 2秒で1回転を繰り返す */
          }

          /* 2. 文字を点滅させる設定 */
          @keyframes flash-text {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.3; }
          }
          .flashing {
            animation: flash-text 1.5s ease-in-out infinite; /* 1.5秒周期で点滅 */
          }

          #loadingArea { display: none; text-align: center; padding-top: 50px; }
        </style>
      </head>
      <body>
        <div id="inputArea" style="display: block;">
          <label for="clipboardData" style="font-weight:bold;">血圧データを貼り付け:</label><br><br>
          <textarea id="clipboardData" rows="10" placeholder="ここに貼り付けてください"></textarea>
          <div class="button-container">
            <input type="button" value="処理を実行" style="padding: 10px 20px;" onclick="runProcess();">
          </div>
        </div>

        <div id="loadingArea" style="display: none;">
          <h3 class="flashing" style="color: #444;">データを照合・更新中...</h3>
          <div class="spinning" style="font-size: 60px; margin: 20px;">⌛</div>
          <p style="color: #666;">完了通知が出るまで、そのままお待ちください。</p>
        </div>

        <script>
          function runProcess() {
            const rawText = document.getElementById('clipboardData').value;
            if (rawText.trim() === '') {
              alert('データが入力されていません。');
              return;
            }

            document.getElementById('inputArea').style.display = 'none';
            document.getElementById('loadingArea').style.display = 'block';

            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(e){ 
                alert('エラー: ' + e); 
                google.script.host.close(); 
              })
              .processInputData(rawText);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '血圧データ入力');
}


// --- [ メイン処理 ] --------------------------------------------------------------------------------

/**
 * クリップボードからのデータを受け取り、クリーニングしてシートに書き込み、メイン処理を継続する
 */
/**
 * クリップボードからのデータを受け取り、全処理を実行する（メイン関数）
 * * @param {string} rawTextFromClipboard 貼り付けられた生テキスト
 */
function processInputData(rawTextFromClipboard) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    try {
        // リンクシートから使用者名を取得
        const existingLinkSheet = getExistingSheet(spreadsheet, LINK_SHEET_NAME);
        const userName = existingLinkSheet.getRange(NAME_CELL).getValue(); 

        // メインデータシート（集約シート）を取得
        const mainSheet = getSheet(spreadsheet, MAIN_SHEET_NAME, 0);

        // ---【修正ポイント：既存データの読み込み】--------------------------------------
        // 集約シートの3行目から、現在入っているデータをすべて読み出す
        let existingData = [];
        if (mainSheet.getLastRow() >= DATA_START_ROW) {
            existingData = mainSheet.getRange(
                DATA_START_ROW, 
                1, 
                mainSheet.getLastRow() - DATA_START_ROW + 1, 
                HEADERS.length
            ).getValues();
        }
        // ----------------------------------------------------------------------------

        // データのクリーニング（既存データを第2引数として渡し、合流・重複排除を行う）
        const { allRecords } = cleanAndFilterData(rawTextFromClipboard, existingData);
        
        if (allRecords.length === 0) {
            ui.alert('警告', '有効な血圧データが見つかりませんでした。', ui.ButtonSet.OK);
            return;
        }

        // 1. メインデータシート（集約シート）の更新
        // ここで「過去データ + 新規データ」の合体版（allRecords）を全書き出しする
        updateMainDataSheet(mainSheet, allRecords, userName); 

        // 2. グラフ専用スプレッドシートの取得、または作成
        const linkSheet = getSheet(spreadsheet, LINK_SHEET_NAME, 1); 
        let graphSpreadsheet = getOrCreateGraphSpreadsheet(spreadsheet, linkSheet, userName); 
        
        // 3. データの分類（月別および時間帯別）
        const { timeSplitData, allMonthlyData } = processAndSplitData(allRecords);

        // 4. メインスプレッドシート内の各月別シートの作成・更新
        // 分割されたデータを元に、表示用の個別シートを再生成する
        updateMonthlySheets(spreadsheet, timeSplitData, 2); 
        
        // 5. グラフ専用スプレッドシートの各期間別データとチャートの更新
        if (graphSpreadsheet) {
            updateGraphDataAndCharts(graphSpreadsheet, allMonthlyData, userName); 
        }

        ui.alert('完了', '過去データを含めてデータの処理とグラフの更新がすべて完了しました！', ui.ButtonSet.OK);

    } catch (e) {
        ui.alert('重大なエラーが発生しました', '詳細情報: ' + e.message, ui.ButtonSet.OK);
        console.error('Stack: ' + e.stack);
    }
}

// --- [ ユーティリティ機能: 初期化とコピー ] ----------------------------------------------------

/**
 * ★ 機能1: スプレッドシートを初期状態に戻す
 */
function initializeSpreadsheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // --- 0. 子シート判定による初期化制限チェック ---
    const existingLinkSheet = getExistingSheet(spreadsheet, LINK_SHEET_NAME);
    if (existingLinkSheet) {
        const parentLinkContent = String(existingLinkSheet.getRange(PARENT_LINK_CELL).getDisplayValue()).trim();

        if (parentLinkContent === PARENT_LINK_HEADER) {
            ui.alert(
                '❌ 初期化制限', 
                'このスプレッドシートは個人データ管理用の子シートです。\n\nこのシートを削除したい場合は、**親シート**に戻って操作してください。', 
                ui.ButtonSet.OK
            );
            return; 
        }
    }
    // ------------------------------------------

    const response = ui.alert(
        '⚠️ 最終確認 (初期化)',
        '「血圧測定データ」シート以外の全てのシート、リンク情報、**グラフ専用スプレッドシートファイル**を削除します。\nスプレッドシートのタイトルも初期名に戻ります。\n\nよろしいですか？',
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
        return;
    }

    try {
        // 1. グラフシートのIDを取得し、存在すればファイルを完全に削除 (ゴミ箱へ)
        if (existingLinkSheet) {
            const graphSpreadsheetId = String(existingLinkSheet.getRange(GRAPH_ID_CELL).getValue()).trim();

            if (graphSpreadsheetId) {
                try {
                    // DriveApp を使用してファイルをゴミ箱へ移動
                    DriveApp.getFileById(graphSpreadsheetId).setTrashed(true);
                    
                    // リンクセル情報(A1, A3)をクリア
                    existingLinkSheet.getRange(GRAPH_ID_CELL).clearContent(); 
                    existingLinkSheet.getRange(GRAPH_LINK_CELL).clearContent(); 
                } catch (e) {
                    Logger.log('Failed to trash graph spreadsheet (ID: ' + graphSpreadsheetId + '). Error: ' + e.message);
                }
            }
            
            // A5以降の子シート/個人シートリンク情報をクリア (親シートでのみ有効)
            if (existingLinkSheet.getLastRow() >= 5) {
                existingLinkSheet.getRange(5, 1, existingLinkSheet.getLastRow() - 4, 3).clearContent();
            }
        }
        
        // 2. メインスプレッドシート内のシートを削除（「血圧測定データ」以外）
        spreadsheet.getSheets().forEach(sheet => {
            if (sheet.getName() !== MAIN_SHEET_NAME) {
                spreadsheet.deleteSheet(sheet);
            }
        });
        
        // 3. スプレッドシート名を初期名に戻す
        spreadsheet.rename(DEFAULT_SPREADSHEET_TITLE);
        
        // 4. メインシートの内容を初期化
        const mainSheet = getSheet(spreadsheet, MAIN_SHEET_NAME, 0); 
        
        if (mainSheet.getLastRow() >= DATA_START_ROW) {
            mainSheet.getRange(DATA_START_ROW, 1, mainSheet.getLastRow() - DATA_START_ROW + 1, mainSheet.getLastColumn()).clearContent();
        }
        
        mainSheet.getRange(1, 1).setValue(DEFAULT_TITLE_ROW).setFontSize(14).setFontWeight('bold').setBackground('#d9ead3');
        mainSheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#b6d7a8');
        mainSheet.setFrozenRows(2); 

        ui.alert('初期化完了', `スプレッドシートを初期状態に戻しました。\n\nグラフ専用ファイルも削除されました（ゴミ箱を確認してください）。`, ui.ButtonSet.OK);

    } catch (e) {
        ui.alert('エラーが発生しました', '初期化処理中にエラー: ' + e.message, ui.ButtonSet.OK);
    }
}


/**
 * ★ 機能2: 現在のスプレッドシートをコピーし、新たな個人名でセットアップする
 */
function createPersonalCopy() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // --- 0. 複製制限チェック (A6セルの内容を確認) ---
    const currentLinkSheet = getExistingSheet(spreadsheet, LINK_SHEET_NAME);
    if (currentLinkSheet) {
        const parentLinkContent = String(currentLinkSheet.getRange(PARENT_LINK_CELL).getDisplayValue()).trim();

        if (parentLinkContent === PARENT_LINK_HEADER) {
            ui.alert(
                '⚠️ 複製制限', 
                'このスプレッドシートは「個人データ管理用」の子シートです。\n新規に個人データを作成する場合は、**元の親スプレッドシート**に戻って操作してください。', 
                ui.ButtonSet.OK
            );
            return;
        }
    }
    // ------------------------------------------------------------------

    // --- 1. 親シートの存在と名前のチェック ---
    let userName = "";

    if (currentLinkSheet) {
        userName = currentLinkSheet.getRange(NAME_CELL).getValue();
    }

    if (!currentLinkSheet || !userName || userName.trim() === '') {
        ui.alert(
            '⚠️ 必須情報の不足', 
            '「⚙️ 血圧データ処理」メニューから「使用者名を変更する」を選択し、名前を設定してから再度実行してください。', 
            ui.ButtonSet.OK
        );
        return;
    }
    // ------------------------------------------

    // 2. 新しい名前を取得するためのダイアログ
    const nameResponse = ui.prompt(
        '新たな個人名の設定',
        '新しいスプレッドシートで使用する**使用者名**を入力してください。',
        ui.ButtonSet.OK_CANCEL
    );

    if (nameResponse.getSelectedButton() !== ui.Button.OK || !nameResponse.getResponseText().trim()) {
        ui.alert('キャンセル', '名前が入力されなかったため、コピー作成をキャンセルしました。', ui.ButtonSet.OK);
        return;
    }
    
    const newUserName = nameResponse.getResponseText().trim();
    const newSpreadsheetTitle = `${newUserName}さんの血圧測定管理`;
    
    try {
        // 3. スプレッドシートをコピー
        const parentUrl = spreadsheet.getUrl(); // 親シートのURLを取得
        const newSpreadsheet = spreadsheet.copy(newSpreadsheetTitle);
        const newUrl = newSpreadsheet.getUrl(); 
        
        // 4. コピーしたスプレッドシートの初期化 
        copyInitializeSheets(newSpreadsheet);

        // 5. 名前を設定し、タイトルなどを更新 (親シートのURLを渡す)
        setUserNameAndTitlesInCopy(newSpreadsheet, newUserName, parentUrl); 
        
        // 6. 親シートの「血圧管理情報」に新しい個人シートのURLを記録する
        recordLinkToLinkSheet(spreadsheet, newUserName, newUrl);
        
        // 7. 完了アラート
        const alertMessage = 
            `「${newUserName}さん」用の個人シートを作成しました。\n\n` +
            `「${LINK_SHEET_NAME}」シートを参照してください。`
            
        if (ui) {
            ui.alert('コピー作成完了', alertMessage, ui.ButtonSet.OK);
        }
        
    } catch (e) {
        if (ui) {
            ui.alert('エラーが発生しました', 'コピー作成中にエラー: ' + e.message, ui.ButtonSet.OK);
        } else {
            Logger.log('コピー作成中にエラーが発生しました: ' + e.message);
        }
    }
}

/**
 * コピーされたスプレッドシートから不要なシートとデータを削除し、初期化する。
 * (initializeSpreadsheetのサブセット的な機能)
 */
function copyInitializeSheets(newSpreadsheet) {
    // 1. メインシート（血圧測定データ）以外のシートを削除
    newSpreadsheet.getSheets().forEach(sheet => {
        if (sheet.getName() !== MAIN_SHEET_NAME) {
            newSpreadsheet.deleteSheet(sheet);
        }
    });

    // 2. メインシートをアクティブにして、データをクリア
    const mainSheet = newSpreadsheet.getSheetByName(MAIN_SHEET_NAME);
    if (mainSheet) {
        newSpreadsheet.setActiveSheet(mainSheet);
        newSpreadsheet.moveActiveSheet(0); // 1番目に移動
        
        // データ範囲をクリア（タイトル行は残す）
        const DATA_START_ROW = 3; 
        if (mainSheet.getLastRow() >= DATA_START_ROW) {
            mainSheet.getRange(DATA_START_ROW, 1, mainSheet.getLastRow() - DATA_START_ROW + 1, mainSheet.getLastColumn()).clearContent();
        }
    }
}


/**
 * コピーしたスプレッドシートの名前設定とシートタイトル更新を行う
 */
function setUserNameAndTitlesInCopy(newSpreadsheet, userName, parentUrl) { 
    const linkSheet = getSheet(newSpreadsheet, LINK_SHEET_NAME, 1);
    
    linkSheet.getRange(GRAPH_ID_CELL).setValue('グラフID (非表示)').setFontColor('white');
    linkSheet.getRange('A2').setValue('使用者名:').setFontWeight('bold').setBackground('#fce5cd');
    linkSheet.setColumnWidth(1, 150);
    linkSheet.setColumnWidth(2, 250);

    linkSheet.getRange(NAME_CELL).setValue(userName).setFontWeight('bold').setFontSize(12).setBackground('#fff2cc');
    
    // 子シートのA6セルに親シートへのリンクを記録する（子シートフラグ）
    const parentLinkFormula = `=HYPERLINK("${parentUrl}", "${PARENT_LINK_HEADER}")`;
    linkSheet.getRange(PARENT_LINK_CELL).setValue(parentLinkFormula); 
    linkSheet.getRange(PARENT_LINK_CELL).setFontSize(14).setFontWeight('bold').setBackground('#f3f3f3');
    
    const mainSheet = newSpreadsheet.getSheetByName(MAIN_SHEET_NAME);
    if (mainSheet) {
        mainSheet.getRange(1, 1).setValue(`${userName}さんの${DEFAULT_TITLE_ROW}`).setFontSize(14).setFontWeight('bold').setBackground('#d9ead3');
    }
}

/**
 * 親シートの「血圧管理情報」に新しい個人シートのURLを記録する
 */
function recordLinkToLinkSheet(parentSpreadsheet, userName, newUrl) {
    const linkSheet = getSheet(parentSpreadsheet, LINK_SHEET_NAME, 1); 

    const HEADER_ROW = 5; 
    
    const currentHeaderContent = linkSheet.getRange(HEADER_ROW, 1).getValue();

    if (currentHeaderContent !== '【作成済み個人シート】') {
         linkSheet.getRange(HEADER_ROW, 1, 1, 3).setValues([['【作成済み個人シート】', '', '']])
             .setFontWeight('bold').setBackground('#fce5cd').mergeAcross();
    }
    
    const dataStartRow = HEADER_ROW + 1;
    
    const existingValues = linkSheet.getRange(dataStartRow, 1, linkSheet.getMaxRows() - dataStartRow + 1, 3).getValues();
    
    let nextRowOffset = 0;
    for (let i = 0; i < existingValues.length; i++) {
        if (existingValues[i][0] === '') {
            break;
        }
        nextRowOffset++;
    }
    
    const nextRow = dataStartRow + nextRowOffset;

    const linkFormula = `=HYPERLINK("${newUrl}", "${userName}さんの血圧測定管理")`;
    
    linkSheet.getRange(nextRow, 1).setValue(userName);
    linkSheet.getRange(nextRow, 2).setValue(linkFormula);
    
    linkSheet.setColumnWidth(1, 150);
    linkSheet.setColumnWidth(2, 400); 
    
    linkSheet.getRange(nextRow, 1, 1, 2).setFontSize(14).setFontWeight('bold');
    linkSheet.getRange(nextRow, 1, 1, 2).setBackground('#ebf1de'); 
}


// --- [ 機能3: 親シートからの子シート削除 ] ----------------------------------------------------

/**
 * 親シートから子シートの一覧を表示し、選択された子シートを削除（Driveからゴミ箱へ移動）する
 */
function deletePersonalCopy() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const linkSheet = getExistingSheet(spreadsheet, LINK_SHEET_NAME);

    // 1. 親シート判定 (A6にPARENT_LINK_HEADERがないこと)
    if (!linkSheet || String(linkSheet.getRange(PARENT_LINK_CELL).getDisplayValue()).trim() === PARENT_LINK_HEADER) {
        ui.alert('❌ 実行制限', 'この機能は**親シート**でのみ実行可能です。', ui.ButtonSet.OK);
        return;
    }

    // 2. リンク一覧のデータを取得（A6以降）
    const HEADER_ROW = 5; 
    const dataStartRow = HEADER_ROW + 1; 
    
    const allDataRange = linkSheet.getRange(dataStartRow, 1, linkSheet.getMaxRows() - dataStartRow + 1, 3);
    
    const allNames = allDataRange.getValues();      
    const allFormulas = allDataRange.getFormulas(); 
    
    const linkRecordsForExecution = []; 
    const names = [];

    allFormulas.forEach((row, index) => {
        
        if (allNames[index][0] && allNames[index][0].toString().trim() !== '') {
            
            const formula = row[1]; 
            
            if (formula && formula.toString().startsWith('=HYPERLINK')) {
                const urlMatch = formula.toString().match(/=HYPERLINK\("([^"]+)"/i);
                const url = urlMatch ? urlMatch[1] : null;
                
                if (url) {
                    linkRecordsForExecution.push({
                        name: allNames[index][0], 
                        url: url,
                        sheetRow: dataStartRow + index
                    });
                    names.push(allNames[index][0]);
                }
            }
        }
    });

    if (linkRecordsForExecution.length === 0) {
        ui.alert('情報', '現在、削除可能な個人シートのリンクは記録されていません。', ui.ButtonSet.OK);
        return;
    }
    
    // PropertiesServiceを使用してリンク情報を一時保存する
    PropertiesService.getScriptProperties().setProperty('temp_linkRecords', JSON.stringify(linkRecordsForExecution));

    // 4. 削除ダイアログを表示
    const selectHtmlTemplate = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style> body { font-family: sans-serif; } select { width: 100%; padding: 8px; margin-bottom: 15px; } </style>
        </head>
        <body>
          <p>削除したい個人シートを選択してください。<br>選択されたファイルと紐づくグラフは**ゴミ箱に移動**されます。</p>
          <select id="targetSheet">
            ${names.map((name, index) => `<option value="${index}">${name}さんの管理簿</option>`).join('')}
          </select>
          <input type="button" value="選択したシートを削除" onclick="deleteSheet();">
          <script>
            function deleteSheet() {
              const select = document.getElementById('targetSheet');
              const index = select.value;
              if (index !== null) {
                google.script.run
                  .withSuccessHandler(function() {
                    alert('削除処理が完了しました。親シートのリンク情報も削除されました。');
                    google.script.host.close();
                  })
                  .withFailureHandler(function(e) {
                    alert('削除処理中にエラーが発生しました: ' + e);
                    google.script.host.close();
                  })
                  .executeDelete(index);
              }
            }
          </script>
        </body>
      </html>
    `;
    const htmlOutput = HtmlService.createHtmlOutput(selectHtmlTemplate)
        .setWidth(400)
        .setHeight(250);
    ui.showModalDialog(htmlOutput, '個人シートの削除');
}

/**
 * deletePersonalCopyからコールバックされる実際の削除実行関数
 */
function executeDelete(indexStr) {
    const index = parseInt(indexStr, 10);
    
    // PropertiesServiceからリンク情報を読み込む
    const tempRecordsString = PropertiesService.getScriptProperties().getProperty('temp_linkRecords');
    
    try {
        if (!tempRecordsString) {
            throw new Error('一時データが見つかりません。');
        }

        const linkRecords = JSON.parse(tempRecordsString);
        const record = linkRecords[index]; 

        if (!record) {
            throw new Error('削除対象レコードが見つかりません。');
        }
        
        // 子シートのファイルIDをURLから抽出
        let childSpreadsheetId = null; 
        const childSpreadsheetIdMatch = record.url.match(/d\/([a-zA-Z0-9_-]+)/);
        
        if (childSpreadsheetIdMatch) {
            childSpreadsheetId = childSpreadsheetIdMatch[1];
        }

        if (childSpreadsheetId) {
            try {
                // 1. 子シートを開き、A1からグラフIDを取得
                const childSpreadsheet = SpreadsheetApp.openById(childSpreadsheetId);
                const childLinkSheet = childSpreadsheet.getSheetByName(LINK_SHEET_NAME);
                
                if (childLinkSheet) {
                    // A1セルからグラフIDを取得
                    const graphFileId = String(childLinkSheet.getRange(GRAPH_ID_CELL).getValue()).trim();

                    if (graphFileId) {
                        // 2. グラフファイルをゴミ箱へ移動
                        DriveApp.getFileById(graphFileId).setTrashed(true);
                        Logger.log(`Associated graph file ${graphFileId} trashed successfully.`);
                    }
                }
            } catch (e) {
                // 子シートが既に削除されている、またはアクセスできない場合
                Logger.log(`Could not open child sheet or delete graph file for URL: ${record.url}. Error: ${e.message}`);
            }
        }
        
        // 3. 子シート本体のファイルをゴミ箱へ移動
        if (childSpreadsheetId) {
            try {
                DriveApp.getFileById(childSpreadsheetId).setTrashed(true);
                Logger.log(`Child sheet file ${childSpreadsheetId} trashed successfully.`);
            } catch (e) {
                // ファイルの削除に失敗しても、リンク情報は削除する
                Logger.log(`Failed to trash child sheet file ${childSpreadsheetId}: ${e.message}. Proceeding to delete link.`);
            }
        }

        // 4. 親シートのリンク一覧から該当行を削除
        const linkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINK_SHEET_NAME);
        if (linkSheet) {
            linkSheet.deleteRow(record.sheetRow); 
        }
        
        return true; 
    } catch(e) {
        throw e;
    } finally {
        // 処理の成功・失敗に関わらず、一時データを削除する
        PropertiesService.getScriptProperties().deleteProperty('temp_linkRecords');
    }
}


// --- [ サブ関数: グラフ専用シートの作成とリンク / データ処理関連 ] ------------------------------------------------------

/**
 * 既存のグラフ専用スプレッドシートを取得する (存在しない場合はnull)
 */
function getGraphSpreadsheetIfExist(linkSheet) {
    const graphSpreadsheetId = String(linkSheet.getRange(GRAPH_ID_CELL).getValue()).trim(); 
    
    if (graphSpreadsheetId) {
        try {
            return SpreadsheetApp.openById(graphSpreadsheetId);
        } catch (e) {
            return null; 
        }
    }
    return null;
}

/**
 * グラフ専用スプレッドシートを取得または新規作成し、メインシートにリンクを設定する
 */
function getOrCreateGraphSpreadsheet(mainSpreadsheet, linkSheet, userName) {
  const ui = SpreadsheetApp.getUi();
  const newGraphTitle = `${userName}さんの${DEFAULT_GRAPH_SPREADSHEET_NAME}`;
  
  let graphSpreadsheet = getGraphSpreadsheetIfExist(linkSheet);

  if (graphSpreadsheet) {
      graphSpreadsheet.rename(newGraphTitle);
      setGraphHyperlink(linkSheet, graphSpreadsheet.getUrl(), userName);
      return graphSpreadsheet;
  }
  
  graphSpreadsheet = SpreadsheetApp.create(newGraphTitle);
  const graphSpreadsheetId = graphSpreadsheet.getId();
  
  linkSheet.getRange(GRAPH_ID_CELL).setValue(graphSpreadsheetId);
  linkSheet.getRange('A1').setFontColor('white');
  setGraphHyperlink(linkSheet, graphSpreadsheet.getUrl(), userName);
  
  graphSpreadsheet.getSheets().forEach(sheet => {
    if (sheet.getName() === 'Sheet1') {
      graphSpreadsheet.deleteSheet(sheet);
    }
  });

  ui.alert('グラフ専用シートを作成しました', `新しいスプレッドシート「${newGraphTitle}」を作成し、リンクを「${LINK_SHEET_NAME}」シートに設定しました。`, ui.ButtonSet.OK);
  
  return graphSpreadsheet;
}

/**
 * リンクシートにハイパーリンクを設定する (A3セル)
 */
function setGraphHyperlink(sheet, url, userName) {
  const linkRange = sheet.getRange(GRAPH_LINK_CELL);
  const linkText = `▶︎ ${userName}さんの長期推移グラフにアクセス`;
  linkRange.setValue(`=HYPERLINK("${url}", "${linkText}")`);
  linkRange.setFontSize(14).setFontWeight('bold').setBackground('#cfe2f3');
  sheet.setColumnWidth(linkRange.getColumn(), 400);
}

/**
* クリップボードからの単一文字列データ（改行区切り）を処理する
*/
function cleanAndFilterData(rawTextFromClipboard,existingData) {
  const uniqueRecordsMap = {};
  if (existingData && existingData.length > 0) {
    existingData.forEach(row => {
      if (!row[0]) return;
      const dKey = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy/MM/dd") : row[0].toString();
      uniqueRecordsMap[`${dKey}_${row[1]}`] = [dKey, row[1], row[2], row[3], row[4]];
    });
  }
  const rawDataLines = rawTextFromClipboard.split(/\r?\n/).filter(line => line.trim().length > 0);
  
  rawDataLines.forEach(rawText => {
    rawText = rawText.trim();
	// rawText（行全体）に対して、4桁の数字より前を削除
    rawText = rawText.replace(/^.*?(?=\d{4})/, '');

    const cells = rawText.split(',').map(s => s.trim());
    
// 日本語表記の正規化
    let dateStr = cells[0].replace(/年|月/g, '/').replace(/日/g, '');
    let timeStr = cells[1].replace(/時/g, ':').replace(/分/g, '');
    const max = parseInt(cells[2], 10);
    const min = parseInt(cells[3], 10);
    
    let pulse = null;
    if (cells.length >= 5 && !isNaN(parseInt(cells[4], 10))) {
        pulse = parseInt(cells[4], 10);
    }

    if (isNaN(max) || isNaN(min)) return;

    const dateParts = dateStr.split('/');
    if (dateParts.length !== 3) return;
    
    const year = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10);
    const day = parseInt(dateParts[2], 10);
    
    if (isNaN(year) || isNaN(month) || isNaN(day)) return;

    const recordDate = new Date(year, month - 1, day);
    if (isNaN(recordDate.getTime())) return;

    let timeSlot = timeStr;
    
    if (timeSlot.includes('朝') || timeSlot.includes('夜')) {
      timeSlot = timeSlot.includes('朝') ? '朝' : '夜';
    } else if (timeSlot.includes(':')) {
        const timeParts = timeSlot.split(':');
        const hour = parseInt(timeParts[0], 10);
        
        if (isNaN(hour)) return;
        
        if (hour >= 4 && hour < 12) timeSlot = '朝';
        else if (hour >= 18 || hour < 4) timeSlot = '夜';
        else return;
    } else {
      return; 
    }
    const dateKey = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const uniqueKey = `${dateKey}_${timeSlot}`;

    uniqueRecordsMap[uniqueKey] = [dateKey, timeSlot, max, min, pulse];
  });
  
  const allRecords = Object.values(uniqueRecordsMap).sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));

  return { cleanedData: allRecords, allRecords };
}

/**
 * データシート（血圧測定データ）を更新する
 */
function updateMainDataSheet(sheet, data, userName) {
    sheet.clearContents();
    
    // 1行目: タイトルにユーザー名を追加
    sheet.getRange(1, 1).setValue(`${userName}さんの${DEFAULT_TITLE_ROW}`).setFontSize(14).setFontWeight('bold').setBackground('#d9ead3');

    // 2行目: ヘッダー
    sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#b6d7a8');
    
    // 3行目以降: データ
    if (data.length > 0) {
        sheet.getRange(DATA_START_ROW, 1, data.length, data[0].length).setValues(data);
    }
    
    sheet.setFrozenRows(2);
    sheet.autoResizeColumns(1, HEADERS.length);
}

/**
 * データを行ごとに月と時間帯で分類する
 */
function processAndSplitData(allRecords) {
  const timeSplitData = {};
  const allMonthlyData = {
    '朝': [],
    '夜': []
  };
  
  allRecords.forEach(row => {
    const dateStr = row[0];
    const timeLabel = row[1];
    
    const month = dateStr.substring(0, 7); 
    const monthSheetName = `${month}/${timeLabel}`;

    const monthlyRow = [
        dateStr, 
        row[2], 
        row[3], 
        row[4] 
    ];

    if (!timeSplitData[monthSheetName]) {
      timeSplitData[monthSheetName] = [];
    }
    timeSplitData[monthSheetName].push(monthlyRow);
    
    allMonthlyData[timeLabel].push(monthlyRow);
  });
  
  ['朝', '夜'].forEach(timeLabel => {
      allMonthlyData[timeLabel].sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
  });

  return { timeSplitData, allMonthlyData };
}

/**
 * 指定されたシートを取得する（存在しない場合は作成し、指定されたインデックスに移動する）
 */
function getSheet(spreadsheet, sheetName, index) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName, index);
  } else {
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(index);
  }
  return sheet;
}

/**
 * 指定されたシートが存在する場合のみ取得する（存在しない場合はnullを返す）
 */
function getExistingSheet(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName);
}

/**
 * メインスプレッドシート内の月別シートを更新する
 */
function updateMonthlySheets(mainSpreadsheet, timeSplitData, startIndex) {
  const sheetNames = Object.keys(timeSplitData).sort();
  const header = ['日付', '最高血圧', '最低血圧', '脈拍'];
  let sheetIndex = startIndex;
  
  // 既存の月別シートを削除
  mainSpreadsheet.getSheets().forEach(sheet => {
    if (sheet.getName().match(/\d{4}\/\d{2}\/\朝|\d{4}\/\d{2}\/\夜/)) {
      mainSpreadsheet.deleteSheet(sheet);
    }
  });

  // 新しい月別シートを作成し、データを書き込む
  sheetNames.forEach(sheetName => {
    const data = timeSplitData[sheetName];
    const newSheet = mainSpreadsheet.insertSheet(sheetName, sheetIndex++);
    
    newSheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#fce5cd');
    newSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    
    newSheet.setFrozenRows(1);
    newSheet.setColumnWidth(1, 100);
  });
}

/**
 * データの配列から最高血圧、最低血圧、脈拍の平均値を計算する
 */
function calculateAverageData(data) {
    if (data.length === 0) {
        return { max: 0, min: 0, pulse: 0 };
    }

    let sumMax = 0;
    let sumMin = 0;
    let sumPulse = 0;
    let count = 0;

    data.forEach(row => {
        const max = row[1];
        const min = row[2];
        const pulse = row[3];

        if (typeof max === 'number' && typeof min === 'number' && typeof pulse === 'number') {
            sumMax += max;
            sumMin += min;
            sumPulse += pulse;
            count++;
        }
    });
    
    if (count === 0) {
        return { max: 0, min: 0, pulse: 0 };
    }

    return {
        max: Math.round(sumMax / count),
        min: Math.round(sumMin / count),
        pulse: Math.round(sumPulse / count)
    };
}

/**
 * 計算された平均値をシートの指定された位置に表示・整形する
 */
function displayAveragesOnSheet(sheet, averages) {
    const startRow = 2; 
    const startCol = 6; 

    // 項目名 - F列
    sheet.getRange(startRow, startCol).setValue('【期間平均値】').setFontWeight('bold').setBackground('#fff2cc').setHorizontalAlignment('center');
    sheet.getRange(startRow, startCol, 1, 2).mergeAcross();

    // 最高血圧 - F列、G列
    sheet.getRange(startRow + 1, startCol).setValue('最高血圧平均:').setFontWeight('bold').setBackground('#fce5cd').setHorizontalAlignment('right');
    sheet.getRange(startRow + 1, startCol + 1).setValue(averages.max).setFontWeight('bold').setBackground('#f4cccc').setHorizontalAlignment('center');

    // 最低血圧 - F列、G列
    sheet.getRange(startRow + 2, startCol).setValue('最低血圧平均:').setFontWeight('bold').setBackground('#fce5cd').setHorizontalAlignment('right');
    sheet.getRange(startRow + 2, startCol + 1).setValue(averages.min).setFontWeight('bold').setBackground('#cfe2f3').setHorizontalAlignment('center');

    // 脈拍 - F列、G列
    sheet.getRange(startRow + 3, startCol).setValue('脈拍平均:').setFontWeight('bold').setBackground('#fce5cd').setHorizontalAlignment('right');
    sheet.getRange(startRow + 3, startCol + 1).setValue(averages.pulse).setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
    
    sheet.setColumnWidth(startCol, 120); 
    sheet.setColumnWidth(startCol + 1, 80);
}


/**
 * グラフ専用スプレッドシートの長期グラフデータとチャートを更新する
 */
function updateGraphDataAndCharts(graphSpreadsheet, allMonthlyData, userName) {
  const header = ['日付', '最高血圧', '最低血圧', '脈拍'];
  const periods = [
    { name: '1ヶ月', days: 30 },
    { name: '3ヶ月', days: 90 },
    { name: '6ヶ月', days: 180 },
    { name: '1年', days: 365 }
  ];
  const now = new Date();
  
  let sheetIndex = 0;
  periods.forEach(period => {
    ['朝', '夜'].forEach(timeLabel => {
      const sheetName = `${period.name}${timeLabel}`;
      let sheet = getSheet(graphSpreadsheet, sheetName, sheetIndex++); 
      
      const allData = allMonthlyData[timeLabel];
      sheet.clear();
      
      if (allData.length === 0) {
        sheet.getRange(1, 1).setValue('データがありません。');
        return;
      }
      
      const filteredData = allData.filter(row => {
        const rowDate = new Date(row[0]);
        const diffTime = Math.abs(now.getTime() - rowDate.getTime());
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        return diffDays <= period.days;
      }).sort((a, b) => new Date(a[0]) - new Date(b[0]));

      if (filteredData.length > 0) {
        const dataToWrite = [header, ...filteredData];
        sheet.getRange(1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
        
        const averages = calculateAverageData(filteredData);
        displayAveragesOnSheet(sheet, averages);

        createOrUpdateChart(sheet, sheetName, filteredData.length, userName, averages); 
      } else {
        sheet.getRange(1, 1).setValue(`直近${period.name}のデータがありません。`);
      }
      
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 100);
      sheet.autoResizeColumns(2, 4);
    });
  });
}

/**
 * 指定されたシートに時系列グラフを作成または更新する
 */
function createOrUpdateChart(sheet, title, dataRows, userName, averages) {
  const chartRange = sheet.getRange(2, 1, dataRows, 4); 
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  const avgTitle = 
    ` (平均: 最高(赤) ${averages.max} / 最低(青) ${averages.min} / 脈拍(緑) ${averages.pulse})`;

  const chartTitle = 
    `${userName}さんの${title} - 血圧と脈拍の推移${avgTitle} 【作成日: ${today}】`;
  
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  const chart = sheet.newChart()
    .asLineChart()
    .addRange(chartRange)
    .setOption('title', chartTitle)
    .setOption('hAxis.title', '日付')
    .setOption('vAxes.0.title', '血圧 (mmHg)')
    .setOption('vAxes.1.title', '脈拍')
    .setOption('series', {
      0: { targetAxisIndex: 0, color: 'red', label: '最高血圧' }, 
      1: { targetAxisIndex: 0, color: 'blue', label: '最低血圧' }, 
      2: { targetAxisIndex: 1, color: 'green', label: '脈拍' }   
    })
    .setOption('height', 600) //400
    .setOption('width', 600) //900
    .setOption('legend.position', 'bottom')
    .setPosition(6, 6, 0, 0)
    .build();

  sheet.insertChart(chart);
}