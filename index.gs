/**
 * 集計期間や動作設定をまとめた定数
 * @constant
 */
const CONFIG = {
  START: "2025/11/01 00:00:00", // @type {string} 集計開始日時（YYYY/MM/DD HH:mm:ss）
  END:   "2025/12/31 00:59:59", // @type {string} 集計終了日時（YYYY/MM/DD HH:mm:ss）
  IGNORE_COLOR: CalendarApp.EventColor.GRAY, // @type {CalendarApp.EventColor} 無視するカレンダーイベントカラー
  HEADER_BG: "#ebebeb", // @type {string} スプシのヘッダー背景色
  TOTAL_BG: "#5db6ffff", // @type {string} スプシの合計行背景色
  MY_NAME: "E", // @type {string} 担当者未指定時のデフォルト名
  WORKER_NAME: {
    "E": "恵美",
    "M": "素子",
    "B": "バトー",
    "S": "サイトー",
  }
};

/* =====================
 * Value Objects（ドメイン層）
 * 値を保持するだけのクラス群
 * ===================== */
/**
 * 分単位の時間を扱うクラス
 * @param {number} minutes 分単位の時間
 */
class Duration {
  constructor(minutes) {
    this.minutes = minutes;
  }
  /**
   * HH:MM 形式の文字列に変換するメソッド
   * @returns {string}
   */
  toHHMM() {
    const h = Math.floor(this.minutes / 60);
    const m = this.minutes % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  }
}

/**
 * カレンダータイトルを解析した結果を保持するクラス
 * @param {string} workName 物件名
 * @param {string} clientName クライアント名
 * @param {string} task 作業内容
 * @param {string[]} workers 担当者一覧
 */
class ParsedTitle {
  constructor(workName, clientName, task, workers) {
    this.workName = workName;
    this.clientName = clientName;
    this.task = task;
    this.workers = workers;
  }
}

/**
 * 集計対象となるカレンダーイベントを表すクラス
 * @param {Object} params
 * @param {string} params.title 元のカレンダータイトル
 * @param {CalendarApp.EventColor} params.color イベントカラー
 * @param {ParsedTitle} params.parsedTitle パース済みタイトル情報
 * @param {number} params.duration 作業時間（分）
*/
class WorkEvent {
  constructor({ title, color, parsedTitle, duration }) {
    this.title = title;
    this.color = color;
    this.workName = parsedTitle.workName;
    this.clientName = parsedTitle.clientName;
    this.task = parsedTitle.task;
    this.workers = parsedTitle.workers;
    this.duration = duration;
  }
}

/**
 * ピボットテーブル用のシートデータの列数を表すクラス
 * @param {string} name 列名
 */
class PivotDefs {
  constructor() {
    this._columns = {
      '物件名': 1,
      'クライアント名': 2,
      '作業内容': 3,
      '担当者': 4,
      '所要時間（分）': 5,
      '所要時間（時間）': 6,
      // English aliases
      'WORK_NAME': 1,
      'CLIENT': 2,
      'TASK': 3,
      'WORKER': 4,
      'MINUTES': 5,
      'HOURS': 6,
    };
  }

  /**
   * 列名や列番号を受け取り、スプレッドシートの列インデックス（1始まり）を返す
   * @param {string|number} name
   * @returns {number|null}
   */
  columnIndex(name) {
    if (!name) return null;
    return this._columns[name];
  }

  /**
   * summarize指定（例: "SUM","COUNT","AVERAGE"）をSpreadsheetApp.PivotTableSummarizeFunction の値に変換して返す。
   * デフォルトは SUM。
   * @param {string} name
   * @returns {SpreadsheetApp.PivotTableSummarizeFunction}
   */
  toSummarizeFunction(name) {
    if (!name) return SpreadsheetApp.PivotTableSummarizeFunction.SUM; // デフォルト
    const key = String(name).trim().toUpperCase(); // 大文字化して空白削除
    switch (key) {
      case 'SUM':    return SpreadsheetApp.PivotTableSummarizeFunction.SUM;
      case 'COUNT':  return SpreadsheetApp.PivotTableSummarizeFunction.COUNTA;
      case 'AVERAGE':return SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE;
      case 'MAX':    return SpreadsheetApp.PivotTableSummarizeFunction.MAX;
      case 'MIN':    return SpreadsheetApp.PivotTableSummarizeFunction.MIN;
      case 'MEDIAN': return SpreadsheetApp.PivotTableSummarizeFunction.MEDIAN;
      case 'PRODUCT':return SpreadsheetApp.PivotTableSummarizeFunction.PRODUCT;

      // 必要に応じてマッピングを追加
      default:
        console.warn(`PivotDefs: Summarize Function が未定義です: "${name}", SUMを使用します。`);
        return SpreadsheetApp.PivotTableSummarizeFunction.SUM;
    }
  }
}

/* =====================
 * Services
 * 業務ロジックを担当するクラス群
 * ===================== */
/**
 * Googleカレンダーの取得を担当するクラス
 * 指定期間のイベントを取得
 * @param {Date} start 開始日時
 * @param {Date} end 終了日時
 * @returns {GoogleAppsScript.Calendar.CalendarEvent[]}
 */
class CalendarService {
  static getEvents(start, end) {
    const calendar = CalendarApp.getCalendarById(
      Session.getActiveUser().getEmail()
    );
    return calendar.getEvents(start, end);
  }
}

/**
 * カレンダーイベントタイトルを解析するクラス
 */
class EventParser {
  /**
   * タイトル文字列を正規化
   * @param {string} title
   * @returns {string}
   */
  static normalize(title) {
    if (typeof title !== "string") return "";
    return title
      .replace(/\u00A0/g, " ") // NBSP → 半角スペース
      .trim() // 先頭・末尾の空白削除
      .normalize("NFKC"); // 全角記号を正規化（Unicodeの互換性正規化）
  }

  /**
   * タイトルを解析して使いやすいデータに変換
   * 入力形式：【物件名[ | クライアント名]】作業内容[ / 担当者1[・担当者2]]
   * @param {string} title
   * @returns {ParsedTitle|null}
   */
  static parse(title) {
    const normalized = this.normalize(title);
    const regex = /^【(.+?)(?:\s*[|｜]\s*(.+?))?】(.+?)(?:\s*[\/／]\s*(.+))?$/;
    const match = normalized.match(regex);
    if (!match) return null;
    const [, workName, clientNameRaw, task, workerRaw] = match; // match[0]は不要なので省略
    const clientName = clientNameRaw ? clientNameRaw.trim() : ""; // 未指定なら空文字
    const workers = workerRaw
      ? workerRaw.split(/\s*[・,]\s*/) // 「・」または「,」で分割
      : [CONFIG.MY_NAME]; // 担当者未指定時はデフォルト名を使用
    return new ParsedTitle(workName.trim(), clientName, task.trim(), workers);
  }
}

/**
 * ユーティリティクラス（共通処理）
 */
const Utils = {
  /**
   * イベントの開始時刻と終了時刻から所要時間を分単位で返す
   * @param {GoogleAppsScript.Calendar.CalendarEvent} event
   * @returns {number} 分単位の所要時間
   */
  minutesBetween(event) {
    const diffMs = event.getEndTime() - event.getStartTime();
    // 四捨五入して分単位にする（分未満の端数対策）
    return Math.round(diffMs / (60 * 1000));
  },

  /**
   * イベントの解析結果と担当者からイベントキーを作成
   * @param {ParsedTitle} parsed 
   * @param {string} worker 
   * @returns {string}
   */
  makeEventKey(parsed, worker) {
    // タイトルそのものではなく構造化フィールドでキー作成（正規化）
    return `${parsed.workName}||${parsed.clientName}||${parsed.task}||${worker}`;
  }
};

/**
 * スプレッドシートへ書き込むクラス
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
class SheetWriter {
  constructor(sheet) {
    this.sheet = sheet;
  }

  /**
   * ヘッダー行を出力するメソッド
   */
  writeHeader() {
    const HEADERS = [
      "物件名",
      "クライアント名",
      "作業内容",
      "担当者",
      "所要時間（分）",
      "所要時間（時間）",
      "カラー",
    ];

    this.sheet.getRange(1, 1, 1, HEADERS.length) // 開始行、開始列、行数、列数
      .setValues([HEADERS])
      .setBackground(CONFIG.HEADER_BG);
    this.sheet.setColumnWidths(1, 3, 200); // 1〜3列目の幅を200に設定

    // 1行目をヘッダーとしてフィルタを設定
    const lastCol = this.sheet.getLastColumn();
    this.sheet.getRange(1, 1, 1, lastCol).createFilter();
    // ヘッダー行を固定
    this.sheet.setFrozenRows(1);
  }

  /**
   * 項目一覧を出力するメソッド
   * @param {WorkEvent[]} events
   */
  writeEvents(events) {
    if (!events || events.length === 0) return;
    const rows = events.map(e => {
      // 担当者コードを表示名に変換（マッピングがなければ元の値を使う）
      const workerDisplay = (e.workers || []).map(w => CONFIG.WORKER_NAME[w] || w).join(", ");
      return [
        e.workName,
        e.clientName,
        e.task,
        workerDisplay,
        Number(e.duration),           // 分
        Number(e.duration) / 60,      // 時間
        e.color || "",
      ];
    });
    this.sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    this.sheet.getRange(2, 6, rows.length).setNumberFormat("0.00"); // 所要時間（時間）列（6列目）を小数1桁表示にする
  }

  /**
   * 合計行を出力するメソッド
   * @param {WorkEvent[]} events
   */
  writeTotal(events) {
    const row = (events.length || 0) + 2; // ヘッダー行 + データ行 + 1行目
    const totalMinutes = (events || []).reduce((sum, e) => sum + (Number(e.duration) || 0), 0); // 分単位の合計時間
    
    this.sheet.getRange(row, 1, 1, 7) // 開始行、開始列、行数、列数
      .setBackground(CONFIG.TOTAL_BG);
    this.sheet.getRange(row, 1).setValue("合計（時間）");
    this.sheet.getRange(row, 5).setValue(new Duration(totalMinutes).toHHMM());
    this.sheet.getRange(row, 6).setValue(totalMinutes / 60); // 時間に変換
    this.sheet.getRange(row, 6, 1, 1).setNumberFormat("0.00"); // 小数点2桁表示
  }

  /**
   * 円グラフを描画
   * @param {number} rowCount
   */
  // drawChart(rowCount) {
  //   if (!rowCount || rowCount < 1) return;
  //   const chart = this.sheet.newChart()
  //     .addRange(this.sheet.getRange(`A1:B${rowCount + 1}`))
  //     .setChartType(Charts.ChartType.PIE)
  //     .setOption("pieSliceText", "value")
  //     .setPosition(1, 9, 0, 0)
  //     .build();
  //   this.sheet.insertChart(chart);
  // }
}

/**
 * 宣言的にピボットを定義するオブジェクト
 * @constant
 */
const PIVOT_DEFINITION = {
  rows:    ["物件名"], // 行グループにする列名
  columns: ["担当者"], // 列グループにする列名
  values: [
    { column: "所要時間（時間）", summarize: "SUM", displayName: "合計時間" }
  ],
};

/**
 * 集計結果をもとにピボットテーブルを作成するクラス
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {PivotDefs} defs
 */
class PivotTableBuilder {
  constructor(spreadsheet, dataSheet, defs = new PivotDefs()) {
    this.spreadsheet = spreadsheet;
    this.dataSheet = dataSheet;
    this.defs = defs;
  }

  // ピボットテーブルの基本構造を作成
  _createPivotSheet(name) {
    const old = this.spreadsheet.getSheetByName(name);
    if (old) this.spreadsheet.deleteSheet(old);
    const sheet = this.spreadsheet.insertSheet(name);
    const pivot = sheet.getRange('A1').createPivotTable(this.dataSheet.getDataRange());
    return { sheet, pivot };
  }

  /**
   * ピボットテーブルを作成するメソッド
   * @param {Object} def
   * @param {string[]} def.rows 行グループの列名一覧
   * @param {string[]} def.columns 列グループの列名一覧
   * @param {Object[]} def.values 値フィールドの定義一覧
   * @param {string} def.values[].column 列名
   * @param {string} def.values[].summarize 集計方法（例: "SUM","COUNT"）
   */
  create(definition, pivotSheetName = `Pivot_${this.dataSheet.getName()}`) {
    const { pivot } = this._createPivotSheet(pivotSheetName);

    // 宣言的に row, column, value を設定
    (definition.rows || []).forEach(name => 
      pivot.addRowGroup(this.defs.columnIndex(name))
    );
    (definition.columns || []).forEach(name =>
      pivot.addColumnGroup(this.defs.columnIndex(name))
    );
    (definition.values || []).forEach(v =>
      pivot.addPivotValue(
        this.defs.columnIndex(v.column),
        this.defs.toSummarizeFunction(v.summarize)
      ).setDisplayName(v.displayName || v.column)
    );
  }
}

/* =====================
 * Orchestrator
 * 全体の流れを制御する
 * ===================== */
/**
 * Googleカレンダーの予定を取得し、タイトルを解析して工数データとしてスプレッドシートに出力する。
 * 1. カレンダーイベントを取得
 * 2. タイトル形式を解析
 * 3. 作業者ごとに工数を集計
 * 4. スプレッドシートに書き込み
 */
function CalendarToSpreadSheet() {
  const startDate = CONFIG.START ? new Date(CONFIG.START) : new Date();
  const endDate = CONFIG.END ? new Date(CONFIG.END) : new Date();

  // シート名指定
  const sheetName =
    `${startDate.getFullYear()}/${startDate.getMonth() + 1}/${startDate.getDate()}`
    + "-" +
    `${endDate.getFullYear()}/${endDate.getMonth() + 1}/${endDate.getDate() - 1}`;

  // スプレッドシートの準備
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // アクティブなスプレッドシートを取得
  const old = ss.getSheetByName(sheetName); // 既存のシートを取得
  if (old) ss.deleteSheet(old); // 既存のシートがあれば削除
  const sheet = ss.insertSheet(sheetName); // 新しいシートを作成

  // 1. カレンダーイベントを取得
  const rawEvents = CalendarService.getEvents(startDate, endDate);

  const eventMap = new Map(); // key: workName||client||task||worker -> WorkEvent
  // カレンダーの予定ごとに処理
  for (const ev of rawEvents) {
    const color = ev.getColor ? ev.getColor() : null;
    // 無視するカラーのイベントはスキップ
    if (color && CONFIG.IGNORE_COLOR && color === CONFIG.IGNORE_COLOR) continue; 

    // 2. タイトル形式を解析
    const title = ev.getTitle();
    if (!title) continue;
    const parsed = EventParser.parse(title); // カレンダータイトルを解析して使いやすいデータに変換
    if (!parsed) {
      Logger.log(`タイトル不正: ${title}`);
      continue;
    }
    const minutes = Utils.minutesBetween(ev); // イベントの所要時間を分単位で取得
    if (minutes <= 0) continue;

    // 3. 作業者ごとに工数を集計
    for (const worker of parsed.workers) {
      const key = Utils.makeEventKey(parsed, worker); // イベントキーを作成
      // 既存のWorkEventがあれば所要時間を加算、なければ新規作成してMapに追加
      if (eventMap.has(key)) {
        eventMap.get(key).duration += minutes;
      } else {
        // 新規作成
        const singleParsed = new ParsedTitle(parsed.workName, parsed.clientName, parsed.task, [worker]); // 作業者は1人ずつ保持
        eventMap.set(key, new WorkEvent({
            title,
            color,
            parsedTitle: singleParsed,
            duration: minutes,
          })
        );
      }
    }
  }
  const events = [...eventMap.values()]; // Mapの値を配列に変換

  // 4. スプレッドシートに書き込み
  // 物件名→クライアント→作業→担当者 でソート
  events.sort((a, b) => {
    const ka = `${a.workName}|${a.clientName}|${a.task}|${a.workers.join(",")}`;
    const kb = `${b.workName}|${b.clientName}|${b.task}|${b.workers.join(",")}`;
    return ka.localeCompare(kb);
  });
  Logger.log(`イベント件数: ${events.length}`);
  const writer = new SheetWriter(sheet); // シート書き込み
  writer.writeHeader();
  writer.writeEvents(events);
  writer.writeTotal(events);

  // ピボットテーブル作成（定義を渡す）
  new PivotTableBuilder(ss, sheet).create({
    rows:    ["物件名"],
    columns: ["担当者"],
    values:  [{ column: "所要時間（時間）", summarize: "SUM" }],
  });
  // writer.drawChart(events.length); //チャート描画（WIP）
}
