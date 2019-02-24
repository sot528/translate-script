DIVIDER = '-';

function myFunction() {

  const SHEET = SpreadsheetApp.getActiveSheet(); //シートを取得
  const ACTIVE_CELL = SHEET.getActiveCell(); //アクティブセルを取得

  // 1列目以外は翻訳しない
  if (ACTIVE_CELL.getColumn() !== 1.0) {
    return false;
  }

  // 1シート目以外は翻訳しない
  if (SHEET.getIndex() !== 1.0) {
    return false;
  }

  const TARGET = SHEET.getRange(ACTIVE_CELL.getRow(), ACTIVE_CELL.getColumn()).getValue().trim();

  // dividerは翻訳しない
  if (TARGET === DIVIDER) {
    return false;
  }

  // 既存の単語は翻訳しない
  const exists = is_exists(SHEET, TARGET);
  if (TARGET !== '' && exists) {
    ACTIVE_CELL.offset(0, 1).setValue(exists + '\n' + TARGET + ': ALREADY EXISTS.');
    ACTIVE_CELL.offset(0, 0).setValue('');
    return false;
  }

  const WORD_INFO = searchWord(TARGET);

  // 変換できなかった単語はGoogleの検索リンクを貼る
  if (WORD_INFO[0].match('weblio辞書で英語学習')) {
    ACTIVE_CELL.offset(0, 1).setValue('https://www.google.com/search?q='
        + TARGET.replace(/\s/g, '+'))
        + '&source=lnt&tbs=lr:lang_1ja&lr=lang_ja';
    return false;
  }

  ACTIVE_CELL.offset(0, 1).setValue(WORD_INFO[0]);
  ACTIVE_CELL.offset(0, 3).setValue(WORD_INFO[1]);
}

// 単語を検索
function searchWord(word) {
  // 空文字なら翻訳を削除
  if (word === '') {
    return ['', ''];
  }

  const WEBLIO_URL = 'http://ejje.weblio.jp/content/';
  try {
    const HTML = UrlFetchApp.fetch(WEBLIO_URL + word).getContentText();
  } catch (e) {
    Logger.log('word:' + word + ', error:' + e.message);
    return ['', '']
  }

  //Browser.msgBox(getAlk(word));
  //Logger.log('search: ' + getAlk(word));

  // 品詞を取得
  const TMP = HTML.match(/<meta\ name="description"\ content="([\s\S]*?)">/i);
  const PART = TMP[1].match(/【([\s\S]*?)】/i) ? TMP[1].match(/【([\s\S]*?)】/i)[0].trim() : '';

  // 音声を取得
  const TMP_SOUND = HTML.match(/https:\/\/([^:]*?)\.mp3/i);
  const SOUND_URL = TMP_SOUND ? TMP_SOUND[0].trim() : '';

  // 正規表現が異なる？ ため意味が取れず。twitter用のデータから抽出
  var meaning = HTML.match(/<meta\ name="twitter:description"\ content="([\s\S]*?)">/i);
  meaning = meaning[1].replace(word + ':', '')
      .replace('&lt;b&gt;', '').replace('&lt;/b&gt;', '').trim();

  // 動詞以外の意味の場合は動詞も付与
  const VERB = PART === '【動詞】' ? '' : getVerbFromAlc(word);

  const TRANSLATED = PART + meaning + VERB;
  return [TRANSLATED, SOUND_URL]
}

// アルクの辞書から取得。動詞以外の意味の言葉に動詞を付与するために利用している
// Weblioではいい感じに動詞が取得できないため
function getVerbFromAlc(word) {
  const URL_ALC = 'https://eow.alc.co.jp/search?q=';
  try {
    const html = UrlFetchApp.fetch(URL_ALC + word).getContentText();
  } catch (e) {
    Logger.log('word:' + word + ', error:' + e.message);
    return '';
  }

  var alc = html.match(/<meta\ name="description"\ content="([\s\S]*?)\ -\ /i);
  
  alc = !alc ? '' : "\n\n" + alc[1].replace(word + ' ', '')
      .replace('&lt;b&gt;', '').replace('&lt;/b&gt;', '').trim();
  alc = alc.match(/[動詞|自動|他動]+/i) ? alc : '';

  return alc;
}

// すでにシートに存在するか確認
function is_exists(sheet, val) {
  const DAT = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  const LOWERCASE_VAL = val.toLowerCase();
  var counter = 0;
  var exsistingTranslation = '';
  for (var i = 1; i < DAT.length; i++) {
    if (DAT[i][0].toLowerCase() === LOWERCASE_VAL) {
      ++counter;
      if (counter >= 2) {
        return exsistingTranslation;
      }else{
        exsistingTranslation = DAT[i][1];
      }
    }
  }
  return false;
}


// ------------------------------------------------------------


// Kindleから単語を追加
function addWordsFromKindle() {
  var cell;
  const WORDS = getWordsFromGmailSendByKindle();
  for (var i in WORDS) {
    cell = getCellForNewWord();
    cell.activate();
    cell.setValue(WORDS[i]);
    main();
  }
  clear();
}

// Kindleから送られたメールの添付CSVから単語を抽出し、その後メールは削除
function getWordsFromGmailSendByKindle() {
  const QUERY = 'のメモ label:inbox from:no-reply@amazon.com';
  const WORDS_START_ROW = 8;
  const WORDS_COLUMN = 3;
  const THREADS = GmailApp.search(QUERY);

  var words = [];

  for (var numThread in THREADS) {
    //添付CSV読み込み
    var csv = Utilities.parseCsv(THREADS[numThread].getMessages()[0].getAttachments()[0].getDataAsString());

    //CSVから単語を抜き出す
    for (var i = WORDS_START_ROW; i < csv.length; i++) {
      words.push(csv[i][WORDS_COLUMN]);
    }

    // スレッドにいくつもメッセージが溜まって面倒なのでアーカイブではなく削除する
    THREADS[numThread].moveToTrash();
  }

  return words;
}

// 次の単語が入るべきセルを取得
function getCellForNewWord() {
  const sheet = SpreadsheetApp.getActiveSheet();
  return sheet.getRange(sheet.getLastRow() + 1, 1);
}


// ------------------------------------------------------------


// シートを綺麗にする
function clear() {
  const SHEET = SpreadsheetApp.getActiveSheet();
  const DATA_RANGE_VALUES = SHEET.getDataRange().getValues();
  const UNNECESSARY_ROWS = [];
  var rowIndex;
  var cell;


  for (var i = 0; DATA_RANGE_VALUES.length > i; i++) {
    rowIndex = i + 1;

    // 不要な行は最後にすべて除却
    if (DATA_RANGE_VALUES[i][0] === DIVIDER || DATA_RANGE_VALUES[i][0] === '') {
      UNNECESSARY_ROWS.push(rowIndex);
      continue;
    }

    // 翻訳されていないものがあれば翻訳
    if (DATA_RANGE_VALUES[i][1] === '') {
      cell = SHEET.getRange(rowIndex, 1);
      cell.activate();
      myFunction();
    }
  }

  // 降順ソート
  UNNECESSARY_ROWS.sort(function (a, b) {
    if (a > b) return -1;
    if (a < b) return 1;
    return 0;
  });

  // 不必要な行は削除
  for (var j = 0; UNNECESSARY_ROWS.length > j; j++) {
    SHEET.deleteRow(UNNECESSARY_ROWS[j]);
  }
}
