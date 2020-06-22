DIVIDER = '-';

// 全ての単語を一個ずつ処理
function all(){
  const SHEET = SpreadsheetApp.getActiveSheet(); //シートを取得
  var active_cell = SHEET.getActiveCell(); //アクティブセルを取得

  while(active_cell.getValue() != ''){
    myFunction();
    active_cell.offset(1, 0).activate();
    active_cell = SHEET.getActiveCell();
  }
}

function myFunction() {

  const SHEET = SpreadsheetApp.getActiveSheet(); //シートを取得
  const ACTIVE_CELL = SHEET.getActiveCell(); //アクティブセルを取得

  // 1列目以外は翻訳しない
  if (ACTIVE_CELL.getColumn() !== 1.0) {
    return false;
  }

  const TARGET = SHEET.getRange(ACTIVE_CELL.getRow(), ACTIVE_CELL.getColumn()).getValue().trim();

  // dividerは翻訳しない
  if (TARGET === DIVIDER) {
    return false;
  }

  const EXAMPLE_SENTENCES = getExampleSentence(TARGET);

  const JAPANESE_MEANING = searchWord(TARGET);
  ACTIVE_CELL.offset(0, 6).setValue(JAPANESE_MEANING[0] + "\n\n" + EXAMPLE_SENTENCES);
}

function getExampleSentence(word) {
  // 空文字なら削除
  if (word === '') {
    return '';
  }

  const WEBLIO_URL = 'https://ejje.weblio.jp/sentence/content/';
  const HTML = UrlFetchApp.fetch(WEBLIO_URL + word).getContentText();

  //Logger.log(HTML);

  // 英語の例文を取得
  var example_sentences = '';
  var english_list = Parser.data(HTML).from('<p class=qotCE>').to('</p>').iterate();
  var japanese_list = Parser.data(HTML).from('<p class=qotCJ>').to('<span>').iterate();
  for ( var i=0; i<english_list.length; i++ ){
    var english = english_list[i].replace(/(<([^>]+)>)/ig,"").replace('例文帳に追加', '');
    var japanese = japanese_list[i].replace(/(<([^>]+)>)/ig,"").replace('.', '');
    example_sentences += english + '. ' + japanese + ";\n";
  }

  return example_sentences;
}


// 単語を検索
function searchWord(word) {
  // 空文字なら翻訳を削除
  if (word === '') {
    return ['', ''];
  }

  const WEBLIO_URL = 'http://ejje.weblio.jp/content/';
  const HTML = UrlFetchApp.fetch(WEBLIO_URL + word).getContentText();

  //Browser.msgBox(getAlk(word));
  //Logger.log(HTML);

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

  const TRANSLATED = (PART + meaning).replace('【意味】', '');
  return [TRANSLATED, SOUND_URL]
}
