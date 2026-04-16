//😃😃基本😃😃基本😃😃基本😃😃基本😃😃基本😃😃基本
//😃😃😃😃設定項目😃😃😃😃設定項目😃😃😃😃設定項目
//😃😃権限判定
function kenGenJudge(param1) {
  const systemManager = "丸山達也";
  // personal に名前を追加するだけで OK
  const personal = [//経費用
    "石井利樹",
    "石井亜也子",
    "浅谷美保",
    "望月凜花"
    
    // ここに増やしていけます
  ];
  if (param1 === systemManager) {
    return "systemManager";
  }
  if (personal.includes(param1)) {
    return "personal";
  }
  return "none";  // どちらにも当てはまらない場合
}

//収支用権限メンバー
function syusiKengen(syain){
  const personal2 = [//収支用
    "石井利樹",
    "石井亜也子",
    "浅谷美保",
    "丸山達也",
    "石井亜也子",
    "有馬和嘉子",
    "松原情恵",
    "テスト担当",
    "宮原亜美",
    "望月凜花"
    // ここに増やしていけます
  ];
  if(personal2.includes(syain)){
    return "OK";
  }
}
function 社名(){
  return "石井工業";
}

function 会社名(){
  return "（株）石井工業";
}
function 正式名称(){
  return "株式会社石井工業";
}
//😃😃メールでの問い合わせ（システム管理）
function 管理者(){
  return "丸山達也"
}
//管理者メール
function 管理Mail(e){
  if(!e){
    return "maruyama@ebisu-ishii.co.jp";
  }else if(e == "経理"){
    return "ishii.ayako@ebisu-ishii.co.jp";
  }else if(e == "社長"){
    return "toshiki@ebisu-ishii.co.jp";
  }else if(e == "システム"){
    return "system.ishii@ebisu-ishii.co.jp";
  }else if(e == "契約請求"){//契約請求😃😃迷惑メールに入る問題がある
    return "keiyaku-seikyu@ebisu-ishii.co.jp";
  }else if("事務"){
    return "asatani@ebisu-ishii.co.jp";//南澤さん
  }
}
//管理電話
function 管理Tel(){
  return "070-6456-8882";
}
function 社長(){
  return "石井利樹";
}
function 勤怠チャットpig(){
  return "https://drive.google.com/file/d/1nmkKSxX8bXCWUQhPIe-KPJ2cEZRhyiuW/view?usp=sharing"//勤怠チャットのお疲れ様につける画像　お疲れ様時コンパネ素材2（共有設定が必要）
}

//😃😃スプレットシートID
function getSPId(e){//eの値によってスプレットシートのIDを返す
  if(e == "業者顧客担当"){
    return "1v_N5rJZOyaUf4_W8xR1itWNeIj_fug3AR48m0HIW-ME";
  }else if(e == "年度フォルダ"){
    return "1yHLjRhNNS-Ht22JltcfxBb8G6_YaEjHTxI-ExL58ZcI";
  }else if(e == "勤怠発行管理"){
    return "1E_ZxoRt7nxChEI6JfVbOm4dh6JZVbXEOoTYgscO2F3E";
  }else if(e == "log記録"){
    return "18Kx7VVfvxU9HVaL9Dc1JbDvJQYhXyxEBpx3fdQ20KzE";
  }else if(e == "休日カレンダー年間計画"){
    return "1ob9jVO0F2VwKlg7vr6Px8HMsoJrLM1aMbg6Zt97sjwc";//【管理】勤怠管理の休日カレンダー年間計画
  }else if(e == "社員用ひな形"){
    return "1WU27lBHBsR4uCAGozmauA8z43nFhXxFI7gngJF1IYm0";
  }else if(e == "支払案内"){
    return "1g6LtwtEFJCIZTfXdqsI_H04RFmFBdTIlM8vQOOQ9_28"//【管理】請求支払→【ひな形】フォルダ
  }else if(e == "経費管理"){
    return "1gXa-0L4HBQ1UhPvDHOdkcdozJlhzA-iIW9w7py5sj8E";//【管理】経費管理の経費管理（発行管理）
  }else if(e == "長谷工完了報告"){
    return "1lsDhEOgNFmB6jMCYo3laGR4IXPObUp5TtesCamIe1Pg";//【管理】請求支払の【管理表長谷工】工事完了報告書
  }else if(e == "QUEUE"){
    return "1XSPkDGJWhcS-1BkPVDxKEYjRdkOpTzgSyYE6xervn5Q";//松田塗装【管理】請求支払→【契約署名】開発フォアルダの（QUEUE）契約署名記録
  }
}
//😃😃スプレットシートのURLを取得する
function getSSUrl(e){//🔷完了🔷
  if(e == "大成内訳"){
    return "https://docs.google.com/spreadsheets/d/10Jzzj926wnNi41U4LMv9bntiCa7gw-w4Bw5mzHoBiyc/edit?gid=848455356#gid=848455356";//松田塗装【管理】請求支払→【ひな形】ファイルの【ひな形役務】大成請求ファイルの様式B（内訳）シートが開くURL
  }
}
//😃😃ひな形ファイルID取得
function ひな形(e){
  if(e == "長谷工完了報告"){
    return "1lsDhEOgNFmB6jMCYo3laGR4IXPObUp5TtesCamIe1Pg";
  }else if(e =="自社請求"){
    return "1jkzG80XyuOVAcBEsBLpJTwUPZ6kQ40OhXRUiVME-pzM";
  }else if(e == "予算書雛形"){
    return "11lGcIPD6L7QMeNRABPMD4bMKXH2gLaacHpGmOFU_dKQ";
  }else if(e == "明細集計用"){
    return "1hN0qxVDt5GNoZ7mplKsNc2wwNA7EoGCHr0t0RNkDelQ";
  }else if(e == "出面明細"){
    return "1Y8K0sXeG5AHxp2L17X6iL2rHgM7CB6n-wtPtTadAIBk";
  }else if(e == "支払い案内"){
    return "1g6LtwtEFJCIZTfXdqsI_H04RFmFBdTIlM8vQOOQ9_28";//【管理】請求支払→【ひな形】ファイル
  }else if(e == "業者別支払リスト"){
    return "1SAW8daKnzVScUpFdE8H_1JIRJ-HuQkYUUHkMpt5dE3k";//【管理】請求支払
  }else if(e == "予算書"){
    return "11lGcIPD6L7QMeNRABPMD4bMKXH2gLaacHpGmOFU_dKQ";//★管理★フォルダ→【ひな形】フォルダ
  }else if(e == "月次請求リスト"){
    return "1gzkCBjBChuxCed-CmBx21rilr2bc3l2C_jRji2QmvlI";//【管理】請求支払→【ひな形】ファイルの【ひな形】月次業者請求まとめ用
  }else if(e == "現場担当確認リスト"){
    return "1jV0lROy7rqLWCpwzaAIdY7oT77mp_gaI06dkEs2hTnc";//【管理】請求支払い→【ひな形】フォルダ→現場別担当者確認リストss
  }else if(e == "請求書"){
    return "1kjcXqXgNW06TddGnyVR8MdRDjra9ezkYob5iVwol3gs";//【管理】請求支払→【ひな形】フォルダ→請求書ひな形（出来高用）
  }else if(e == "協力会"){
    return "1ViuY0uvQZNoCaU1C5qAkoqk-FcWWlmsefK2X6DFhEbM";//【管理】請求支払→◇業者管理◇（保存)の協力会の案内
  }else if(e == "覚書"){
    return "1ESE-Hy-gqFkWhQRDyhThHDPubPr18MjY7XuQ300TwXQ";//【管理】請求支払→◇業者管理◇（保存)の覚書ひな形（請求受付説明付）
  }else if(e == "社員毎月ひな型"){
    return "1WU27lBHBsR4uCAGozmauA8z43nFhXxFI7gngJF1IYm0";//【管理】勤怠管理→ひな形フォルダの社員（毎月）ひな形
  }else if(e == "勤怠管理用"){
    return "1Aeidke-ArPhZEWHSQocObNkkJsUaewfpCKNL0R6plPw";//【管理】勤怠管理→ひな形フォルダの管理用（毎月）
  }else if(e == "経費管理ひな形" ){
    return "1Kh93RZN2jwDEhAhprP2Jfp3xI0ABd9riQ8nU4dom3YI";//【管理】経費管理→ひな形フォルダの【ひな形】経費管理用（毎月） 
  }else if(e == "経費集計PDF用"){
    return "1pdGD7IgE-zWCHTnK8kQpmCzsDUBLN4r76pglsSrutPo";//【管理】経費管理→ひな形フォルダの【ひな形】集計表PDF用
  }else if(e == "社員用経費精算"){
    return "1u86kIX6tk_aNrmpGOSGM2NROAjtOqNTcci0FV65ebtY";//【管理】経費管理→ひな形フォルダの【ひな形】社員用経費精算書
  }else if(e == "経費管理SS"){
    return "1XmNS2nwwRFT6OzmxTVI8-fMgc1XlQZ3NaDXEzpvaSP4";//【管理】経費管理→ひな形フォルダの【ひな形】経費管理SS
  }else if(e == "業者支払台帳"){
    return "1_pgflNKp3Y2xPaASkziABt8_B-Sxv2BAcSun2FjETQ8";//【管理】請求支払→業者支払台帳の【ひな形】★業者支払い台帳
  }else if(e == "休日カレンダー"){
    return "1qSKtNclc0rFget7e4-cuXgtM07_Nkyptba_2u7guM8M";//【管理】勤怠管理→年間休日カレンダーの【ひな形】休日カレンダー年間計画
  }else if(e == "月まとめ支払表"){
    return "1vny1k3GGMcepSZJefNHnNKQiNHAF3btDsSl8gP0eOOM";//【 管理】請求支払→【ひな形】ファイルの【ひな形】月まとめ支払い表（現場別）
  }else if(e == "業者注文請書"){
    return "16qJi4_EiTLniQ2-GzgoOjTVlNyNgoEI10j_xj3nTFUk";//★管理フォルダ★→ひな形フォルダの【ひな形】業者注文請書.xlsm のコピー
  }else if(e == "注文請書単独"){
    return "1oqiV3uNxXL_oJI5YCZXJizF_fz00utUJ9Bo63NfyAFg";//★管理フォルダ★→ひな形フォルダの【ひな形】業者注文請書（単発）.xlsm  のコピー 
  }else if(e == "大成工事請求"){
    return "1IcelXr4N_abjU8y_FwHcu1eWdWZS9NqiSJx5U7BM12Y";//【管理】請求支払→【ひな形】ファイルの【ひな形工事】大成請求 のコピー
  }else if(e == "大成役務請求"){
    return "10Jzzj926wnNi41U4LMv9bntiCa7gw-w4Bw5mzHoBiyc";//【管理】請求支払→【ひな形】ファイルの【ひな形役務】大成請求 のコピー
  }else if(e == "石井請負注文"){
    return "1a-jbxunpTyc8p61zxU3bMInEo3wZvB7C0MaWf4a0rvg";//★管理フォルダ★→ひな形フォルダの【ひな形】石井請負注文請書 のコピー
  }else if(e == "CSV変換用"){
    return "1E8q4zf61-MkViQyyfE6itBbJI1BA3QaRMliV_hfOEns";//★管理フォルダ★→ひな形フォルダのCSV変換用 のコピー
  }else if(e == "SS担当出来高"){
    return "1aDgOvF1p3RAQ6JiqxDPdjq1dYPx1HPVZIEihzCVhOHk";//【管理】請求支払→【ひな形】ファイルのSS担当者出来高反映用
  }
}
//◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆
function getSetUrl(e){
  //◆デプロイを管理のURLを↓↓に
  const urlF = () => {
    if(!e){
      return "https://script.google.com/a/macros/s/AKfycbz3Sgm0A3gRIl---r1qSYAkE65ZNo0wbGP08YLw4wc_8NTEyKD3ydq5hIspxiZ-U8qHmQ/exec"//工事台帳データベース（システム）
    }else if(e == "業者新規Url"){
      return "https://script.google.com/a/macros/s/AKfycbxEmV92UMGhQDlm370j76QvTHbWgwJRnn2M9uzFoBfFnnuZE96BFocAyrzhpxYg4iCM/exec";//【WEBアプリ】会社情報（新規）
    }else if(e == "担当者出来高"){
      return "https://script.google.com/a/macros/s/AKfycbxiwcII0NhX2b72BLYuDj7FC_hVZadoTzYxqAQGXJrmvtAJ8iA3-O-zJgx22hKv_LLK/exec";//【webｱﾌﾟﾘ】担当者用（出来高処理）
    }else if(e == "勤怠アプリ"){
      return "https://script.google.com/a/macros/s/AKfycbzZGziiTY9hA-UaS2VnhWHlZFGacUNjTciykl8TYo05PEyKqd3Juct5aUt0rFw1ubahIQ/exec";//勤怠管理【webアプリ2024】
    }else if(e == "勤怠情報"){
      return "https://script.google.com/a/macros/s/AKfycbwwv9uRS8n1gPdj-tv6G2DmnoXsSfJqRJuai-2rZu0cJUE8gJs8N_suSWeQJjDqsflFQQ/exec";//【webアプリ】勤怠アプリ情報確認パネル
    }else if(e == "業者請求出来高"){
      return "https://script.google.com/a/macros/s/AKfycbxH5YToR9FOU6TqUJh7lCYx-WWm3tv4q9_UPXsDi78lAdiR_ltFJHuRd1QAubxbpUwpUQ/exec";//【webｱﾌﾟﾘ】業者請求用（出来高請求用）
    }else if(e == "会社情報年次"){
      return "https://script.google.com/a/macros/s/AKfycbx7LBtQbDr5IifIDkMvHRIvuP8x4syOYHABy6BQLklU16O5z8llgiZEaYZXps9AwY16vQ/exec"//【webアプリ】会社情報（年次）
    }else if(e == "経費アプリ"){
      return "https://script.google.com/a/macros/s/AKfycbzam1Nl8d1T4VolI5vflZy459xc9LEA_bB50GneO3eicEV2YEDpX2WL6oVYq0nXHnOY/exec";//経費アプリUrl
    }else if(e == "予算書アプリ"){
      return "https://script.google.com/a/macros/s/AKfycbxYYDbl2CSy6BTeAO-jTEdBbNvBclOGpuDNEzAQgwlpt8ovS1gE6WiZzkuhtY4M3ng/exec";//予算書アプリ（案内発注）
    }else if(e == "収支アプリ"){
      return "https://script.google.com/macros/s/AKfycbxkE4wqOe0SoOQRaRYsWm6dUGt42r7AnhPEy-dWlbFMf_SI_hBpnVm1CCYaSmkcenfD/exec";//収支管理フォルダ【webアプリ】
    }else if(e == "契約書署名"){//🔷OK🔷
      return "https://script.google.com/macros/s/AKfycbz6d_t0ZmOv1qaGZ7HtPBVD7gZfjPobMl5XPqIEFke7XR3XL96GlHveMZnDT6kKcumeNQ/exec";//契約書署名【アプリ】
    }else if(e == "新社員情報"){
      return "https://script.google.com/macros/s/AKfycby8RgKMiCiXFypAFj3rV0Wv60XFbXDQrVd0E_HHuaKODecDwDOYHz0qTVUaG-25WIeC/exec";//新社員情報確認【webApp】
    }
  }
  const url = urlF();
  a = url.slice(0,26);
  b = url.slice(32);
  c = "a/macros/ebisu-ishii.co.jp";
  urlA = a + c + b;
  Logger.log(urlA);
  return urlA;
}
//◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆URLを指定する◆

//😃追加😃
function 台帳SetUrl2() {
  const domain = 'ebisu-ishii.co.jp';

  let yosan  = 'https://script.google.com/a/macros/s/AKfycbxYYDbl2CSy6BTeAO-jTEdBbNvBclOGpuDNEzAQgwlpt8ovS1gE6WiZzkuhtY4M3ng/exec'; // 予算
  let syusi  = 'https://script.google.com/macros/s/AKfycbxkE4wqOe0SoOQRaRYsWm6dUGt42r7AnhPEy-dWlbFMf_SI_hBpnVm1CCYaSmkcenfD/exec'; // 収支
  let keihi  = 'https://script.google.com/a/macros/s/AKfycbzam1Nl8d1T4VolI5vflZy459xc9LEA_bB50GneO3eicEV2YEDpX2WL6oVYq0nXHnOY/exec'; // 経費
  let kinzyo = 'https://script.google.com/a/macros/s/AKfycbwwv9uRS8n1gPdj-tv6G2DmnoXsSfJqRJuai-2rZu0cJUE8gJs8N_suSWeQJjDqsflFQQ/exec'; // 勤怠情報

  const rn = [yosan, syusi, keihi, kinzyo];

  const fix = (u) => {
    const base = 'https://script.google.com/';
    if (!u.startsWith(base)) return u;
    const rest = u.slice(base.length); // 例: "macros/..." or "a/macros/..." or "a/<domain>/macros/..."
    if (rest.startsWith('macros/')) {
      // ドメインなし → 付与
      return `${base}a/${domain}/${rest}`;
    }
    if (rest.startsWith('a/')) {
      const parts = rest.split('/'); // ["a", "macros" or "<domain>", ...]
      if (parts[1] === 'macros') {
        // 順序が誤り: /a/macros/... → /a/<domain>/macros/...
        parts.splice(1, 0, domain); // ["a", "<domain>", "macros", ...]
        return base + parts.join('/');
      }
      // 既に /a/<何か>/macros/... → ドメインを強制的に置換
      parts[1] = domain;
      return base + parts.join('/');
    }
    return u;
  };

  for (let i = 0; i < rn.length; i++) {
    rn[i] = fix(rn[i]);
  }
  return rn;
}


//😃😃チャット送信先
function chatUrl(e){
  if(e == "管理1"){
    return "https://chat.googleapis.com/v1/spaces/7XSXnkAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=UaDx9eyXbms4EwHrcz5lY_TXuTzLd66kr6drC7Yq5PE"
      //丸山url
  }else if(e == "管理2"){
    return "https://chat.googleapis.com/v1/spaces/AAAA1s3kXVw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=JpoMK1-_4CiHEd3SIIVaFzrNfmh7SZtng20cbBOOg0k";//アプリメッセージ
  }else if(e == "経理"){
    return "https://chat.googleapis.com/v1/spaces/k7nYeUAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=YgIpgPYDtDppe83YwFyAMLsK3T7nnMwmJNwV63HQVdU";//石井亜也子
  }else if(e == "社長"){
    return "https://chat.googleapis.com/v1/spaces/gL04eUAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cPmZdv3k4QAQHAluOKA0lTih_-1YMz90ZFncsZWph5A";//社長
  }else if(e == "部長"){
    return "https://chat.googleapis.com/v1/spaces/uqKYeUAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=oFFAyBXNzqZ-o1kt4l0LH7rX2D9yFcLwdTbiUEetvdI"//亀卦川
  }else if(e == "全体丸山"){
    return "https://chat.googleapis.com/v1/spaces/AAAALAFDA04/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=yDyK46DIs8ghea25T2HQACI2CFdzn-SdOzNt2QGaSA0";
  }else if(e == "注文書丸山"){
    return "https://chat.googleapis.com/v1/spaces/AAAAqlbD8WU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=IX7NMpJTHgRdERDEUf6gjTf1NHd8Q3v8U4dsVrdsI4o";
  }
}

//😃😃メールアドレス
function メアド(e){
  if(e == "管理"){
    return "maruyama@ebisu-ishii.co.jp";
  }else if(e == "社長"){
    return "toshiki@ebisu-ishii.co.jp";
  }else if(e == "経理"){
    return "ishii.ayako@ebisu-ishii.co.jp";
  }else if(e == "請求支払"){
    return "keiyaku-seikyu@ebisu-ishii.co.jp";
  }else if(e == "システム"){
    return "system.ishii@ebisu-ishii.co.jp"
  }
}

//😃😃メール送信時の署名
//😃😃メール送信時の署名
function メール署名(e){
  if(!e){
    return "◇◇◇◇◇◇◇◇◇◇\n株式会社石井工業\n契約請求\n丸山達也\n070-6456-8882\nmaruyama@ebisu-ishii.co.jp\n◇◇◇◇◇◇◇◇◇◇";
  }else if(e == 2){
    return `◇◇◇◇◇◇◇◇◇◇
    株式会社　石井工業
    契約請求
    丸山達也
    maruyama@ebisu-ishii.co.jp
    070-6456-8882
    ◇◇◇◇◇◇◇◇◇◇
    `;
  }else if(e == 3){
    return "<p>◇◇◇◇◇◇◇◇◇◇</p><p>株式会社石井工業</p><p>契約請求</p><p>丸山達也</p><p>070-6456-8882</p><p>maruyama@ebisu-ishii.co.jp</p><p>◇◇◇◇◇◇◇◇◇◇</P>"
  }
}

//😃😃ドライブのIDを取得する
function ドライブId(e){
  if(e == "注文請書"){
    return "0AAfqjEosHAYcUk9PVA";
  }else if(e == "年間カレンダー"){//【管理】勤怠管理の年間カレンダーの最新フォルダ
    return "1G99vfn7EpgNf7-vlUs46qq5Kzw06tNay";
  }else if(e == "年間休日カレンダー"){
    return "1bkFY5i0fL8CTq02NfF3lnPp39M22tzAq"//【管理】勤怠管理の年間カレンダー
  }else if(e == "建設業許可"){
    return "1hhgWClibrbdwg-rHj5SmbSo1pQjWS8-V";//【管理】請求支払→◇業者管理◇保存→建設業許可フォルダ
  }else if(e == "保存インボイス"){
    return "1Z3aOdW0kSTIq32I5qYH3Z7sK96_g0uDE";//【管理】請求支払→◇業者管理◇保存→(保存)インボイス
  }else if(e == "一時フォルダ1"){
    return "12xqAGM_8vlz2LyboNvOAEXkG396-P0OO";//一時保管フォルダ（マイドライブ内）
  }else if(e == "一時フォルダ2"){
    return "1nvXo1Ipph-isFBJ4V2XTQECNJIkyPAtk";//一時保管フォルダ（管理フォルダ→ひな形フォルダ内）
  }else if(e == "3ヶ月フォルダ"){
    return "1eVw99s4IYRGs941EzRyYU1e1MJgthPIW";//【管理】請求支払→【契約署名用】開発フォルダ→（契約用）一時フォルダ（３ヶ月）
  }else if(e == "総務共有"){
    return "0AAIT7X4k3F1nUk9PVA";//総務共有ドライブ
  }else if(e == "一般共有"){
    return "1aALv7FvJMVKxrpkcsc0ekO8tapV8MF_M";//【管理】請求支払→一般共有用フォルダ
  }else if(e == "業者支払台帳"){
    return "182Ye-1jjVKGZ-PCIqMALl_tNSSTfvsG-";//【管理】請求支払→業者支払台帳
  }else if(e == "出面まとめフォルダ"){
    return "1QVBC8TwUQ_G8FfXsxssf-aSTqT9GFPrp";//【管理】勤怠管理→出面明細PDFまとめ用一時フォルダ
  }else if(e == "一時フォルダ丸山"){
    return "1uFRaywzGLBD8IWkwTZ_n4p2HCelTOpEZ";//丸山のマイドライブ
  }else if(e == "PDFまとめ用"){
    return "1jtcL4wJX0F3KcZ0TUuGPQl__bfiVkw9r";//【管理】請求支払→PDFまとめ用
  }else if(e == "一般共有用フォルダ"){
    return "1aALv7FvJMVKxrpkcsc0ekO8tapV8MF_M";//【管理】請求支払→一般共有用フォルダ
  }else if(e == "就業規則"){
    return "1l2djzUJ87b73JSbRkdmcDsrl3-rxKnU2";//【管理】勤怠管理→就業規則→最新版
  }else if(e == "覚書一時保管"){
    return "1sYngOC85Wy8ofkErRaeHUylhy8-U2R1q";//【管理】請求支払→◇業者管理◇（保管）→覚書(一時保管）消してもOK
  }else if(e == "安全協力会"){
    return "1ksUHq5Tu-XmuspNKIZu5u0wDwKRF7_Ph";//【管理】請求支払→◇業者管理◇フォルダ→（保存）安全協力会
  }else if(e == "電子文書覚書"){
    return "1xr4L6N5osI1JEJ5Yavn7i4ucO2MIc3pu";//【管理】請求支払→◇業者管理◇フォルダ→（保存）電子文書覚書
  }
}
//添付用説明PDFリンク
function pdfLink(e){
  if(e == "出来高請求説明"){
    return "https://drive.google.com/file/d/1LgLnkLG6bNJtKFyLmTmNxe2OfhWqcnPr/view?usp=share_link";//【管理】請求支払→◇業者管理◇（保存）→出来高請求のやり方のMicrosoft Word - 出来高請求のやり方
  }else if(e == "覚書説明"){
    return "1ResS3CTo8v-nx0s49ooHT7XnPCTDEvue";//【管理】請求支払→◇業者管理◇（保存）→出来高請求のやり方のアドビサインのやり方説明
  }else if(e == "覚書印"){
    return "1hqQJYxROrh0fGPHHc-JlM3-ofE4C3ss_";//印鑑Id //【管理】請求支払→◇業者管理◇（保存）→承認印フォルダ内
  }else if(e == "経費入力マニュアル"){
    return "1qtiexsVj6aMQD89N6HMQ_BQQolwIV267";//【管理】経費管理→マニュアルの【経費精算】基本入力マニュアル◆石井工業20250124
  }else if(e == "経費管理マニュアル"){
    return "1nmDRCf51tuKR7F1oP0ty1YySY-E3_mM1";//【管理】経費管理→マニュアルの【経費精算】管理パネルマニュアル◆石井工業20250217
  }else if(e == "経費管理スクショ"){
    return "1Yrd0j6n6FwzCo61WsaZGN9nlSC-KXOlB"//【管理】経費管理のスクリーンショット 2024-11-17 105625（経費の確認OKがされてないものがある場合にメールに添付する写真）
  }else if(e == "お疲れ様時コンパネ素材"){
    return "https://drive.google.com/file/d/1nmkKSxX8bXCWUQhPIe-KPJ2cEZRhyiuW/view?usp=sharing";//【管理】勤怠管理のお疲れ様時コンパネ素材2
  }else if(e == "新規業者登録やり方"){
    return "https://drive.google.com/file/d/135zBaOY5FbtxsyVd80__hTC4oLp5FYJn/view?usp=share_link";//【管理】請求支払→◇業者管理◇→新規会社登録やり方の新規会社登録のやり方
  }else if(e == "電子契約やり方"){//🔷OK🔷
    return "https://drive.google.com/file/d/1-Cq0Nep2f4J5WYdTU9E2p8diMz09MfoG/view?usp=drivesdk";//URL【管理】請求支払→【ひな形】ファイルの契約書署名と印鑑方法.pdf
  }else if(e == "契約印"){//🔷OK🔷
    return "1TfkXg6HkTsFghItg6FNgvhBT77TVCFZS";//ID【管理】請求支払→【ひな形】ファイルの契約印.jpg
  }else if("契約署名やり方"){//🔷OK🔷
    return "1-Cq0Nep2f4J5WYdTU9E2p8diMz09MfoG"//ID【管理】請求支払→【ひな形】ファイルの契約書署名と印鑑方法PDF
  }
}
//😃経費アプリ
function keihiKengen(name){
  const keihi = ["丸山達也","石井利樹","石井亜也子","浅谷美保","宮原亜美","望月凜花"];
  return keihi.includes(name) ? "OK" : "NG";
}

function getkinds(e){
  if(!e){
    return ["担当者","職人","総務部","管理監督者","外部監査","外注","リモート","事務所","テスト"];
  }else if(e == "管理監督"){
    return ["選択", "石井利樹", "丸山達也", "亀卦川亨"];
  }
}
function カレンダーid(e){
  if(e == "管理"){
    return "system.ishii@ebisu-ishii.co.jp";//システム管理
  }else if(e == "総務部"){
    return "c_62201i9sggtol25cfvgpn41i0o@group.calendar.google.com"//総務部カレンダー
  }
}
function getAdmin(){//権限のある者のIDを返す（新社員情報確認で使用）
  return ["102","103","136"];
}
function adminId(){
  return "103";//管理者のid
}
//😃追加社員リストに追加する人（管理者とか）のROW
function 社員リスト追加(){
  return [36,40,46,54];
}
//😃覚書と協力会のファイルネーム（Adobeサイン反映用）
function ファイル名PDF(e){
  if(e == "協力会"){
    return "安全協力会ご加入のご案内◆";
  }else if(e == "覚書"){
    return "電子文書についての覚書◆"
  }
}
//😃ロゴマーク
function ロゴId(){
  return "15Z0KlyN3SWCDmXJhhCCYJmj_WcG60JEI";
}
//😃😃設定項目😃😃😃😃設定項目😃😃😃😃設定項目😃😃😃😃設定項目😃😃😃😃設定項目😃😃
