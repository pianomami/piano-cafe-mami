const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, ShadingType, WidthType, LevelFormat,
  Table, TableRow, TableCell, Header, Footer, PageNumber,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const MAIN   = "B5896A";
const DARK   = "8B6F47";
const LIGHT  = "F5F1ED";
const GRAY   = "8B7B6E";
const WHITE  = "FFFFFF";
const BLACK  = "2C2C2C";

// ─── ヘルパー ─────────────────────────────────────
const sp = (before=60, after=60) => ({ spacing: { before, after } });

function title(text, size=48) {
  return new Paragraph({
    children: [new TextRun({ text, size, bold:true, color:DARK, font:"Arial" })],
    alignment: AlignmentType.CENTER,
    ...sp(0, 80),
  });
}

function rule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MAIN, space: 4 } },
    ...sp(0, 0),
  });
}

function sectionTitle(text) {
  return [
    new Paragraph({
      children: [new TextRun({ text, size:26, bold:true, color:WHITE, font:"Arial" })],
      shading: { fill: MAIN, type: ShadingType.CLEAR },
      indent: { left: 200 },
      keepNext: true,
      keepLines: true,
      ...sp(300, 0),
    }),
  ];
}

function body(text, { indent=false, bold=false, color=BLACK, size=19 }={}) {
  return new Paragraph({
    children: [new TextRun({ text, size, font:"Arial", bold, color })],
    indent: indent ? { left: 360 } : undefined,
    ...sp(50, 50),
  });
}

function check(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: "✅  ", size:19, font:"Arial" }),
      new TextRun({ text, size:19, font:"Arial", color:BLACK }),
    ],
    indent: { left: 360 },
    ...sp(40, 40),
  });
}

function note(text) {
  return new Paragraph({
    children: [new TextRun({ text: "※ " + text, size:17, font:"Arial", color:GRAY, italics:true })],
    indent: { left: 420 },
    ...sp(30, 30),
  });
}

function subheading(text) {
  return new Paragraph({
    children: [new TextRun({ text, size:21, bold:true, color:DARK, font:"Arial" })],
    indent: { left: 200 },
    border: { left: { style: BorderStyle.THICK, size: 12, color: MAIN, space: 8 } },
    keepNext: true,
    keepLines: true,
    ...sp(200, 60),
  });
}

function box(children) {
  const b = { style: BorderStyle.SINGLE, size: 4, color: MAIN };
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        borders: { top:b, bottom:b, left:b, right:b },
        shading: { fill: "FDF9F6", type: ShadingType.CLEAR },
        margins: { top:200, bottom:200, left:280, right:280 },
        width: { size:9026, type:WidthType.DXA },
        children,
      })]
    })]
  });
}

function infoRow(label, value) {
  const bL = { style: BorderStyle.SINGLE, size: 2, color:"DDDDDD" };
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({
        borders: { top:bL, bottom:bL, left:bL, right:bL },
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        margins: { top:100, bottom:100, left:160, right:160 },
        width: { size:2800, type:WidthType.DXA },
        children: [new Paragraph({ children:[new TextRun({ text:label, size:19, bold:true, color:DARK, font:"Arial" })], ...sp(0,0) })]
      }),
      new TableCell({
        borders: { top:bL, bottom:bL, left:bL, right:bL },
        margins: { top:100, bottom:100, left:160, right:160 },
        width: { size:6226, type:WidthType.DXA },
        children: [new Paragraph({ children:[new TextRun({ text:value, size:19, font:"Arial", color:BLACK })], ...sp(0,0) })]
      }),
    ]
  });
}

function infoTable(rows) {
  return new Table({
    width: { size:9026, type:WidthType.DXA },
    columnWidths: [2800, 6226],
    rows: rows.map(([l,v]) => infoRow(l, v)),
  });
}

const blank = () => new Paragraph({ children:[new TextRun("")], ...sp(40,40) });

// ─── ドキュメント ──────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font:"Arial", size:19, color:BLACK } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width:11906, height:16838 },
        margin: { top:1000, bottom:1100, left:1100, right:1100 },
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text:"やさしいピアノカフェ ", size:16, color:GRAY, font:"Arial" }),
              new TextRun({ text:"\t入会のしおり", size:16, color:GRAY, font:"Arial" }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { bottom: { style: BorderStyle.SINGLE, size:4, color:"DDDDDD", space:4 } },
            ...sp(0,40),
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text:"© やさしいピアノカフェ  ", size:16, color:GRAY, font:"Arial" }),
              new TextRun({ text:"\t", size:16 }),
              new TextRun({ children:[PageNumber.CURRENT], size:16, color:GRAY, font:"Arial" }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { top: { style: BorderStyle.SINGLE, size:4, color:"DDDDDD", space:4 } },
            ...sp(40,0),
          })
        ]
      })
    },
    children: [

      // ── タイトルエリア ──
      new Paragraph({
        children: [new TextRun({ text:"🎹  やさしいピアノカフェ ", size:22, color:MAIN, bold:true, font:"Arial" })],
        alignment: AlignmentType.CENTER,
        ...sp(0, 60),
      }),
      new Paragraph({
        children: [new TextRun({ text:"入 会 の し お り", size:52, bold:true, color:DARK, font:"Arial" })],
        alignment: AlignmentType.CENTER,
        ...sp(0, 40),
      }),
      rule(),
      new Paragraph({
        children: [new TextRun({ text:"ご入会いただきありがとうございます。大切なことをまとめました。不明点はいつでもLINEへ😊", size:18, color:GRAY, font:"Arial" })],
        alignment: AlignmentType.CENTER,
        ...sp(60, 40),
      }),

      blank(),

      // ── 1. ウェルカム ──
      ...sectionTitle("😊  まみより、ようこそ！"),
      blank(),
      box([
        new Paragraph({
          children:[new TextRun({ text:"ご入会、本当にありがとうございます。", size:19, bold:true, font:"Arial", color:DARK })],
          ...sp(0,80),
        }),
        body("「弾いてみたい」と思ったその気持ちを大切に、一緒に進んでいきましょう。うまくいかない日も、練習できなかった週も、焦らなくて大丈夫。まずは音楽を楽しむことを一番に考えています。"),
        body("わからないことや不安なことは、いつでも気軽にLINEで聞いてください。どんな小さなことでも、まみがお答えします。"),
      ]),

      blank(),

      // ── 2. レッスンの進め方 ──
      ...sectionTitle("🎹  レッスンの進め方"),
      blank(),

      subheading("🌟 1曲マスタープログラム（3ヶ月・30分×10回・60,000円一括）"),
      check("レッスンは基本的に曜日・時間を固定してお届けします"),
      check("お互いの都合が合う場合は振替可能です。早めにLINEでご相談ください"),
      check("初回に目標曲とロードマップを一緒に決めます"),
      check("全レッスン後にアーカイブ動画＋フィードバックメッセージをお届けします"),
      check("「毎日まみにLINE」サポート付き（1日1往復目安）"),
      check("月1回グループ音楽カフェに無料参加できます"),
      check("3ヶ月後にミニ発表会があります（参加・顔出しは自由）"),
      note("全10回のレッスンは、3ヶ月間続けていただいた方向けの回数です。途中で退会される場合は、レッスン回数は月3回のペースで計算させていただきます"),
      note("定員5名限定。お支払い：前日までにPayPay or 振込"),

      subheading("🎹 定期レッスン（30分×月2回・10,000円/月）"),
      check("レッスンは基本的に曜日・時間を固定してお届けします"),
      check("お互いの都合が合う場合は振替可能です。早めにLINEでご相談ください"),
      check("全レッスン後にアーカイブ動画＋フィードバックメッセージをお届けします"),
      check("週1回までLINEで質問・ご相談OK"),
      check("月1回グループ音楽カフェに無料参加できます"),

      subheading("🎵 単発レッスン（30分・5,500円/回）"),
      check("受けたい時に予約・都度お支払いのシンプルなレッスン"),
      check("全レッスン後にアーカイブ動画をお届けします"),
      check("グループ音楽カフェに300円で参加できます"),

      subheading("☕ グループ音楽カフェ（月1回・無料〜500円）"),
      check("受講生（1曲マスター・定期レッスン）は無料"),
      check("単発レッスンの方は300円、一般の方は500円で参加できます"),
      check("顔出し不要。音楽の話・好きな曲の紹介・日常のおしゃべりなど"),

      subheading("💬 個人おしゃべりカフェ（30分・3,000円）"),
      check("1対1でまみとゆったりおしゃべりする時間"),
      check("顔出し不要。音楽のことも、日常のことも何でもOK"),

      blank(),

      // ── 3. LINEサポート ──
      ...sectionTitle("💬  LINEサポートについて"),
      blank(),

      subheading("🌟 1曲マスタープログラムの方（毎日1往復）"),
      body("レッスン期間中は、1日1往復を目安にLINEでやりとりできます。"),
      check("練習中の疑問・弾けない箇所のご相談"),
      check("今日の練習報告・感想のシェア"),
      check("世間話・ちょっとした雑談も大歓迎です😊"),
      check("まみからも、週に数回「練習どうですか？」など声かけメッセージをお送りします😊"),
      note("毎日でもOK、質問したい時だけでもOK、連絡しない日があっても全然大丈夫です😊"),

      blank(),
      subheading("🎹 定期レッスンの方（週1回まで）"),
      body("レッスンの合間に、週1回までLINEで質問・ご相談ができます。"),
      check("練習中の疑問・弾けない箇所のご相談"),
      check("レッスン日程に関するご相談"),

      blank(),
      subheading("🎵 単発レッスンの方"),
      body("LINEサポートは含まれません。ご質問はレッスン当日にお気軽にどうぞ。"),

      blank(),
      note("お返事は順番にお届けします。すぐに返せない場合もありますがご了承ください"),
      note("LINEはテキストメッセージのやりとりが基本です。動画の添削・送受信は行っていません"),
      note("一度にたくさんのご質問をいただいた場合は、1つずつ日にちをかけてお答えします。あせらず丁寧にお返しします😊"),

      blank(),

      // ── 4. お支払い ──
      ...sectionTitle("💴  お支払いについて"),
      blank(),
      infoTable([
        ["お支払い方法", "PayPay  または  銀行振込"],
        ["PayPay ID", "（ご入会手続き時にLINEでお知らせします）"],
        ["振込先", "（ご入会手続き時にLINEでお知らせします）"],
        ["体験レッスン", "レッスン前日までにお支払いください"],
        ["通常レッスン", "毎月、決められた日までにお支払いください"],
      ]),

      blank(),

      // ── 4.5 オンラインレッスンについて ──
      ...sectionTitle("💻  オンラインレッスンについて"),
      blank(),

      subheading("通信環境について"),
      check("安定したインターネット環境（Wi-Fiなど）をご用意ください"),
      note("接続状況によっては、レッスンの実施やお申込みが難しい場合があります。特に体験レッスンの際に接続がうまくいかない場合は、本申込みにお進みいただけないこともございますので、あらかじめご了承ください"),

      blank(),
      subheading("マイク・カメラについて"),
      check("スマートフォン・タブレット・パソコンなど、マイクとカメラが使える端末をご用意ください"),
      check("レッスン中はビデオ・マイクをオンにしてご参加ください"),

      blank(),
      subheading("身だしなみについて"),
      body("オンラインでも気持ちよくレッスンを進められるよう、人と顔を合わせても問題のない服装でのご参加をお願いいたします。"),

      blank(),
      subheading("レッスンの中止について"),
      body("言葉の暴力や不適切な映像など、レッスンの目的に関係のない行為があった場合には、こちらからレッスンを中止・終了させていただくことがあります。"),

      blank(),

      // ── 5. 退会・キャンセル ──
      ...sectionTitle("📋  退会・キャンセルについて"),
      blank(),

      subheading("コースの退会"),
      body("【3ヶ月集中パーソナルプログラム】"),
      body("翌月開始前までにLINEでご連絡いただければ、使用月数×20,000円を差し引いた残額をご返金します。", { indent:true }),
      infoTable([
        ["1ヶ月で退会の場合", "60,000円 − 20,000円 ＝ 40,000円 返金"],
        ["2ヶ月で退会の場合", "60,000円 − 40,000円 ＝ 20,000円 返金"],
      ]),
      note("退会のご連絡は翌月開始の3日前までにお願いします"),

      blank(),
      body("【定期レッスン】"),
      body("退会をご希望の場合は、次回お支払い分の3日前までにLINEでご連絡ください。お支払い済み分の返金はできません。", { indent:true }),

      blank(),
      body("【単発レッスン】"),
      body("都度払いのため、退会のお手続きは不要です。", { indent:true }),

      blank(),
      subheading("レッスンのキャンセル"),
      infoTable([
        ["前日までのキャンセル", "無料"],
        ["当日キャンセル", "レッスン料の全額をいただきます"],
      ]),

      blank(),

      // ── 6. 申し込みの流れ ──
      ...sectionTitle("📋  申し込みから始めるまでの流れ"),
      blank(),
      infoTable([
        ["STEP 1", "無料相談（15分）を申し込む\nLINEまたはホームページのフォームから"],
        ["STEP 2", "まみからLINEで日程のご連絡\n2営業日以内にお返しします"],
        ["STEP 3", "Zoomで無料相談（15分）\n弾きたい曲・不安なことを何でも"],
        ["STEP 4", "体験レッスン（30分・1,000円）\n実際のレッスンをお試し"],
        ["STEP 5", "コースを選んでご入会\nお支払い後、次回からレッスンスタート！"],
      ]),
      note("無理な勧誘は一切ありません。体験のみでも大歓迎です😊"),

      blank(),

      // ── 7. Zoom ──
      ...sectionTitle("💻  Zoomの準備と設定方法"),
      blank(),
      subheading("アプリのインストール（初回のみ）"),
      check("スマートフォンの方：App Store（iPhone）または Google Play（Android）で「Zoom」を検索してインストール"),
      check("タブレット・iPadの方：同様にアプリをインストール"),
      check("パソコンの方：zoom.us からダウンロード"),
      note("インストールは無料です。アカウント登録は不要です（参加するだけならOK）"),

      blank(),
      subheading("レッスン当日の接続方法"),
      check("まみからLINEで「Zoomリンク」を送ります"),
      check("そのリンクをタップするだけで入れます"),
      check("名前を入力する画面が出たら、お好みの名前（本名でもあだ名でもOK）を入れてください"),
      check("「ビデオをオン」「マイクをオン」にしてご参加ください"),

      blank(),
      subheading("カメラの角度（大切です！）"),
      box([
        body("ピアノが弾けているか確認するため、鍵盤が映るようにカメラを設置してください。"),
        new Paragraph({ children:[new TextRun({ text:"おすすめの置き方", size:19, bold:true, color:DARK, font:"Arial" })], ...sp(80,40) }),
        check("スマホを横向きにして、ピアノの右横か左横に立てる"),
        check("鍵盤と手元が斜め上から映る角度に調整する"),
        check("スマホスタンドや本に立てかけると安定します（100均でOK）"),
        note("顔が映らなくても大丈夫です。鍵盤が映ることが一番重要です😊"),
      ]),

      blank(),
      subheading("接続に不安な方へ"),
      check("無料相談でZoom接続の練習ができます！初めてでも安心してください"),
      check("つながらない・音が聞こえないなどはLINEでご連絡を。一緒に解決します"),

      blank(),

      // ── 8. 励まし ──
      ...sectionTitle("🌿  練習できない日があっても大丈夫"),
      blank(),
      box([
        body("忙しくて練習できなかった週があっても、どうか責めないでください。"),
        body("「今週は弾けなかった」とLINEで教えてくれるだけでOKです。次のレッスンで一緒に取り戻せるよう、まみが考えます。"),
        body("長く続けることより、楽しく続けることが一番大事です 🎹", { bold:true, color:DARK }),
      ]),

      blank(),

      // ── 9. 連絡先 ──
      ...sectionTitle("📱  ご連絡・ご質問"),
      blank(),
      body("レッスンのこと、お支払いのこと、Zoomのこと、何でも気軽にどうぞ。"),
      infoTable([
        ["LINE公式", "https://lin.ee/YNL3n7c"],
        ["YouTube", "https://www.youtube.com/@mamipianocafe"],
        ["Instagram", "https://www.instagram.com/mamipianocafe/"],
      ]),

      blank(),
      rule(),
      blank(),

      new Paragraph({
        children: [new TextRun({ text:"一緒にピアノを楽しみましょう！🎹", size:24, bold:true, color:DARK, font:"Arial" })],
        alignment: AlignmentType.CENTER,
        ...sp(40, 0),
      }),
      new Paragraph({
        children: [new TextRun({ text:"やさしいピアノカフェ", size:20, color:GRAY, font:"Arial" })],
        alignment: AlignmentType.CENTER,
        ...sp(20, 0),
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("C:\\Users\\kuras\\Desktop\\mayumiclaude\\入会のしおり_v12.docx", buf);
  console.log("✅ 完成！");
});
