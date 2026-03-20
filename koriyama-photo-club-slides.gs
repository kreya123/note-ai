/**
 * 郡山写真部 活動内容まとめ - Google Slides 自動生成スクリプト
 *
 * 使い方：
 * 1. Google ドライブで新規スプレッドシートを開く
 * 2. 拡張機能 > Apps Script を開く
 * 3. このスクリプトを貼り付けて保存
 * 4. 「createKoriyamaPhotoClubSlides」を選択して実行
 * 5. Googleドライブに「郡山写真部 活動まとめ」というスライドが作成される
 */

function createKoriyamaPhotoClubSlides() {
  const presentation = SlidesApp.create("郡山写真部 活動まとめ");

  // ─────────────────────────────────────────────
  // スライド 1：タイトル
  // ─────────────────────────────────────────────
  const slide1 = presentation.getSlides()[0];
  slide1.getBackground().setSolidFill("#1a1a2e");

  const title = slide1.insertTextBox("郡山写真部\n活動内容まとめ");
  title.setLeft(60).setTop(140).setWidth(760).setHeight(160);
  const titleStyle = title.getText().getTextStyle();
  titleStyle.setFontSize(48).setBold(true).setForegroundColor("#e8e8f0");
  title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const sub = slide1.insertTextBox("郡山市SNSフォトブランディングプロジェクト\n2018年9月発足");
  sub.setLeft(60).setTop(310).setWidth(760).setHeight(80);
  sub.getText().getTextStyle().setFontSize(22).setForegroundColor("#a0a8c0");
  sub.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const date = slide1.insertTextBox("調査日：2026年3月20日");
  date.setLeft(60).setTop(490).setWidth(760).setHeight(30);
  date.getText().getTextStyle().setFontSize(14).setForegroundColor("#606080");
  date.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // ─────────────────────────────────────────────
  // スライド 2：概要・目的
  // ─────────────────────────────────────────────
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide2.getBackground().setSolidFill("#f7f7fb");
  addSectionHeader_(slide2, "01", "郡山写真部とは");

  const overview = [
    "▍ 正式名称",
    "　郡山SNSフォトプロジェクト・郡山写真部",
    "",
    "▍ 発足",
    "　2018年9月（郡山市主導）",
    "",
    "▍ 目的",
    "　市民が「住んでいるからこそ」の視点でまちの魅力を撮影し、",
    "　SNSで発信。「郡山に行きたい！」と思ってもらうことを目標に、",
    "　観光誘客につなげる取り組み。",
    "",
    "▍ ハッシュタグ",
    "　#郡山写真部",
  ].join("\n");

  const box2 = slide2.insertTextBox(overview);
  box2.setLeft(60).setTop(130).setWidth(760).setHeight(360);
  box2.getText().getTextStyle().setFontSize(16).setForegroundColor("#222240");

  // ─────────────────────────────────────────────
  // スライド 3：メンバー構成
  // ─────────────────────────────────────────────
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide3.getBackground().setSolidFill("#f7f7fb");
  addSectionHeader_(slide3, "02", "メンバー構成");

  const members = [
    "▍ 規模",
    "　約 30 名（年度ごとに公募・抽選あり）",
    "",
    "▍ 構成",
    "　・郡山市民を中心に、東京からの移住者・現役大学生など多様な層",
    "　・Instagramフォロワー1,000人以上の上級者から初心者まで幅広い",
    "",
    "▍ 参加条件",
    "　・郡山在住／在勤／在学",
    "　・SNSでの情報発信に意欲がある方",
    "",
    "▍ 公式Instagram",
    "　@koriyama_photo（フォロワー約630人・投稿133件）",
  ].join("\n");

  const box3 = slide3.insertTextBox(members);
  box3.setLeft(60).setTop(130).setWidth(760).setHeight(360);
  box3.getText().getTextStyle().setFontSize(16).setForegroundColor("#222240");

  // ─────────────────────────────────────────────
  // スライド 4：主な活動一覧
  // ─────────────────────────────────────────────
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide4.getBackground().setSolidFill("#1a1a2e");
  addSectionHeader_(slide4, "03", "主な活動一覧", "#e8e8f0", "#4a4a7e");

  const activities = [
    [0, "#3a6bc4", "SNS発信", "Instagram中心に#郡山写真部タグで継続発信"],
    [1, "#2e8b57", "フォトまちセミナー", "プロ講師を招いた勉強会（年4回程度）"],
    [2, "#b5651d", "写真展・フォトフェス", "年1回の大規模展示会（最大1,500名来場）"],
    [3, "#7b3fa0", "撮影ツアー", "郡山フォトスポットを巡るツアー"],
    [4, "#c0392b", "フォトスポットガイド制作", "デジタルフォトマップ・PDFを公開"],
  ];

  activities.forEach(([i, color, name, desc]) => {
    const y = 130 + i * 70;
    const dot = slide4.insertShape(SlidesApp.ShapeType.RECTANGLE);
    dot.setLeft(60).setTop(y + 8).setWidth(6).setHeight(36);
    dot.getFill().setSolidFill(color);
    dot.getBorder().setTransparent();

    const label = slide4.insertTextBox(name);
    label.setLeft(80).setTop(y).setWidth(200).setHeight(36);
    label.getText().getTextStyle().setFontSize(16).setBold(true).setForegroundColor(color);

    const descBox = slide4.insertTextBox(desc);
    descBox.setLeft(290).setTop(y + 4).setWidth(520).setHeight(30);
    descBox.getText().getTextStyle().setFontSize(15).setForegroundColor("#c8c8e0");
  });

  // ─────────────────────────────────────────────
  // スライド 5：フォトまちセミナー詳細
  // ─────────────────────────────────────────────
  const slide5 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide5.getBackground().setSolidFill("#f7f7fb");
  addSectionHeader_(slide5, "04", "フォトまちセミナー（勉強会）");

  const seminar = [
    "▍ 開催頻度：年4回程度",
    "",
    "▍ 内容",
    "　・カメラ基礎・構図の理論",
    "　・RAWデータ現像・レタッチ実践",
    "　・ポートレート撮影会",
    "　・商品撮影テクニック",
    "　・ポスター制作実践（フォトフェス向け）",
    "",
    "▍ 講師",
    "　プロカメラマン 佐久間正人 氏 など",
    "",
    "▍ 特徴",
    "　初心者から中上級者まで対応した段階的なカリキュラム",
  ].join("\n");

  const box5 = slide5.insertTextBox(seminar);
  box5.setLeft(60).setTop(130).setWidth(760).setHeight(360);
  box5.getText().getTextStyle().setFontSize(16).setForegroundColor("#222240");

  // ─────────────────────────────────────────────
  // スライド 6：写真展・フォトフェス実績
  // ─────────────────────────────────────────────
  const slide6 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide6.getBackground().setSolidFill("#f7f7fb");
  addSectionHeader_(slide6, "05", "写真展・フォトフェス実績");

  const exhibitions = [
    ["2019年", "東京・京橋ギャラリーで写真展「こおりやま」開催\n「こ・お・り・や・ま」の頭文字をテーマにした5セクション構成"],
    ["2021年", "屋外写真展「フォトフェス〜あったか、ばんだいあたみ！〜」\n磐梯熱海温泉エリアで初の屋外開催"],
    ["2022年", "エスパル郡山で約1,500名、ほっとあたみで約870名来場\n郡山市立美術館と連携（永遠のソール・ライター展）"],
  ];

  exhibitions.forEach(([year, detail], i) => {
    const y = 140 + i * 120;
    const yearBox = slide6.insertTextBox(year);
    yearBox.setLeft(60).setTop(y).setWidth(120).setHeight(40);
    yearBox.getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor("#3a6bc4");

    const detailBox = slide6.insertTextBox(detail);
    detailBox.setLeft(200).setTop(y).setWidth(620).setHeight(80);
    detailBox.getText().getTextStyle().setFontSize(15).setForegroundColor("#222240");

    if (i < exhibitions.length - 1) {
      const line = slide6.insertLine(SlidesApp.LineCategory.STRAIGHT,
        60, y + 95, 820, y + 95);
      line.getLineFill().setSolidFill("#ddddee");
    }
  });

  // ─────────────────────────────────────────────
  // スライド 7：外部連携・メディア
  // ─────────────────────────────────────────────
  const slide7 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide7.getBackground().setSolidFill("#1a1a2e");
  addSectionHeader_(slide7, "06", "外部連携・メディア実績", "#e8e8f0", "#4a4a7e");

  const media = [
    ["PHaT PHOTO「Hello Local」連載",
     "写真誌PHaT PHOTOウェブ連載「フォトまち便り Hello Local」に参加\nvol.15・18・25・28・30・33・35・41・44・47 など多数の回を担当"],
    ["東京カメラ部タイアップ",
     "「こおりやま広域圏×東京カメラ部 つながるフォトコンテスト」を毎年開催\n受賞作品を渋谷ヒカリエ「東京カメラ部写真展」で展示"],
    ["観光行政との連携",
     "部員撮影写真が郡山市HP・フリーペーパー・観光パンフ・観光PRに二次利用\n（公式連携の仕組みあり）"],
  ];

  media.forEach(([title, detail], i) => {
    const y = 135 + i * 110;
    const numBox = slide7.insertTextBox(`0${i + 1}`);
    numBox.setLeft(60).setTop(y).setWidth(50).setHeight(40);
    numBox.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor("#4a8cdd");

    const titleBox = slide7.insertTextBox(title);
    titleBox.setLeft(120).setTop(y).setWidth(700).setHeight(30);
    titleBox.getText().getTextStyle().setFontSize(17).setBold(true).setForegroundColor("#e8e8f0");

    const detailBox = slide7.insertTextBox(detail);
    detailBox.setLeft(120).setTop(y + 32).setWidth(700).setHeight(60);
    detailBox.getText().getTextStyle().setFontSize(14).setForegroundColor("#a0a8c0");
  });

  // ─────────────────────────────────────────────
  // スライド 8：まとめ
  // ─────────────────────────────────────────────
  const slide8 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide8.getBackground().setSolidFill("#1a1a2e");

  const header8 = slide8.insertTextBox("まとめ");
  header8.setLeft(60).setTop(50).setWidth(760).setHeight(50);
  header8.getText().getTextStyle().setFontSize(32).setBold(true).setForegroundColor("#e8e8f0");

  const points = [
    "郡山市が主導する「観光×写真×SNS」の行政プロジェクト",
    "約30名の公募制市民が年間を通じて組織的に活動",
    "勉強会・写真展・撮影ツアーと多角的なプログラムを展開",
    "PHaT PHOTOへの連載・東京カメラ部との協働で全国発信",
    "撮影成果物は観光PR素材として行政に二次利用される仕組み",
  ];

  points.forEach((point, i) => {
    const y = 120 + i * 72;
    const numCircle = slide8.insertShape(SlidesApp.ShapeType.ELLIPSE);
    numCircle.setLeft(60).setTop(y).setWidth(36).setHeight(36);
    numCircle.getFill().setSolidFill("#3a6bc4");
    numCircle.getBorder().setTransparent();

    const numLabel = slide8.insertTextBox(`${i + 1}`);
    numLabel.setLeft(60).setTop(y + 2).setWidth(36).setHeight(30);
    numLabel.getText().getTextStyle().setFontSize(16).setBold(true).setForegroundColor("#ffffff");
    numLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const pointBox = slide8.insertTextBox(point);
    pointBox.setLeft(110).setTop(y + 4).setWidth(710).setHeight(36);
    pointBox.getText().getTextStyle().setFontSize(17).setForegroundColor("#d0d8f0");
  });

  // ─────────────────────────────────────────────
  // スライド 9：参考資料
  // ─────────────────────────────────────────────
  const slide9 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide9.getBackground().setSolidFill("#f7f7fb");
  addSectionHeader_(slide9, "", "参考資料");

  const sources = [
    "郡山市公式ウェブサイト（郡山SNSフォトプロジェクト）",
    "　https://www.city.koriyama.lg.jp/bunka_sports_kanko/kanko/4/koriyamaphoto/index.html",
    "",
    "郡山市観光協会（フォトフェス記事）",
    "　https://www.kanko-koriyama.gr.jp/",
    "",
    "PHaT PHOTO「フォトまち便り Hello Local」",
    "　https://phat-ext.com/",
    "",
    "東京カメラ部「つながるフォトコンテスト2025」",
    "　https://tokyocameraclub.com/koriyamakoikiken/contest2025/",
    "",
    "公式Instagram: @koriyama_photo",
    "　https://www.instagram.com/koriyama_photo/",
  ].join("\n");

  const box9 = slide9.insertTextBox(sources);
  box9.setLeft(60).setTop(130).setWidth(760).setHeight(360);
  box9.getText().getTextStyle().setFontSize(13).setForegroundColor("#444466");

  // 完了通知
  const url = presentation.getUrl();
  Logger.log("スライド作成完了: " + url);
  SlidesApp.getActivePresentation && Browser.msgBox("スライド作成完了！\n\n" + url);
}

// ── ヘルパー：セクションヘッダー追加 ──────────────────────────────
function addSectionHeader_(slide, num, title, textColor, accentColor) {
  textColor = textColor || "#1a1a2e";
  accentColor = accentColor || "#3a6bc4";

  if (num) {
    const numBox = slide.insertTextBox(num);
    numBox.setLeft(60).setTop(40).setWidth(60).setHeight(50);
    numBox.getText().getTextStyle().setFontSize(36).setBold(true).setForegroundColor(accentColor);
  }

  const titleBox = slide.insertTextBox(title);
  titleBox.setLeft(num ? 120 : 60).setTop(48).setWidth(700).setHeight(50);
  titleBox.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor(textColor);

  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, 60, 108, 820, 108);
  line.getLineFill().setSolidFill(accentColor);
  line.setWeight(2);
}
