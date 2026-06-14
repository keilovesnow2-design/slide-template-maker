// 商品紹介テンプレート（LP・営業資料スタイル）
// 講義資料との主な違い:
//   表紙   : 中央揃えヒーロー + メリットピル（講義は左揃え+縦ストライプ）
//   2枚目  : 3横行レイアウト（講義は3縦カード）
//   3枚目  : 3横行 + 右メリットバッジ（講義は4手順縦カード）
//   5枚目  : 大型CTAボタン + アクションチップ（講義はまとめ3カード）
//   ヘッダー: 英語ラベル + 細横バー（講義はサークル番号）

function shadow() { return { type: "outer", blur: 8, offset: 3, color: "000000", opacity: 0.12, angle: 135 }; }

function buildProductIntro({ theme, subtitle, slideCount, colorTheme }) {
  const T = getTheme(colorTheme);
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = theme;

  addPiCoverSlide(pres, T, theme, subtitle);
  addPiPainSlide(pres, T, theme);
  addPiFeaturesSlide(pres, T, theme);
  addPiBeforeAfterSlide(pres, T, theme);
  addPiCtaSlide(pres, T, theme);

  return pres;
}

// ===============================================================
// 共通ヘッダー
// 講義資料の「細縦ストライプ + サークル番号」とは全く別デザイン
// ===============================================================
function addPiHeader(s, pres, T, sectionLabel, title) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(sectionLabel, {
    x: 0.5, y: 0.17, w: 9, h: 0.3,
    fontFace: "Arial Black", fontSize: 10,
    color: T.accent2, align: "left"
  });
  s.addText(title, {
    x: 0.5, y: 0.43, w: 9, h: 0.7,
    fontFace: "Meiryo", fontSize: 25, bold: true,
    color: T.dark, align: "left", valign: "middle"
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.2, w: 9, h: 0.04,
    fill: { color: T.border }, line: { type: "none" }
  });
}

function addPiFooter(s, pres, T, label) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(label, {
    x: 8.5, y: 5.2, w: 1.3, h: 0.425,
    fontFace: "Arial", fontSize: 11,
    color: T.white, align: "center", margin: 0
  });
}

// ===============================================================
// Slide 1: 表紙（中央揃えヒーロー）
// 講義資料: 左揃え + 左縦ストライプ + サークル装飾
// 商品紹介: 中央揃え + フル幅バー + 3メリットピル
// ===============================================================
function addPiCoverSlide(pres, T, theme, subtitle) {
  const s = pres.addSlide();
  s.background = { color: T.coverBg || T.dark };

  if (T.coverGradient) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 10, h: 5.625,
      fill: { type: "gradient", gradientType: "linear", color: T.coverGradient, angle: 135 },
      line: { type: "none" }
    });
  }

  // 上部フル幅アクセントバー（講義の細縦ストライプと対比）
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.1,
    fill: { color: T.coverStripeColor || T.accent2 }, line: { type: "none" }
  });

  // 商品紹介バッジ（中央）
  s.addShape(pres.ShapeType.rect, {
    x: 3.85, y: 0.32, w: 2.3, h: 0.44,
    fill: { color: T.coverTagBg || T.accent2 }, line: { type: "none" }
  });
  s.addText("商 品 紹 介", {
    x: 3.85, y: 0.34, w: 2.3, h: 0.4,
    fontFace: "Meiryo", fontSize: 12, bold: true,
    color: T.coverTagText || T.dark, align: "center", margin: 0
  });

  // 商品名（中央・大）
  s.addText(theme, {
    x: 0.5, y: 0.92, w: 9, h: 2.3,
    fontFace: "Meiryo", fontSize: 50, bold: true,
    color: T.coverTitle || T.white, align: "center", valign: "middle", paraSpaceAfter: 4
  });

  // センター区切りライン
  s.addShape(pres.ShapeType.rect, {
    x: 3.5, y: 3.3, w: 3.0, h: 0.06,
    fill: { color: T.coverStripeColor || T.accent2 }, line: { type: "none" }
  });

  // キャッチコピー（中央）
  const catchphrase = subtitle || "あなたのビジネスを次のステージへ";
  s.addText(catchphrase, {
    x: 0.5, y: 3.45, w: 9, h: 0.65,
    fontFace: "Meiryo", fontSize: 18, italic: true,
    color: T.coverBg ? T.accent1 : T.accent2, align: "center"
  });

  // 3つのメリットピル（ボーダーのみ・背景なし）
  const pills = ["⚡  簡単に使える", "🏆  高品質・高実績", "💬  手厚いサポート"];
  pills.forEach((pill, i) => {
    const px = 1.05 + i * 2.78;
    s.addShape(pres.ShapeType.rect, {
      x: px, y: 4.22, w: 2.55, h: 0.5,
      fill: { type: "none" },
      line: { color: T.accent2, width: 1.5 }
    });
    s.addText(pill, {
      x: px, y: 4.24, w: 2.55, h: 0.46,
      fontFace: "Meiryo", fontSize: 12, bold: true,
      color: T.accent2, align: "center", valign: "middle"
    });
  });

  // フッター
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText("1 / 5", {
    x: 8.5, y: 5.2, w: 1.3, h: 0.425,
    fontFace: "Arial", fontSize: 11,
    color: T.white, align: "center", margin: 0
  });
}

// ===============================================================
// Slide 2: お客様の課題（3横行 + 番号パネル）
// 講義資料: 「○○とは？」3縦カード
// 商品紹介: 「こんなお悩みありませんか？」3横行・赤番号パネル
// ===============================================================
function addPiPainSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiHeader(s, pres, T, "PROBLEM", "こんなお悩みありませんか？");

  const problems = [
    { num: "01", icon: "⏰",
      title: "時間・コストの課題",
      body: `${theme}なしでは、想定以上の時間とコストがかかり続けます` },
    { num: "02", icon: "😔",
      title: "品質がばらつく",
      body: `作業の質が安定せず、期待する成果が出にくい状況が続きます` },
    { num: "03", icon: "🔄",
      title: "改善できない",
      body: `問題が起きても原因が特定できず、改善のサイクルが止まります` },
  ];

  problems.forEach((p, i) => {
    const y = 1.35 + i * 1.05;

    // 横長行カード（講義の縦カードとは正反対の方向）
    s.addShape(pres.ShapeType.rect, {
      x: 0.4, y, w: 9.2, h: 0.9,
      fill: { color: T.white }, line: { color: T.border, width: 1 }, shadow: shadow()
    });

    // 左：赤番号パネル
    s.addShape(pres.ShapeType.rect, {
      x: 0.4, y, w: 1.1, h: 0.9,
      fill: { color: "B91C1C" }, line: { type: "none" }
    });
    s.addText(p.num, {
      x: 0.4, y: y + 0.08, w: 1.1, h: 0.44,
      fontFace: "Arial Black", fontSize: 24, bold: true,
      color: T.white, align: "center", valign: "middle", margin: 0
    });
    s.addText(p.icon, {
      x: 0.4, y: y + 0.5, w: 1.1, h: 0.36,
      fontSize: 18, align: "center"
    });

    // タイトル（✗付き）
    s.addText(`✗  ${p.title}`, {
      x: 1.62, y: y + 0.08, w: 7.7, h: 0.38,
      fontFace: "Meiryo", fontSize: 15, bold: true,
      color: "B91C1C", valign: "middle"
    });

    // 説明文
    s.addText(p.body, {
      x: 1.62, y: y + 0.5, w: 7.7, h: 0.34,
      fontFace: "Meiryo", fontSize: 12,
      color: T.text, valign: "middle"
    });
  });

  // 解決バナー（下部）
  s.addShape(pres.ShapeType.rect, {
    x: 0.4, y: 4.5, w: 9.2, h: 0.52,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(`▶  これらすべての悩みを  ${theme}  が解決します！`, {
    x: 0.4, y: 4.52, w: 9.2, h: 0.48,
    fontFace: "Meiryo", fontSize: 13, bold: true,
    color: T.white, align: "center", valign: "middle", margin: 0
  });

  addPiFooter(s, pres, T, "2 / 5");
}

// ===============================================================
// Slide 3: 選ばれる3つの理由（3横行 + 右メリットバッジ）
// 講義資料: 手順 STEP 1〜4 縦カード4列
// 商品紹介: FEATURE 01〜03 横行 + 右端にメリットバッジ
// ===============================================================
function addPiFeaturesSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiHeader(s, pres, T, "SOLUTION", "選ばれる3つの理由");

  const featureColors = [T.accent1, T.accent2, T.stepAltBg || T.dark];
  const features = [
    { num: "01", icon: "⚡",
      title: "シンプルで使いやすい",
      body: `${theme}は直感的に操作でき、難しい知識は不要。誰でもすぐに使い始められます。`,
      merit: "時間ロスをゼロに",
      color: featureColors[0] },
    { num: "02", icon: "✨",
      title: "高い品質と信頼性",
      body: `一貫した品質を提供し、期待以上の成果をお届けします。導入実績が品質を証明します。`,
      merit: "確実な成果を実現",
      color: featureColors[1] },
    { num: "03", icon: "🛡️",
      title: "充実したサポート体制",
      body: `導入後も安心のサポートを整えています。疑問はすぐに解決できます。`,
      merit: "安心して長く使える",
      color: featureColors[2] },
  ];

  features.forEach((f, i) => {
    const y = 1.35 + i * 1.1;

    // 横長行（幅9.2インチ ← 講義の縦カード幅2.15と全く異なる）
    s.addShape(pres.ShapeType.rect, {
      x: 0.4, y, w: 9.2, h: 0.95,
      fill: { color: T.white }, line: { color: T.border, width: 1 }, shadow: shadow()
    });

    // 左：カラー番号パネル（数字 + アイコン）
    s.addShape(pres.ShapeType.rect, {
      x: 0.4, y, w: 1.3, h: 0.95,
      fill: { color: f.color }, line: { type: "none" }
    });
    s.addText(f.num, {
      x: 0.4, y: y + 0.08, w: 1.3, h: 0.45,
      fontFace: "Arial Black", fontSize: 26, bold: true,
      color: T.white, align: "center", valign: "middle", margin: 0
    });
    s.addText(f.icon, {
      x: 0.4, y: y + 0.52, w: 1.3, h: 0.38,
      fontSize: 18, align: "center"
    });

    // タイトル
    s.addText(f.title, {
      x: 1.82, y: y + 0.08, w: 5.1, h: 0.38,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.dark, valign: "middle"
    });

    // 説明文
    s.addText(f.body, {
      x: 1.82, y: y + 0.49, w: 5.1, h: 0.38,
      fontFace: "Meiryo", fontSize: 11,
      color: T.text, valign: "top"
    });

    // 右：メリットバッジ（✓ + 一言）
    s.addShape(pres.ShapeType.rect, {
      x: 7.12, y: y + 0.2, w: 2.3, h: 0.55,
      fill: { color: f.color }, line: { type: "none" }
    });
    s.addText(`✓  ${f.merit}`, {
      x: 7.12, y: y + 0.2, w: 2.3, h: 0.55,
      fontFace: "Meiryo", fontSize: 12, bold: true,
      color: T.white, align: "center", valign: "middle", margin: 0
    });
  });

  s.addText(`${theme} で、あなたのビジネスをもっとシンプルに・もっと強く`, {
    x: 0.4, y: 4.72, w: 9.2, h: 0.3,
    fontFace: "Meiryo", fontSize: 12, bold: true,
    color: T.accent1, align: "center"
  });

  addPiFooter(s, pres, T, "3 / 5");
}

// ===============================================================
// Slide 4: 変化フロー（上段：課題タイル → 中段：導入バー → 下段：ベネフィットタイル）
// ===============================================================
function addPiBeforeAfterSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiHeader(s, pres, T, "STORY", `${theme}導入でこう変わる！`);

  // 上段：3つの課題タイル（暗い塗りつぶし）
  const befores = [
    { icon: "⏰", title: "時間がかかる" },
    { icon: "😔", title: "品質がばらつく" },
    { icon: "🔄", title: "改善できない" },
  ];
  befores.forEach((b, i) => {
    const x = 0.35 + i * 3.15;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.35, w: 3.0, h: 0.85,
      fill: { color: T.beforeHeaderBg || T.text }, line: { type: "none" }, shadow: shadow()
    });
    s.addText(`${b.icon}  ${b.title}`, {
      x, y: 1.37, w: 3.0, h: 0.81,
      fontFace: "Meiryo", fontSize: 16, bold: true,
      color: T.white, align: "center", valign: "middle"
    });
  });

  // 中段：フルワイドのアクセントバー（導入ブリッジ）
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 2.26, w: 10, h: 0.62,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(`▼  ${theme}  を導入することで…  ▼`, {
    x: 0.5, y: 2.28, w: 9, h: 0.58,
    fontFace: "Meiryo", fontSize: 16, bold: true,
    color: T.white, align: "center", valign: "middle", margin: 0
  });

  // 下段：3つのベネフィットタイル（白ベース・アクセントボーダー）
  const afterColors = [T.accent1, T.accent2, T.stepAltBg || T.dark];
  const afters = [
    { icon: "🚀", title: "作業時間を大幅に短縮", body: "手作業をなくして\n素早く高品質な\nアウトプットを実現" },
    { icon: "✨", title: "品質が安定・向上",     body: "ミスが減り\n一貫した成果が\n生まれ続けます" },
    { icon: "📈", title: "改善サイクルが加速",   body: "データで課題を把握し\n素早く改善を\n繰り返せます" },
  ];
  afters.forEach((a, i) => {
    const x = 0.35 + i * 3.15;
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.94, w: 3.0, h: 2.06,
      fill: { color: T.white }, line: { color: afterColors[i], width: 2 }, shadow: shadow()
    });
    // 上部アクセントライン
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.94, w: 3.0, h: 0.07,
      fill: { color: afterColors[i] }, line: { type: "none" }
    });
    s.addText(a.icon, {
      x, y: 2.98, w: 3.0, h: 0.5,
      fontSize: 22, align: "center"
    });
    s.addText(a.title, {
      x: x + 0.1, y: 3.5, w: 2.8, h: 0.45,
      fontFace: "Meiryo", fontSize: 13, bold: true,
      color: afterColors[i], align: "center", valign: "middle"
    });
    s.addText(a.body, {
      x: x + 0.1, y: 3.97, w: 2.8, h: 0.9,
      fontFace: "Meiryo", fontSize: 11,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 3
    });
  });

  addPiFooter(s, pres, T, "4 / 5");
}

// ===============================================================
// Slide 5: まとめ・CTA（大型ボタン + アクションチップ）
// 講義資料: まとめ3カード（定義・手順・行動）
// 商品紹介: カードなし・大型CTAボタン + 3行動チップ
// ===============================================================
function addPiCtaSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.summaryBg || T.dark };
  if (T.coverGradient) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 10, h: 5.625,
      fill: { type: "gradient", gradientType: "linear", color: T.coverGradient, angle: 135 },
      line: { type: "none" }
    });
  }

  // 上部ライン
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: T.accent2 }, line: { type: "none" }
  });

  // NEXT ACTION ラベル（「まとめ」とは全く別のメッセージ）
  s.addText("NEXT ACTION", {
    x: 0.5, y: 0.2, w: 9, h: 0.32,
    fontFace: "Arial Black", fontSize: 11,
    color: T.accent2, align: "center"
  });

  // メインCTAヘッドライン（中央・大）
  s.addText(`今すぐ ${theme} を始めましょう`, {
    x: 0.5, y: 0.6, w: 9, h: 0.88,
    fontFace: "Meiryo", fontSize: 30, bold: true,
    color: T.white, align: "center", valign: "middle"
  });

  // 区切りライン
  s.addShape(pres.ShapeType.rect, {
    x: 3.5, y: 1.55, w: 3.0, h: 0.05,
    fill: { color: T.accent2 }, line: { type: "none" }
  });

  // 3つのメリット（チェック形式・カードなし）
  const benefits = [
    "シンプルで使いやすく、すぐに成果が出る",
    "高品質なサポートで、安心して続けられる",
    "導入実績多数、多くのお客様に選ばれています",
  ];
  benefits.forEach((b, i) => {
    s.addText(`✅  ${b}`, {
      x: 1.8, y: 1.72 + i * 0.5, w: 6.4, h: 0.44,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.white, align: "left", valign: "middle"
    });
  });

  // 大型CTAボタン（講義には存在しない要素）
  s.addShape(pres.ShapeType.rect, {
    x: 1.3, y: 3.22, w: 7.4, h: 0.9,
    fill: { color: T.accent2 }, line: { type: "none" }, shadow: shadow()
  });
  s.addText("▶  お問い合わせ・資料請求はこちら", {
    x: 1.3, y: 3.24, w: 7.4, h: 0.86,
    fontFace: "Meiryo", fontSize: 19, bold: true,
    color: T.dark, align: "center", valign: "middle", margin: 0
  });

  // 3つのアクションチップ（ボーダーのみ）
  const actions = ["📧 メールで相談", "📞 電話でお問い合わせ", "📄 無料資料ダウンロード"];
  actions.forEach((a, i) => {
    const ax = 0.5 + i * 3.05;
    s.addShape(pres.ShapeType.rect, {
      x: ax, y: 4.25, w: 2.8, h: 0.52,
      fill: { type: "none" },
      line: { color: T.accent2, width: 1.5 }
    });
    s.addText(a, {
      x: ax, y: 4.27, w: 2.8, h: 0.48,
      fontFace: "Meiryo", fontSize: 11, bold: true,
      color: T.accent2, align: "center", valign: "middle"
    });
  });

  // フッター
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.footerBarColor || T.dark }, line: { type: "none" }
  });
  s.addText(`${theme}　　5 / 5`, {
    x: 0.3, y: 5.2, w: 9.4, h: 0.425,
    fontFace: "Meiryo", fontSize: 11,
    color: T.footerText || T.accent2, align: "center", margin: 0
  });
}
