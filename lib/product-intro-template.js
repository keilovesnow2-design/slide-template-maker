// 関数名は addPi... で統一（lecture-template.js との衝突を完全回避）
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
// 共通ヘッダー（講義資料の「円+数字」と全く異なるデザイン）
// ===============================================================
function addPiSlideHeader(s, pres, T, sectionLabel, title) {
  // 上部アクセントバー（細）
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  // セクションラベル（小さいアルファベット）
  s.addText(sectionLabel, {
    x: 0.5, y: 0.17, w: 9, h: 0.3,
    fontFace: "Arial Black", fontSize: 10,
    color: T.accent2, align: "left"
  });
  // タイトル
  s.addText(title, {
    x: 0.5, y: 0.43, w: 9, h: 0.7,
    fontFace: "Meiryo", fontSize: 25, bold: true,
    color: T.dark, align: "left", valign: "middle"
  });
  // 区切り線
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.2, w: 9, h: 0.04,
    fill: { color: T.border }, line: { type: "none" }
  });
}

// ページ番号フッター
function addPiPageNum(s, pres, T, label) {
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
// Slide 1: 表紙（左右分割ヒーローデザイン）
// ===============================================================
function addPiCoverSlide(pres, T, theme, subtitle) {
  const s = pres.addSlide();
  s.background = { color: T.dark };

  // 左エリアグラデーション
  if (T.coverGradient) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 7.2, h: 5.625,
      fill: { type: "gradient", gradientType: "linear", color: T.coverGradient, angle: 120 },
      line: { type: "none" }
    });
  }

  // 右パネル（アクセント1カラー）
  s.addShape(pres.ShapeType.rect, {
    x: 7.1, y: 0, w: 2.9, h: 5.625,
    fill: { color: T.accent1 }, line: { type: "none" }
  });

  // 上部ライン（左エリアのみ）
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 7.1, h: 0.07,
    fill: { color: T.coverStripeColor || T.accent2 }, line: { type: "none" }
  });

  // 商品紹介バッジ
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 0.25, w: 2.3, h: 0.42,
    fill: { color: T.coverTagBg || T.accent2 }, line: { type: "none" }
  });
  s.addText("商 品 紹 介", {
    x: 0.5, y: 0.27, w: 2.3, h: 0.38,
    fontFace: "Meiryo", fontSize: 12, bold: true,
    color: T.coverTagText || T.dark, align: "center", margin: 0
  });

  // 商品名（大きく・左寄せ）
  s.addText(theme, {
    x: 0.5, y: 0.85, w: 6.3, h: 2.45,
    fontFace: "Meiryo", fontSize: 40, bold: true,
    color: T.coverTitle || T.white, align: "left", valign: "middle", paraSpaceAfter: 6
  });

  // 区切りライン
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 3.45, w: 4.2, h: 0.06,
    fill: { color: T.coverStripeColor || T.accent2 }, line: { type: "none" }
  });

  // キャッチコピー
  const catchphrase = subtitle || "あなたのビジネスを次のステージへ";
  s.addText(catchphrase, {
    x: 0.5, y: 3.6, w: 6.3, h: 0.65,
    fontFace: "Meiryo", fontSize: 17, italic: true,
    color: T.coverBg ? T.accent1 : T.accent2, align: "left"
  });

  // 右パネル：KEY POINT ラベル
  s.addText("KEY POINT", {
    x: 7.15, y: 0.22, w: 2.7, h: 0.38,
    fontFace: "Arial Black", fontSize: 10,
    color: T.white, align: "center"
  });
  s.addShape(pres.ShapeType.rect, {
    x: 7.35, y: 0.65, w: 2.3, h: 0.04,
    fill: { color: T.accent2 }, line: { type: "none" }
  });

  // 右パネル：3つのポイント
  const pts = [
    { icon: "⚡", sub: "SIMPLE",  body: "シンプルで\n使いやすい" },
    { icon: "🏆", sub: "QUALITY", body: "高い品質と\n確かな実績" },
    { icon: "💬", sub: "SUPPORT", body: "充実した\nサポート体制" },
  ];
  pts.forEach((p, i) => {
    const iy = 0.8 + i * 1.52;
    s.addText(p.icon, {
      x: 7.1, y: iy, w: 2.9, h: 0.52,
      fontSize: 22, align: "center"
    });
    s.addText(p.sub, {
      x: 7.1, y: iy + 0.54, w: 2.9, h: 0.28,
      fontFace: "Arial Black", fontSize: 8,
      color: T.accent2, align: "center"
    });
    s.addText(p.body, {
      x: 7.1, y: iy + 0.82, w: 2.9, h: 0.55,
      fontFace: "Meiryo", fontSize: 12, bold: true,
      color: T.white, align: "center", valign: "top", paraSpaceAfter: 2
    });
    if (i < 2) {
      s.addShape(pres.ShapeType.rect, {
        x: 7.4, y: iy + 1.46, w: 2.1, h: 0.03,
        fill: { color: T.accent2 }, line: { type: "none" }
      });
    }
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
// Slide 2: お客様の課題（Pain Points）
// ===============================================================
function addPiPainSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiSlideHeader(s, pres, T, "PROBLEM", "こんなお悩みありませんか？");

  // 強調バナー
  s.addShape(pres.ShapeType.rect, {
    x: 0.4, y: 1.35, w: 9.2, h: 0.75,
    fill: { color: T.dark }, line: { type: "none" }, shadow: shadow()
  });
  s.addText(`「${theme}」を導入していないと…`, {
    x: 0.4, y: 1.37, w: 9.2, h: 0.71,
    fontFace: "Meiryo", fontSize: 17, bold: true,
    color: T.white, align: "center", valign: "middle", margin: 0
  });

  // 3課題カード
  const problems = [
    { icon: "⏰", title: "時間・コストの課題",
      body: `${theme}なしでは\n想定以上の時間と\nコストがかかります` },
    { icon: "😔", title: "品質のばらつき",
      body: `作業の質が安定せず\n期待する成果が\n出にくい状況が続きます` },
    { icon: "🔄", title: "改善のしにくさ",
      body: `問題が起きても\n原因特定や改善が\n難しい状況です` },
  ];

  problems.forEach((p, i) => {
    const x = 0.4 + i * 3.1;
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.25, w: 2.95, h: 2.65,
      fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
    });
    // 左赤アクセント
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.25, w: 0.07, h: 2.65,
      fill: { color: "B91C1C" }, line: { type: "none" }
    });
    s.addText(p.icon, {
      x: x + 0.1, y: 2.35, w: 2.75, h: 0.55,
      fontSize: 26, align: "center"
    });
    s.addText("✗  " + p.title, {
      x: x + 0.1, y: 2.95, w: 2.75, h: 0.52,
      fontFace: "Meiryo", fontSize: 13, bold: true,
      color: "B91C1C", align: "center", valign: "middle"
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.25, y: 3.5, w: 2.45, h: 0.03,
      fill: { color: T.border }, line: { type: "none" }
    });
    s.addText(p.body, {
      x: x + 0.1, y: 3.6, w: 2.75, h: 1.1,
      fontFace: "Meiryo", fontSize: 12,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 3
    });
  });

  // 底部メッセージ
  s.addText(`▶  これらのお悩みを ${theme} が解決します！`, {
    x: 0.4, y: 4.97, w: 9.2, h: 0.28,
    fontFace: "Meiryo", fontSize: 11, bold: true,
    color: T.accent1, align: "center"
  });

  addPiPageNum(s, pres, T, "2 / 5");
}

// ===============================================================
// Slide 3: 特徴・できること（Features）
// ===============================================================
function addPiFeaturesSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiSlideHeader(s, pres, T, "SOLUTION", `${theme}でできること — 3つの特徴`);

  const featureColors = [T.accent1, T.accent2, T.stepAltBg || T.dark];
  const features = [
    { num: "01", icon: "⚡",
      title: "シンプルで使いやすい",
      body: `${theme}は直感的に操作でき\n難しい知識は不要。\n誰でもすぐに使い始められます。`,
      merit: "→ 時間ロスをゼロに",
      color: featureColors[0] },
    { num: "02", icon: "✨",
      title: "高い品質と信頼性",
      body: `${theme}は一貫した品質を提供。\n期待以上の成果をお届けし\n導入実績が品質を証明します。`,
      merit: "→ 成果を確実に上げる",
      color: featureColors[1] },
    { num: "03", icon: "🛡️",
      title: "充実したサポート",
      body: `導入後も安心の\nサポート体制を整えています。\n疑問はすぐに解決できます。`,
      merit: "→ 安心して長く使える",
      color: featureColors[2] },
  ];

  features.forEach((f, i) => {
    const x = 0.4 + i * 3.1;

    // カード外枠
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.38, w: 2.95, h: 3.62,
      fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
    });

    // 上部細カラーバー
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.38, w: 2.95, h: 0.06,
      fill: { color: f.color }, line: { type: "none" }
    });

    // FEATURE + 番号ヘッダー（ダーク背景・白テキスト）
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.44, w: 2.95, h: 0.5,
      fill: { color: T.dark }, line: { type: "none" }
    });
    s.addText(`FEATURE  ${f.num}`, {
      x, y: 1.46, w: 2.95, h: 0.46,
      fontFace: "Arial Black", fontSize: 13, bold: true,
      color: f.color, align: "center", margin: 0
    });

    // アイコン（大）
    s.addText(f.icon, {
      x, y: 2.0, w: 2.95, h: 0.62,
      fontSize: 28, align: "center"
    });

    // タイトル
    s.addText(f.title, {
      x: x + 0.1, y: 2.65, w: 2.75, h: 0.55,
      fontFace: "Meiryo", fontSize: 13, bold: true,
      color: T.dark, align: "center", valign: "middle"
    });

    // 区切り
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.3, y: 3.23, w: 2.35, h: 0.03,
      fill: { color: T.border }, line: { type: "none" }
    });

    // 本文
    s.addText(f.body, {
      x: x + 0.1, y: 3.3, w: 2.75, h: 1.38,
      fontFace: "Meiryo", fontSize: 12,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 3
    });

    // メリットバンド（カード下端）
    s.addShape(pres.ShapeType.rect, {
      x, y: 4.72, w: 2.95, h: 0.28,
      fill: { color: f.color }, line: { type: "none" }
    });
    s.addText(f.merit, {
      x, y: 4.72, w: 2.95, h: 0.28,
      fontFace: "Meiryo", fontSize: 11, bold: true,
      color: T.white, align: "center", valign: "middle", margin: 0
    });
  });

  addPiPageNum(s, pres, T, "3 / 5");
}

// ===============================================================
// Slide 4: 導入前 / 導入後（Before / After）
// ===============================================================
function addPiBeforeAfterSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  addPiSlideHeader(s, pres, T, "COMPARISON", "導入前 vs 導入後 — 何が変わる？");

  // ── Before パネル ──
  s.addShape(pres.ShapeType.rect, {
    x: 0.4, y: 1.33, w: 4.1, h: 3.72,
    fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.4, y: 1.33, w: 4.1, h: 0.54,
    fill: { color: T.beforeHeaderBg || T.text }, line: { type: "none" }
  });
  s.addText("導 入 前　😔", {
    x: 0.4, y: 1.35, w: 4.1, h: 0.5,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const befores = [
    "手作業で時間がかかる",
    "ミスが発生しやすい",
    "ノウハウが属人化する",
    "改善の糸口が見えない",
    "コストが想定以上にかさむ",
  ];
  befores.forEach((item, i) => {
    s.addText(`✗  ${item}`, {
      x: 0.6, y: 2.02 + i * 0.47, w: 3.7, h: 0.42,
      fontFace: "Meiryo", fontSize: 13, color: "B91C1C", valign: "middle"
    });
  });

  // ── 中央：導入矢印 ──
  s.addShape(pres.ShapeType.rect, {
    x: 4.58, y: 2.25, w: 0.88, h: 1.65,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText("導\n入\n↓", {
    x: 4.58, y: 2.25, w: 0.88, h: 1.65,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", valign: "middle", paraSpaceAfter: 4
  });

  // ── After パネル ──
  s.addShape(pres.ShapeType.rect, {
    x: 5.52, y: 1.33, w: 4.08, h: 3.72,
    fill: { color: T.white }, line: { color: T.accent2, width: 2 }, shadow: shadow()
  });
  s.addShape(pres.ShapeType.rect, {
    x: 5.52, y: 1.33, w: 4.08, h: 0.54,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(`導 入 後　✨`, {
    x: 5.52, y: 1.35, w: 4.08, h: 0.5,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const afters = [
    "作業時間を大幅に短縮",
    "ミスが激減し品質向上",
    "ノウハウをチームで共有",
    "データで改善策を把握",
    "コストを最適化できる",
  ];
  afters.forEach((item, i) => {
    s.addText(`✓  ${item}`, {
      x: 5.72, y: 2.02 + i * 0.47, w: 3.7, h: 0.42,
      fontFace: "Meiryo", fontSize: 13, color: "065F46", valign: "middle"
    });
  });

  addPiPageNum(s, pres, T, "4 / 5");
}

// ===============================================================
// Slide 5: まとめ・CTA
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

  // タイトル
  s.addText("まとめ", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontFace: "Meiryo", fontSize: 27, bold: true,
    color: T.white, align: "left"
  });

  // 区切り線
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 0.88, w: 9, h: 0.04,
    fill: { color: T.summaryCardBorder || T.accent2 }, line: { type: "none" }
  });

  // 3価値カード
  const values = [
    { num: "01", icon: "🚀",
      title: `${theme}の強み`,
      body: `シンプルな操作性と\n高い品質で\n初めての方でも\n安心してお使いいただけます。` },
    { num: "02", icon: "📈",
      title: "スムーズな導入",
      body: `最短で導入を開始でき\n手厚いサポートで\nスムーズに\n定着します。` },
    { num: "03", icon: "✉️",
      title: "今すぐ行動を",
      body: `まずは${theme}の\n効果を体験してみて\nください。\n次の一歩はここから！` },
  ];

  values.forEach((v, i) => {
    const x = 0.35 + i * 3.15;

    s.addShape(pres.ShapeType.rect, {
      x, y: 1.05, w: 2.95, h: 3.28,
      fill: { color: T.cardBg || T.darkCard },
      line: { color: T.summaryCardBorder || T.accent2, width: 1.5 },
      shadow: shadow()
    });
    // カードヘッダー
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.05, w: 2.95, h: 0.52,
      fill: { color: T.accent1 }, line: { type: "none" }
    });
    s.addText(v.num, {
      x, y: 1.07, w: 2.95, h: 0.48,
      fontFace: "Arial Black", fontSize: 17, bold: true,
      color: T.white, align: "center", margin: 0
    });
    // アイコン
    s.addText(v.icon, {
      x, y: 1.62, w: 2.95, h: 0.56,
      fontSize: 24, align: "center"
    });
    // タイトル
    s.addText(v.title, {
      x: x + 0.1, y: 2.22, w: 2.75, h: 0.55,
      fontFace: "Meiryo", fontSize: 13, bold: true,
      color: T.cardTitle || T.accent2, align: "center", valign: "middle"
    });
    // 区切り
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.25, y: 2.8, w: 2.45, h: 0.03,
      fill: { color: T.summaryCardBorder || T.accent2 }, line: { type: "none" }
    });
    // 本文
    s.addText(v.body, {
      x: x + 0.1, y: 2.87, w: 2.75, h: 1.33,
      fontFace: "Meiryo", fontSize: 11,
      color: T.cardBody || "CBD5E1", align: "center", valign: "top", paraSpaceAfter: 4
    });
  });

  // CTAバンド（黒スライドのフッター前に入れる）
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 4.43, w: 10, h: 0.68,
    fill: { color: T.accent2 }, line: { type: "none" }
  });
  s.addText(`まずは ${theme} をお試しください　　▶ お問い合わせ・資料請求はこちら`, {
    x: 0.3, y: 4.45, w: 9.4, h: 0.64,
    fontFace: "Meiryo", fontSize: 13, bold: true,
    color: T.dark, align: "center", valign: "middle", margin: 0
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
