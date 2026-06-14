const shadow = () => ({ type: "outer", blur: 8, offset: 3, color: "000000", opacity: 0.12, angle: 135 });

function buildProductIntro({ theme, subtitle, slideCount, colorTheme }) {
  const T = getTheme(colorTheme);
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = theme;

  addCoverSlide(pres, T, theme, subtitle);
  addOverviewSlide(pres, T, theme);
  addFeaturesSlide(pres, T, theme);
  addUseCasesSlide(pres, T, theme);
  addCtaSlide(pres, T, theme);

  return pres;
}

// ===============================================================
// Slide 1: 表紙
// ===============================================================
function addCoverSlide(pres, T, theme, subtitle) {
  const s = pres.addSlide();
  s.background = { color: T.coverBg || T.dark };
  if (T.coverGradient) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 10, h: 5.625,
      fill: { type: "gradient", gradientType: "linear", color: T.coverGradient, angle: 135 },
      line: { type: "none" }
    });
  }

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: T.coverStripeColor || T.accent1 }, line: { color: T.coverStripeColor || T.accent1 }
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 0.6, y: 0.5, w: 0.18, h: 0.18,
    fill: { color: T.accent2 }, line: { color: T.accent2 }
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 0.9, y: 0.5, w: 0.12, h: 0.12,
    fill: { color: T.coverStripeColor || T.accent1 }, line: { color: T.coverStripeColor || T.accent1 }
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.55, y: 0.85, w: 2.8, h: 0.38,
    fill: { color: T.coverTagBg || T.accent1 }, line: { color: T.coverTagBg || T.accent1 }
  });
  s.addText("商品紹介", {
    x: 0.55, y: 0.87, w: 2.8, h: 0.34,
    fontFace: "Meiryo", fontSize: 13, bold: true,
    color: T.coverTagText || T.white, align: "center", margin: 0
  });

  s.addText(theme, {
    x: 0.55, y: 1.45, w: 8.5, h: 2.4,
    fontFace: "Meiryo", fontSize: 44, bold: true,
    color: T.coverTitle || T.white, align: "left", paraSpaceAfter: 8
  });

  if (subtitle) {
    if (T.coverTitleFrame) {
      s.addShape(pres.ShapeType.rect, {
        x: 0.55, y: 3.82, w: 3.5, h: 0.035,
        fill: { color: T.coverTitleFrame }, line: { type: "none" }
      });
    }
    s.addText(subtitle, {
      x: 0.55, y: 3.9, w: 8.5, h: 0.6,
      fontFace: "Meiryo", fontSize: 20,
      color: T.coverBg ? T.accent1 : T.accent2, align: "left", italic: true
    });
  }

  if (T.coverTitleFrame) {
    const fx = 0.4, fy = 1.28, fw = 8.8, fh = 3.22;
    const cl = 0.4, ct = 0.035;
    [
      { x: fx,           y: fy,            w: cl, h: ct },
      { x: fx,           y: fy,            w: ct, h: cl },
      { x: fx + fw - cl, y: fy,            w: cl, h: ct },
      { x: fx + fw - ct, y: fy,            w: ct, h: cl },
      { x: fx,           y: fy + fh - ct,  w: cl, h: ct },
      { x: fx,           y: fy + fh - cl,  w: ct, h: cl },
      { x: fx + fw - cl, y: fy + fh - ct,  w: cl, h: ct },
      { x: fx + fw - ct, y: fy + fh - cl,  w: ct, h: cl },
    ].forEach(c => s.addShape(pres.ShapeType.rect, {
      x: c.x, y: c.y, w: c.w, h: c.h,
      fill: { color: T.coverTitleFrame }, line: { type: "none" }
    }));
  }

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });
  s.addText("1 / 5", {
    x: 8.5, y: 5.2, w: 1.3, h: 0.425,
    fontFace: "Arial", fontSize: 11,
    color: T.white, align: "center", margin: 0
  });
}

// ===============================================================
// Slide 2: 商品概要（解決する課題）
// ===============================================================
function addOverviewSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });

  s.addShape(pres.ShapeType.ellipse, {
    x: 0.35, y: 0.3, w: 0.6, h: 0.6,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });
  s.addText("1", {
    x: 0.35, y: 0.32, w: 0.6, h: 0.56,
    fontFace: "Arial Black", fontSize: 22, bold: true,
    color: T.white, align: "center", margin: 0
  });

  s.addText(`${theme}が解決する3つの課題`, {
    x: 1.1, y: 0.3, w: 7.5, h: 0.65,
    fontFace: "Meiryo", fontSize: 26, bold: true,
    color: T.dark, valign: "middle"
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.25, y: 1.1, w: 9.5, h: 0.04,
    fill: { color: T.border }, line: { color: T.border }
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.35, y: 1.25, w: 9.3, h: 0.9,
    fill: { color: T.accent1 }, line: { color: T.accent1 }, shadow: shadow()
  });
  s.addText("多くのお客様が感じているお悩みを解決します", {
    x: 0.35, y: 1.27, w: 9.3, h: 0.86,
    fontFace: "Meiryo", fontSize: 18, bold: true,
    color: T.white, align: "center", valign: "middle", margin: 0
  });

  const issues = [
    { x: 0.35, icon: "📉", title: "時間・コストの課題",
      body: `${theme}なしでは\n想定以上の時間と\nコストがかかります` },
    { x: 3.6,  icon: "😟", title: "品質のばらつき",
      body: `作業の質が安定せず\n期待する成果が\n出にくい状況が続きます` },
    { x: 6.85, icon: "🔄", title: "改善のしにくさ",
      body: `問題が起きても\n原因特定や改善が\n難しい状況です` },
  ];

  issues.forEach(p => {
    s.addShape(pres.ShapeType.rect, {
      x: p.x, y: 2.35, w: 3.0, h: 2.65,
      fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
    });
    s.addShape(pres.ShapeType.rect, {
      x: p.x, y: 2.35, w: 3.0, h: 0.08,
      fill: { color: T.accent2 }, line: { type: "none" }
    });
    s.addText(p.icon, { x: p.x, y: 2.5, w: 3.0, h: 0.7, fontSize: 30, align: "center" });
    s.addText(p.title, {
      x: p.x + 0.15, y: 3.25, w: 2.7, h: 0.55,
      fontFace: "Meiryo", fontSize: 15, bold: true,
      color: T.accent1, align: "center"
    });
    s.addText(p.body, {
      x: p.x + 0.15, y: 3.82, w: 2.7, h: 1.05,
      fontFace: "Meiryo", fontSize: 13,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 3
    });
  });

  addPageNum(s, pres, T, "2 / 5");
}

// ===============================================================
// Slide 3: 特徴・メリット
// ===============================================================
function addFeaturesSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });

  s.addShape(pres.ShapeType.ellipse, {
    x: 0.35, y: 0.3, w: 0.6, h: 0.6,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });
  s.addText("2", {
    x: 0.35, y: 0.32, w: 0.6, h: 0.56,
    fontFace: "Arial Black", fontSize: 22, bold: true,
    color: T.white, align: "center", margin: 0
  });

  s.addText(`${theme}の3つの特徴`, {
    x: 1.1, y: 0.3, w: 7.5, h: 0.65,
    fontFace: "Meiryo", fontSize: 28, bold: true,
    color: T.dark, valign: "middle"
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.25, y: 1.1, w: 9.5, h: 0.04,
    fill: { color: T.border }, line: { color: T.border }
  });

  const featureColors = [T.accent1, T.accent2, T.stepAltBg || T.dark];
  const features = [
    { num: "1", label: "FEATURE", icon: "⚡",
      title: "シンプルで使いやすい",
      body: `${theme}は直感的に操作でき\n難しい知識は不要です。\n誰でもすぐに使い始められます。`,
      color: featureColors[0] },
    { num: "2", label: "FEATURE", icon: "✨",
      title: "高い品質と信頼性",
      body: `${theme}は一貫した品質を提供し\n期待以上の成果をお届けします。\n導入実績が品質を証明します。`,
      color: featureColors[1] },
    { num: "3", label: "FEATURE", icon: "🛡️",
      title: "充実したサポート",
      body: `導入後も安心の\nサポート体制を整えています。\n疑問はすぐに解決できます。`,
      color: featureColors[2] },
  ];

  features.forEach((f, i) => {
    const x = 0.35 + i * 3.15;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.3, w: 2.95, h: 3.65,
      fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.3, w: 2.95, h: 0.55,
      fill: { color: f.color }, line: { type: "none" }
    });
    s.addText(`${f.label} ${f.num}`, {
      x, y: 1.32, w: 2.95, h: 0.51,
      fontFace: "Arial Black", fontSize: 14, bold: true,
      color: T.white, align: "center", margin: 0
    });
    s.addText(f.icon, { x, y: 2.0, w: 2.95, h: 0.65, fontSize: 28, align: "center" });
    s.addText(f.title, {
      x: x + 0.12, y: 2.72, w: 2.71, h: 0.65,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.dark, align: "center", valign: "middle", paraSpaceAfter: 2
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.25, y: 3.4, w: 2.45, h: 0.03,
      fill: { color: T.border }, line: { color: T.border }
    });
    s.addText(f.body, {
      x: x + 0.12, y: 3.5, w: 2.71, h: 1.25,
      fontFace: "Meiryo", fontSize: 12,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 4
    });
  });

  addPageNum(s, pres, T, "3 / 5");
}

// ===============================================================
// Slide 4: 活用シーン（Before / After）
// ===============================================================
function addUseCasesSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.light };

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });

  s.addShape(pres.ShapeType.ellipse, {
    x: 0.35, y: 0.3, w: 0.6, h: 0.6,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });
  s.addText("3", {
    x: 0.35, y: 0.32, w: 0.6, h: 0.56,
    fontFace: "Arial Black", fontSize: 22, bold: true,
    color: T.white, align: "center", margin: 0
  });

  s.addText(`${theme}の活用シーン`, {
    x: 1.1, y: 0.3, w: 7.5, h: 0.65,
    fontFace: "Meiryo", fontSize: 26, bold: true,
    color: T.dark, valign: "middle"
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.25, y: 1.1, w: 9.5, h: 0.04,
    fill: { color: T.border }, line: { color: T.border }
  });

  // Before
  s.addShape(pres.ShapeType.rect, {
    x: 0.35, y: 1.25, w: 4.2, h: 3.7,
    fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.35, y: 1.25, w: 4.2, h: 0.5,
    fill: { color: T.beforeHeaderBg || T.text }, line: { type: "none" }
  });
  s.addText("導入前", {
    x: 0.35, y: 1.27, w: 4.2, h: 0.46,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const beforeItems = ["手作業で時間がかかる", "ミスが発生しやすい", "ノウハウが属人化する", "改善の糸口が見えない", "コストが想定以上にかさむ"];
  beforeItems.forEach((item, i) => {
    s.addText(`✗  ${item}`, {
      x: 0.55, y: 1.9 + i * 0.46, w: 3.8, h: 0.4,
      fontFace: "Meiryo", fontSize: 13, color: "B91C1C", valign: "middle"
    });
  });

  s.addText("→", {
    x: 4.65, y: 2.75, w: 0.5, h: 0.6,
    fontSize: 28, bold: true, color: T.accent1, align: "center"
  });

  // After
  s.addShape(pres.ShapeType.rect, {
    x: 5.25, y: 1.25, w: 4.4, h: 3.7,
    fill: { color: T.white }, line: { color: T.accent2, width: 2 }, shadow: shadow()
  });
  s.addShape(pres.ShapeType.rect, {
    x: 5.25, y: 1.25, w: 4.4, h: 0.5,
    fill: { color: T.accent1 }, line: { type: "none" }
  });
  s.addText(`${theme}導入後 ✨`, {
    x: 5.25, y: 1.27, w: 4.4, h: 0.46,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const afterItems = ["作業時間を大幅に短縮", "ミスが激減し品質向上", "ノウハウをチームで共有", "データで改善策を把握", "コストを最適化できる"];
  afterItems.forEach((item, i) => {
    s.addText(`✓  ${item}`, {
      x: 5.45, y: 1.9 + i * 0.46, w: 4.0, h: 0.4,
      fontFace: "Meiryo", fontSize: 13, color: "065F46", valign: "middle"
    });
  });

  addPageNum(s, pres, T, "4 / 5");
}

// ===============================================================
// Slide 5: まとめ・CTA
// ===============================================================
function addCtaSlide(pres, T, theme) {
  const s = pres.addSlide();
  s.background = { color: T.summaryBg || T.dark };
  if (T.coverGradient) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 10, h: 5.625,
      fill: { type: "gradient", gradientType: "linear", color: T.coverGradient, angle: 135 },
      line: { type: "none" }
    });
  }

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: T.accent2 }, line: { color: T.accent2 }
  });

  s.addText("まとめ", {
    x: 0.55, y: 0.25, w: 9, h: 0.7,
    fontFace: "Meiryo", fontSize: 30, bold: true,
    color: T.white, align: "left"
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.55, y: 1.05, w: 9, h: 0.04,
    fill: { color: T.summaryCardBorder || T.accent1 }, line: { color: T.summaryCardBorder || T.accent1 }
  });

  const cards = [
    {
      num: "01", title: `${theme}の強み`,
      body: `シンプルな操作性と\n高い品質で\n初めての方でも\n安心してお使いいただけます。`
    },
    {
      num: "02", title: "導入のしやすさ",
      body: `最短で導入を開始でき\n手厚いサポートで\nスムーズに\n定着します。`
    },
    {
      num: "03", title: "今すぐ始めよう",
      body: `まずは${theme}の\n効果を体験してみて\nください。\n次の一歩はここから！`
    },
  ];

  cards.forEach((c, i) => {
    const x = 0.35 + i * 3.15;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.2, w: 2.95, h: 3.7,
      fill: { color: T.cardBg || T.darkCard }, line: { color: T.summaryCardBorder || T.accent1, width: 1.5 }, shadow: shadow()
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.2, w: 2.95, h: 0.55,
      fill: { color: T.accent1 }, line: { type: "none" }
    });
    s.addText(c.num, {
      x, y: 1.22, w: 2.95, h: 0.51,
      fontFace: "Arial Black", fontSize: 18, bold: true,
      color: T.white, align: "center", margin: 0
    });
    s.addText(c.title, {
      x: x + 0.12, y: 1.88, w: 2.71, h: 0.85,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.cardTitle || T.accent2, align: "center", valign: "middle", paraSpaceAfter: 3
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.25, y: 2.75, w: 2.45, h: 0.03,
      fill: { color: T.summaryCardBorder || T.accent1 }, line: { color: T.summaryCardBorder || T.accent1 }
    });
    s.addText(c.body, {
      x: x + 0.12, y: 2.85, w: 2.71, h: 1.95,
      fontFace: "Meiryo", fontSize: 12,
      color: T.cardBody || "CBD5E1", align: "center", valign: "top", paraSpaceAfter: 5
    });
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.footerBarColor || T.accent2 },
    line: { color: T.footerBarColor || T.accent2 }
  });
  s.addText(`${theme}　　5 / 5`, {
    x: 0.3, y: 5.2, w: 9.4, h: 0.425,
    fontFace: "Meiryo", fontSize: 11,
    color: T.footerText, align: "center", margin: 0
  });
}

// ===============================================================
// ページ番号フッター（共通）
// ===============================================================
function addPageNum(s, pres, T, label) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.2, w: 10, h: 0.425,
    fill: { color: T.accent1 }, line: { color: T.accent1 }
  });
  s.addText(label, {
    x: 8.5, y: 5.2, w: 1.3, h: 0.425,
    fontFace: "Arial", fontSize: 11,
    color: T.white, align: "center", margin: 0
  });
}
