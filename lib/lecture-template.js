const shadow = () => ({ type: "outer", blur: 8, offset: 3, color: "000000", opacity: 0.12, angle: 135 });

function buildLecture({ theme, subtitle, slideCount, colorTheme }) {
  const T = getTheme(colorTheme);
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = theme;

  addCoverSlide(pres, T, theme, subtitle, "講義資料");
  addWhatIsSlide(pres, T, theme);
  addStepsSlide(pres, T, theme);
  addComparisonSlide(pres, T, theme);
  addSummarySlide(pres, T, theme);

  return pres;
}

// ===============================================================
// Slide 1: 表紙
// ===============================================================
function addCoverSlide(pres, T, theme, subtitle, tagLabel) {
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
  s.addText(tagLabel, {
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
// Slide 2: ○○とは？
// ===============================================================
function addWhatIsSlide(pres, T, theme) {
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

  s.addText(`${theme}とは？`, {
    x: 1.1, y: 0.3, w: 7.5, h: 0.65,
    fontFace: "Meiryo", fontSize: 28, bold: true,
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
  s.addText(`${theme} = 目的を達成するための体系的な方法・知識・スキルの総称です`, {
    x: 0.35, y: 1.27, w: 9.3, h: 0.86,
    fontFace: "Meiryo", fontSize: 18, bold: true,
    color: T.white, align: "center", valign: "middle", margin: 0
  });

  const points = [
    { x: 0.35, icon: "📌", title: "定義",       body: `${theme}を一言で表すと\n「目的達成のための\n技術・方法」` },
    { x: 3.6,  icon: "💡", title: "なぜ重要か", body: `現代において\n${theme}は\n必須スキルのひとつです` },
    { x: 6.85, icon: "🎯", title: "活用場面",   body: `${theme}は\n様々なシーンで\n幅広く活用できます` },
  ];

  points.forEach(p => {
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
// Slide 3: 手順・作り方
// ===============================================================
function addStepsSlide(pres, T, theme) {
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

  s.addText(`${theme}の手順`, {
    x: 1.1, y: 0.3, w: 7.5, h: 0.65,
    fontFace: "Meiryo", fontSize: 28, bold: true,
    color: T.dark, valign: "middle"
  });

  s.addShape(pres.ShapeType.rect, {
    x: 0.25, y: 1.1, w: 9.5, h: 0.04,
    fill: { color: T.border }, line: { color: T.border }
  });

  const stepColors = [T.accent1, T.accent2, T.accent1, T.stepAltBg || T.dark];
  const steps = [
    { num: "1", title: "目標を決める",   body: `${theme}で\n何を達成したいかを\n明確にする`,       color: stepColors[0] },
    { num: "2", title: "基礎を学ぶ",     body: `${theme}の\n基本概念・用語を\n理解する`,           color: stepColors[1] },
    { num: "3", title: "実践する",       body: `実際に${theme}を\n試してフィードバックを\n得る`,   color: stepColors[2] },
    { num: "4", title: "改善・応用する", body: `${theme}をより深く\n理解して\n応用範囲を広げる`, color: stepColors[3] },
  ];

  steps.forEach((step, i) => {
    const x = 0.35 + i * 2.35;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.3, w: 2.15, h: 3.5,
      fill: { color: T.white }, line: { color: T.border, width: 1.5 }, shadow: shadow()
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.3, w: 2.15, h: 0.55,
      fill: { color: step.color }, line: { type: "none" }
    });
    s.addText(`STEP ${step.num}`, {
      x, y: 1.32, w: 2.15, h: 0.51,
      fontFace: "Arial Black", fontSize: 14, bold: true,
      color: T.white, align: "center", margin: 0
    });
    s.addText(step.title, {
      x: x + 0.1, y: 2.0, w: 1.95, h: 0.7,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.dark, align: "center", valign: "middle", paraSpaceAfter: 2
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.2, y: 2.72, w: 1.75, h: 0.03,
      fill: { color: T.border }, line: { color: T.border }
    });
    s.addText(step.body, {
      x: x + 0.1, y: 2.82, w: 1.95, h: 1.85,
      fontFace: "Meiryo", fontSize: 12,
      color: T.text, align: "center", valign: "top", paraSpaceAfter: 4
    });
    if (i < steps.length - 1) {
      s.addText("▶", {
        x: x + 2.05, y: 2.65, w: 0.3, h: 0.35,
        fontSize: 14, color: T.accent1, align: "center"
      });
    }
  });

  s.addText(`✅ STEP を順番に進めれば ${theme} の基礎を習得できます`, {
    x: 0.35, y: 4.86, w: 9.3, h: 0.28,
    fontFace: "Meiryo", fontSize: 11, bold: true,
    color: T.accent1, align: "center"
  });

  addPageNum(s, pres, T, "3 / 5");
}

// ===============================================================
// Slide 4: 比較・実例
// ===============================================================
function addComparisonSlide(pres, T, theme) {
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

  s.addText(`${theme}の比較・実例`, {
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
  s.addText(`${theme}なし`, {
    x: 0.35, y: 1.27, w: 4.2, h: 0.46,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const beforeItems = ["時間がかかる", "品質がばらつく", "ノウハウが属人化する", "改善サイクルが遅い", "再現性が低い"];
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
  s.addText(`${theme}あり ✨`, {
    x: 5.25, y: 1.27, w: 4.4, h: 0.46,
    fontFace: "Meiryo", fontSize: 14, bold: true,
    color: T.white, align: "center", margin: 0
  });

  const afterItems = ["スピードが大幅に向上", "品質が安定する", "チームで共有できる", "改善サイクルが加速", "高い再現性を実現"];
  afterItems.forEach((item, i) => {
    s.addText(`✓  ${item}`, {
      x: 5.45, y: 1.9 + i * 0.46, w: 4.0, h: 0.4,
      fontFace: "Meiryo", fontSize: 13, color: "065F46", valign: "middle"
    });
  });

  addPageNum(s, pres, T, "4 / 5");
}

// ===============================================================
// Slide 5: まとめ
// ===============================================================
function addSummarySlide(pres, T, theme) {
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

  const takeaways = [
    {
      num: "01", title: `${theme}とは何か`,
      body: `${theme}は\n目標達成のための\n体系的な技術・方法です。\nまず基本概念を理解しましょう。`
    },
    {
      num: "02", title: "ステップで習得",
      body: `${theme}は\n段階的に学べます。\n最初から完璧でなくて OK。\n実践しながら改善を続けましょう。`
    },
    {
      num: "03", title: "今すぐ始めよう",
      body: `${theme}の習得は\n小さな一歩から。\n「まずやってみる」\nことが一番大切です！`
    },
  ];

  takeaways.forEach((t, i) => {
    const x = 0.35 + i * 3.15;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.2, w: 2.95, h: 3.7,
      fill: { color: T.cardBg || T.darkCard }, line: { color: T.summaryCardBorder || T.accent1, width: 1.5 }, shadow: shadow()
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.2, w: 2.95, h: 0.55,
      fill: { color: T.accent1 }, line: { type: "none" }
    });
    s.addText(t.num, {
      x, y: 1.22, w: 2.95, h: 0.51,
      fontFace: "Arial Black", fontSize: 18, bold: true,
      color: T.white, align: "center", margin: 0
    });
    s.addText(t.title, {
      x: x + 0.12, y: 1.88, w: 2.71, h: 0.85,
      fontFace: "Meiryo", fontSize: 14, bold: true,
      color: T.cardTitle || T.accent2, align: "center", valign: "middle", paraSpaceAfter: 3
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.25, y: 2.75, w: 2.45, h: 0.03,
      fill: { color: T.summaryCardBorder || T.accent1 }, line: { color: T.summaryCardBorder || T.accent1 }
    });
    s.addText(t.body, {
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
