// 5種類の色パターン定義
// 各テンプレートは T.accent1 / T.accent2 / T.dark などの意味ベースのキーで参照する
const THEMES = {
  "teal-navy": {
    label:      "ティール × ネイビー",
    accent1:    "028090",   // メインアクセント
    accent2:    "02C39A",   // サブアクセント
    dark:       "1A1A2E",   // 暗い背景・見出し
    darkCard:   "0D2137",   // まとめスライドのカード背景
    light:      "F8FAFC",   // 明るい背景
    text:       "475569",   // 本文テキスト
    white:      "FFFFFF",
    border:     "E0F2F1",   // 薄い区切り線・カード枠
    footerText: "1A1A2E",   // まとめスライドフッター文字色
  },
  "blue-white": {
    label:      "ブルー × ホワイト",
    accent1:    "1565C0",   // ミディアムブルー
    accent2:    "42A5F5",   // スカイブルー
    dark:       "0D47A1",   // ロイヤルブルー
    darkCard:   "0A2D6E",   // ディープネイビー
    light:      "FFFFFF",   // 純白（コンテンツスライド背景）
    text:       "1E2D3D",   // ディープブルーグレー（本文テキスト）
    white:      "FFFFFF",
    border:     "BBDEFB",   // ライトブルー枠
    footerText: "0D47A1",   // ロイヤルブルー（スカイブルーフッターバー上）
  },
  "orange-dark": {
    label:      "オレンジ × ダーク",
    accent1:    "E65100",   // ディープオレンジ
    accent2:    "FF9800",   // オレンジ
    dark:       "1C1C1C",   // チャコール
    darkCard:   "111111",   // ディープチャコール
    light:      "FFF8F3",   // ウォームオフホワイト
    text:       "4E342E",   // ウォームダークブラウン
    white:      "FFFFFF",
    border:     "FFE0B2",   // ライトオレンジ枠
    footerText: "1C1C1C",   // チャコール（オレンジフッターバー上・高コントラスト）
    beforeHeaderBg: "2C2C2C",  // 比較「なし/導入前」ヘッダー：チャコール
  },
  "pink-beige": {
    label:      "ピンク × ベージュ",
    accent1:    "B5737A",   // くすみローズ
    accent2:    "CF9FA3",   // ブラッシュピンク
    dark:       "4A2E35",   // ダークワインブラウン（テキスト・ステップヘッダー用）
    darkCard:   "E8D5CE",   // ライトベージュ
    light:      "FEFBF7",   // コンテンツスライド背景
    text:       "5C3D2E",   // ウォームブラウン
    white:      "FFFFFF",
    border:     "E8D0C8",   // ウォームピンク枠
    footerText: "4A2E35",   // ダークブラウン（ブラッシュフッターバー上）
    coverBg:       "FAF0E8",   // 表紙背景：アイボリー
    coverTitle:    "FAF0E8",   // 表紙タイトル：アイボリー（dark背景と同色バグを修正）
    summaryBg:     "7B4A54",   // まとめ背景：ミディアムローズ
    cardBg:        "FAF0E8",   // まとめカード背景：アイボリー
    cardBody:      "4A2E35",   // カード本文：ダークブラウン
    cardTitle:     "B5737A",   // カードタイトル：ダスティローズ
    stepAltBg:     "8A5A62",   // STEP4/FEATURE3ヘッダー：ディープローズ
    beforeHeaderBg:"9B8589",   // 比較「なし/導入前」ヘッダー：ミューテッドモーブ
  },
  "mono-gold": {
    label:      "モノトーン × ゴールド",
    accent1:    "252535",   // ディープグラファイト（帯・カードヘッダー）
    accent2:    "C09A38",   // アンティークゴールド（山吹色を排除）
    dark:       "121220",   // ほぼブラック（表紙・まとめ基底）
    darkCard:   "2C2C42",   // スレートパープル（背景から浮いて見える）
    light:      "E8E3DA",   // ウォームグレージュ（コンテンツ背景）
    text:       "383848",   // ブルーグレー
    white:      "FFFFFF",
    border:     "C09A38",   // アンティークゴールド枠
    footerText: "C09A38",   // ゴールド文字
    footerBarColor: "0D0D1A",   // 最深ダーク帯
    summaryCardBorder: "C09A38", // まとめカードのゴールド枠
    beforeHeaderBg: "1C1C30",   // 比較「導入前」ヘッダー
    coverStripeColor: "C09A38", // 表紙左ストライプ：ゴールドで存在感
    coverTagBg:   "C09A38",     // 表紙タグバッジ：ゴールド背景
    coverTagText: "0D0D1A",     // 表紙タグ文字：最深ダーク
    coverTitleFrame: "C09A38",  // タイトルエリアのゴールドフレーム
    // 表紙・まとめ：広い明暗差で立体的な奥行き
    coverGradient: [
      { position: 0,   color: "3A3A54" },   // 左上：明るいスレート紫
      { position: 35,  color: "181828" },   // 中：ダーク
      { position: 100, color: "07070D" }    // 右下：ほぼ黒
    ],
  },
};

const DEFAULT_THEME = "teal-navy";

function getTheme(key) {
  return THEMES[key] || THEMES[DEFAULT_THEME];
}
