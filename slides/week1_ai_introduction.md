---
marp: true
theme: default
paginate: true
size: 16:9
style: |
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap');

  :root {
    --navy: #0F172A;
    --blue: #3B82F6;
    --slate: #64748B;
    --light: #F1F5F9;
    --green: #10B981;
    --amber: #F59E0B;
  }

  section {
    font-family: 'Noto Sans JP', 'Hiragino Sans', sans-serif;
    background: #FFFFFF;
    color: var(--navy);
    padding: 70px 90px;
    justify-content: flex-start;
    line-height: 1.8;
  }

  h1 {
    color: var(--navy);
    font-size: 44px;
    font-weight: 700;
    margin: 0 0 40px 0;
    padding: 0;
    border: none;
    letter-spacing: -0.02em;
  }

  h2 {
    color: var(--blue);
    font-size: 20px;
    font-weight: 600;
    margin: 0 0 32px 0;
    text-transform: uppercase;
    letter-spacing: 0.15em;
  }

  p {
    font-size: 24px;
    line-height: 1.8;
    color: var(--slate);
    margin: 0 0 20px 0;
  }

  ul {
    margin: 0;
    padding: 0;
    list-style: none;
  }

  li {
    font-size: 26px;
    line-height: 1.7;
    color: var(--slate);
    padding-left: 36px;
    margin-bottom: 24px;
    position: relative;
  }

  li::before {
    content: "";
    position: absolute;
    left: 0;
    top: 14px;
    width: 10px;
    height: 10px;
    background: var(--blue);
    border-radius: 50%;
  }

  strong {
    color: var(--navy);
    font-weight: 700;
  }

  em {
    font-style: normal;
    color: var(--amber);
    font-weight: 600;
  }

  code {
    background: var(--light);
    padding: 6px 16px;
    border-radius: 8px;
    font-size: 22px;
    font-family: 'Noto Sans JP', sans-serif;
    color: var(--navy);
  }

  blockquote {
    background: var(--light);
    border-left: 5px solid var(--blue);
    margin: 32px 0;
    padding: 28px 36px;
    border-radius: 0 12px 12px 0;
  }

  blockquote p {
    margin: 0;
    font-size: 24px;
    color: var(--navy);
  }

  section.cover {
    background: linear-gradient(145deg, #0F172A 0%, #1E3A5F 100%);
    justify-content: center;
    align-items: center;
    text-align: center;
    padding: 90px;
  }

  section.cover h1 {
    color: #FFFFFF;
    font-size: 64px;
    margin-bottom: 20px;
    letter-spacing: -0.02em;
  }

  section.cover h2 {
    color: #94A3B8;
    font-size: 28px;
    font-weight: 400;
    letter-spacing: 0.05em;
    text-transform: none;
    margin: 0;
  }

  section.cover p {
    color: #64748B;
    font-size: 20px;
    margin-top: 48px;
  }

  section.divider {
    background: var(--light);
    justify-content: center;
    align-items: flex-start;
    padding: 90px;
  }

  section.divider h2 {
    color: var(--blue);
    font-size: 18px;
    margin-bottom: 20px;
  }

  section.divider h1 {
    font-size: 52px;
    margin-bottom: 20px;
  }

  section.divider p {
    font-size: 26px;
  }

  section.accent {
    background: var(--navy);
    justify-content: center;
    align-items: center;
    text-align: center;
  }

  section.accent h1 {
    color: #FFFFFF;
    font-size: 56px;
  }

  section.accent p {
    color: #94A3B8;
    font-size: 28px;
  }

  section.accent li {
    color: #E2E8F0;
    font-size: 32px;
    text-align: left;
  }

  section.accent li::before {
    background: var(--green);
  }

  section.work {
    background: linear-gradient(145deg, #3B82F6 0%, #2563EB 100%);
  }

  section.work h1 {
    color: #FFFFFF;
  }

  section.work h2 {
    color: rgba(255,255,255,0.7);
  }

  section.work p, section.work li {
    color: #FFFFFF;
  }

  section.work li::before {
    background: #FFFFFF;
  }

  section.work blockquote {
    background: rgba(255,255,255,0.15);
    border-left-color: #FFFFFF;
  }

  section.work blockquote p {
    color: #FFFFFF;
  }

  .two-col {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 60px;
    margin-top: 20px;
  }

  .badge {
    display: inline-block;
    background: var(--blue);
    color: #FFFFFF;
    font-size: 14px;
    font-weight: 600;
    padding: 6px 16px;
    border-radius: 20px;
    margin-bottom: 16px;
  }

  .num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 48px;
    height: 48px;
    background: var(--blue);
    color: #FFFFFF;
    font-size: 24px;
    font-weight: 700;
    border-radius: 50%;
    margin-right: 20px;
  }

  .highlight {
    background: linear-gradient(transparent 60%, #DBEAFE 60%);
    padding: 0 4px;
  }

  .warning {
    background: #FEF3C7;
    border-left: 5px solid var(--amber);
    padding: 24px 32px;
    border-radius: 0 12px 12px 0;
    margin: 24px 0;
  }

  .success {
    background: #D1FAE5;
    border-left: 5px solid var(--green);
    padding: 24px 32px;
    border-radius: 0 12px 12px 0;
    margin: 24px 0;
  }
---

<!-- _class: cover -->
<!-- _paginate: false -->

## AI活用実践研修

# Week 1
# AIと出会う

12週間プログラム

---

<!-- _class: divider -->

## TODAY'S GOAL

# 今日のゴール

60分後、AIと会話できるようになる

---

## GOAL

# 60分後のあなた

- AIに**質問**ができる
- AIの回答を**修正**させられる
- 「思ったより**簡単**」と感じている

---

## INTRODUCTION

# まずは自己紹介

- お名前
- 担当している業務
- AI、使ったことありますか？

---

## OVERVIEW

# 12週間で手に入れるもの

- 業務マニュアル **3本**
- 新人教育AIボット **1つ**

---

## 12 WEEKS JOURNEY

# 12週間のステップ

<div class="two-col">
<div>

**前半：基礎**
- Week 1：AIと出会う
- Week 2：業務の棚卸し
- Week 3：マニュアル作成①
- Week 4：マニュアル作成②
- Week 5：マニュアル作成③
- Week 6：中間ふりかえり

</div>
<div>

**後半：応用**
- Week 7：AIボット設計
- Week 8：ボット構築①
- Week 9：ボット構築②
- Week 10：テスト・改善
- Week 11：運用準備
- Week 12：最終発表

</div>
</div>

---

<!-- _class: divider -->

## WHAT IS AI?

# AIって何？

難しく考えなくて大丈夫です

---

## WHAT IS AI?

# 24時間いる、物知りな先輩

テキストで質問すると、テキストで答えてくれる。

ただし、**必ず正しいとは限らない**。

---

## AI CAN

# AIにできること

- 文章を書く・要約する
- アイデアを出す
- 質問に答える
- 手順を整理する

---

## AI CAN'T

# AIにできないこと

- *最新情報*を調べる（学習が古い）
- *100%正確*な回答をする
- *あなたの判断*を代わりにする

> だから、AIの回答は必ず確認しましょう

---

<!-- _class: accent -->

# 3つのルール

これだけ守ればOK

---

## RULES

# AIを使うときのルール

- **個人情報を入れない**（名前、住所、電話番号）
- **お客様の情報を入れない**
- **回答は必ず確認する**（間違いもある）

---

<!-- _class: work -->

## HANDS-ON

# 実践ワーク

AIに3回質問してみよう

⏱ 25分

---

## STEP 1

# まずは挨拶してみる

やること：AIに自己紹介する

> 「こんにちは。私はカフェで働いています。」

→ AIが返事をしてくれることを確認

---

## STEP 2

# 仕事のことを聞いてみる

やること：業務に関する質問をする

- 「カフェラテの美味しい作り方は？」
- 「お客様を待たせないコツは？」
- 「レジ締めの手順を教えて」

---

## STEP 3

# 回答を修正させる

やること：さっきの回答を修正してもらう

- 「もっと**短く**まとめて」
- 「**箇条書き**で整理して」
- 「**新人向け**に書き直して」

---

## HOMEWORK

# 今週の宿題

**「人によってやり方が違う」業務を3つ書き出す**

- 所要時間：10分
- 提出方法：LINEに送信
- 締切：次回研修の前日まで

---

<!-- _class: cover -->
<!-- _paginate: false -->

# お疲れ様でした

## 次回：Week 2「業務の棚卸し」
