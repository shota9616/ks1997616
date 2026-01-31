---
marp: true
theme: default
paginate: true
style: |
  section {
    font-family: 'Hiragino Sans', 'Noto Sans JP', 'Yu Gothic', sans-serif;
    background-color: #F7FAFC;
    color: #2D3748;
    padding: 60px 80px;
    line-height: 1.6;
  }
  h1 {
    color: #1A365D;
    font-size: 2.2em;
    font-weight: 700;
    border-bottom: 3px solid #2B6CB0;
    padding-bottom: 16px;
    margin-bottom: 40px;
  }
  h2 {
    color: #2B6CB0;
    font-size: 1.6em;
    font-weight: 600;
    margin-bottom: 24px;
  }
  ul {
    line-height: 2.0;
    font-size: 1.1em;
  }
  li {
    margin-bottom: 16px;
  }
  strong {
    color: #1A365D;
    font-weight: 700;
  }
  em {
    color: #DD6B20;
    font-style: normal;
    font-weight: 600;
  }
  code {
    background: #EDF2F7;
    padding: 8px 16px;
    border-radius: 6px;
    font-size: 0.95em;
    color: #2D3748;
  }
  blockquote {
    border-left: 4px solid #38A169;
    padding-left: 24px;
    margin: 24px 0;
    color: #4A5568;
    font-size: 1.1em;
  }
  section.lead {
    display: flex;
    flex-direction: column;
    justify-content: center;
    text-align: center;
    background: linear-gradient(135deg, #1A365D 0%, #2B6CB0 100%);
    color: #ffffff;
  }
  section.lead h1 {
    color: #ffffff;
    border-bottom: 3px solid #38A169;
    font-size: 2.8em;
  }
  section.lead h2 {
    color: #E2E8F0;
    font-weight: 400;
  }
  section.invert {
    background-color: #1A365D;
    color: #ffffff;
  }
  section.invert h1 {
    color: #ffffff;
    border-bottom-color: #38A169;
  }
  section.invert strong {
    color: #38A169;
  }
  .step-number {
    font-size: 4em;
    font-weight: 800;
    color: #2B6CB0;
    opacity: 0.2;
    position: absolute;
    right: 80px;
    top: 60px;
  }
  .highlight-box {
    background: #EBF8FF;
    border-left: 4px solid #2B6CB0;
    padding: 20px 24px;
    border-radius: 0 8px 8px 0;
    margin: 24px 0;
  }
  .warning-box {
    background: #FFFAF0;
    border-left: 4px solid #DD6B20;
    padding: 20px 24px;
    border-radius: 0 8px 8px 0;
    margin: 24px 0;
  }
  .success-box {
    background: #F0FFF4;
    border-left: 4px solid #38A169;
    padding: 20px 24px;
    border-radius: 0 8px 8px 0;
    margin: 24px 0;
  }
  footer {
    color: #A0AEC0;
    font-size: 0.7em;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 24px 0;
    font-size: 1em;
  }
  th {
    background: #1A365D;
    color: #ffffff;
    padding: 12px 16px;
    text-align: left;
    font-weight: 600;
  }
  td {
    padding: 12px 16px;
    border-bottom: 1px solid #E2E8F0;
  }
  tr:nth-child(even) {
    background: #EDF2F7;
  }
  ol {
    line-height: 2.0;
    font-size: 1.1em;
  }
  ol li {
    margin-bottom: 12px;
  }
---

<!-- _class: lead -->

# Week1「AIと出会う」

## AI活用実践研修（12週間）

若杉様カフェスタッフ向け

---

# 今日のゴール

<div class="success-box">

**この60分で達成すること**

</div>

- **AIに3回質問**できるようになる
- 「質問 → 回答 → 修正」のサイクルを体験する

---

# 自己紹介タイム

それぞれ教えてください

- **お名前**
- **担当している業務**
- **AIへの印象**（一言で）

---

# 12週間の全体像

| 週 | テーマ |
|:---:|:---|
| Week 1-2 | AIと出会う・業務を整理する |
| Week 3-6 | 接客・ドリンク・フードのマニュアル作成 |
| Week 7-10 | マニュアルの改善・仕上げ |
| Week 11-12 | 教育AIボットを作る |

<div class="highlight-box">

**最終ゴール：マニュアル3本 ＋ 教育AIボット**

</div>

---

<!-- _class: lead -->

# AIとは？

---

# AIってなに？

<div class="highlight-box">

**すごく賢いチャット相手**

</div>

- 質問すると、すぐに答えてくれる
- 何度でも聞き直せる
- 怒られない、待たされない

---

# AIができること・できないこと

| できること | できないこと |
|:---|:---|
| 文章を作る・直す | お店の今日の売上を見る |
| アイデアを出す | コーヒーを淹れる |
| 質問に答える | お客様の顔を覚える |
| 言い換えを提案する | 100%正しい答えを出す |

---

<!-- _class: lead -->

# セキュリティルール

---

<!-- _class: invert -->

# 絶対に入力しないこと

## *3つのNG*

- **お客様の名前・連絡先**
- **パスワード・暗証番号**
- **売上などの機密情報**

---

# AIを使うときの約束

<div class="warning-box">

**AIの回答は必ず確認する**

</div>

- AIは間違えることがある
- 「本当かな？」と疑う習慣を
- 困ったら周りの人に相談

---

<!-- _class: lead -->

# 実践ワーク

AIに3回質問してみよう

---

# 実践ワークの流れ

<div class="step-number">3</div>

今日は **3つの質問** をします

1. **挨拶する** — まずはAIと話してみる
2. **業務の質問** — カフェの仕事について聞く
3. **修正を依頼** — 回答を直してもらう

---

# 質問① AIに挨拶する

<div class="step-number">1</div>

## やること

AIに話しかけて、返事をもらう

<div class="highlight-box">

**入力例**
「こんにちは。私はカフェで働いています。よろしくお願いします。」

</div>

---

# 質問① やってみましょう

## 手順

1. AIチャット画面を開く
2. 挨拶を入力する
3. 送信ボタンを押す
4. 返事を読む

---

# 質問② 業務について聞く

<div class="step-number">2</div>

## やること

カフェの仕事について質問する

<div class="highlight-box">

**入力例**
「カフェで新人スタッフに教えるべきことを3つ教えてください」

</div>

---

# 質問② やってみましょう

## 他の質問例

- 「お客様への挨拶の例を教えて」
- 「忙しい時間帯の優先順位の決め方は？」
- 「レジ対応で気をつけることは？」

---

# 質問③ 回答を修正させる

<div class="step-number">3</div>

## やること

AIの回答を、もっと良くする指示を出す

<div class="highlight-box">

**入力例**
「もっと短くまとめてください」
「カフェ初心者向けに書き直してください」

</div>

---

# 質問③ やってみましょう

## 修正指示の例

- 「箇条書きにしてください」
- 「もっと具体的に教えてください」
- 「敬語を使った言い方にしてください」

---

<!-- _class: invert -->

# 今週の宿題

---

# 宿題

<div class="highlight-box">

**「人によってやり方が違う業務」を3つ書き出す**

</div>

## 例
- ドリンクの作り方
- お客様への声かけのタイミング
- 片付けの順番

**提出**：次回のZoom研修で発表（気軽にでOK）

---

# 次回予告

## Week2「業務を棚卸しする」

- 今日書き出した業務を深掘り
- マニュアル化する業務を決める

---

<!-- _class: lead -->

# お疲れ様でした

今日できたこと：AIに質問して、回答を修正できた

来週もお楽しみに！

---
