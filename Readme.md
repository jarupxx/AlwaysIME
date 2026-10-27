# AlwaysIME
<div align="center">
  <img
    src="https://github.com/user-attachments/assets/2bab68cb-b733-40d7-9427-7ff4a00c2890"
    alt="AlwaysIMEの4コマ漫画 画像生成：Canva（AI）"
    width="472"
  />
  <p>画像生成：Canva（AI）</p>
</div>

## 概要


AlwaysIME は、Windows のタスクトレイに常駐し、アプリに合わせて IME の状態を切り替えるツールです。入力方法を気にせず、日本語をそのまま入力できます。

## 主な機能

- アプリやウィンドウタイトルに応じて、IME をオン／オフ
- 一定時間入力がなければ、IME を初期状態に戻す
- 対応する Microsoft IME のスペース幅・句読点・キー割り当てを設定
- 自動起動、設定のバックアップと復元

## はじめかた

1. ZIP を、ドキュメントやデスクトップなど**書き込み可能な場所**に展開します。`C:\Program Files` は避けてください。
2. `AlwaysIME.exe` を起動します。
3. タスクトレイのアイコンを右クリックし、「設定を表示」を選びます。

`AlwaysIME.dll.config.default` は起動に必要な設定のひな形です。削除・編集しないでください。

## 設定・詳細説明

設定画面で、IME を切り替えるアプリやタイトル、未入力タイムアウト、自動起動などを設定できます。詳しくは [`Resources/AlwaysIME 説明書.md`](Resources/AlwaysIME%20説明書.md) をご覧ください。

## 動作環境と制限

- Windows 10 / 11
- WinUI / UWP などのモダン UI アプリには対応していません。
- 「MS-IME入力設定」は、Windows 11 の新しい Microsoft IME（互換モードをオフ）専用です。ATOK と Google 日本語入力には対応していません。
- 管理者権限で動くアプリを操作するには、AlwaysIME も管理者権限で起動してください。

## ランタイム

.NET デスクトップ ランタイム 8.0 が必要です。[ダウンロード](https://dotnet.microsoft.com/ja-jp/download)

## 更新・ライセンス

更新時は、新しい ZIP を同じ場所に上書き展開してください。設定ファイルはそのまま残ります。ライセンスは MIT License です（同梱の `LICENSE.txt` を参照）。

[最新版を GitHub Releases から入手](https://github.com/jarupxx/AlwaysIME/releases)
