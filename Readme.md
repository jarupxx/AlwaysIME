# AlwaysIME

## 概要
IMEオンにするツール

## 使い方
起動するだけです。

アプリを行き来するとIMEが勝手に入力モードを切り替えてしまいますが、
すかさず入力モードを元に戻します。

IMEの設定はグローバルで、自分が切り替えた状態でいつもいて欲しい。そんな願いを叶えます。

## 仕様
常駐型。

アクティブウインドウを監視してIMEの制御を行います。同時に連携するアプリを呼ぶことができます。

自身でIMEオフにするアプリに行くと状態がおかしくなります。
そういった相性が悪いアプリは```AlwaysIME.exe.config```の```"AppList"```に登録してください。

タスクマネージャーには効きません。
