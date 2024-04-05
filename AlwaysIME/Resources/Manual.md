# AlwaysIME

## 詳細設定

AlwaysIME.dll.config を参考に設定してください。

- AlwaysIMEMode

  IMEの制御方法を指定します。

  1：IMEが自動でオンになります。

  2：IMEを自動でオン/オフする動作を反転します。

  3：グローバル。システム全体でIMEを管理します。

  参考：WindowsではウインドウやアプリのタブごとにIMEを管理していて、どのアプリもオフから始まります。

- IsDarkMode

  ダークモード対応にします。

  on:ダーク

  off:ライト

- ImeOffList（複数登録）

  既定でIMEオフにするアプリ名を指定します。複数登録する場合はカンマ,で区切ってください。

  英語の文書の作成やデータ入力、コマンドで使うアプリ名を登録しておくと良いです。

  ```value="EXCEL,WindowsTerminal"```

- ImeOffTitle（複数登録、正規表現）

  IMEオフにするタイトルを指定します。

  拡張子を登録しておくと文書作成とプログラミングで使い分けれると思います。

  正規表現が使用できます。

  ```value="英語の文書.txt,\.py"```

- PassList（複数登録）

  無視するアプリ名を指定します。

  グローバルの時、IMEの管理から指定したアプリは除外されます。

  自身でIMEオフにするアプリを巻き込むと思ったように動かないので。

  ```value="Jane2ch"```

- intervalTime

  更新間隔をミリ秒（ms）で指定します。1000 = 1秒間隔

  ```value="500"```

- NoKeyInputTime

  キーボードやマウス未入力でIMEオンにする時間を分（min）で指定します。

- SuspendFewTime

  常駐アイコンから一時的に無効化できます。

  「少し無効」にしてから再開するまでの時間を分（min）で指定します。

- SuspendTime

  常駐アイコンから一時的に無効化できます。

  「しばらく無効」にしてから再開するまでの時間を分（min）で指定します。

- OnActivatedAppList（複数登録）

  アクティブに移動を監視するアプリ名を指定します。そのアプリを使用中はネットを遮断といった使い方ができます。

  ```value="WINWORD,MassiGra"```

- OnActivatedAppPath

  OnActivatedAppListのアプリ名と連携するアプリケーションを指定します。

  ファイアウォールを操作するのに管理者権限が必要だったので [SkipUAC](https://www.sordum.org/16219/skip-uac-prompt-v1-0/) を指定しました。

  ```value="C:\Program Files\SkipUAC\SkipUAC.exe"```

- OnActivatedArgv

  上記で指定したアプリケーションの引数を指定します。不要なら```value=""```にして下さい。

  ```value="/ID foo"```

- BackgroundAppList（複数登録）

  監視中のアプリが非アクティブになってから数分後に次で指定するアプリが呼び出されますが、違うアプリをトリガーにしたい場合はアプリ名を指定して下さい。

  ```value="chrome,msedge"```

- BackgroundAppPath

  BackgroundAppListのアプリ名と連携するアプリケーションを指定します。

  ```value="C:\Program Files\SkipUAC\SkipUAC.exe"```

- BackgroundArgv

  上記で指定したアプリケーションの引数を指定します。不要なら```value=""```にして下さい。

  ```value="/ID bar"```

- DelayBackgroundTime

  非アクティブになってから連携するアプリケーションを呼び出すまでの遅延を分（min）で指定します。

  ```value="3"```

- Punctuation（トグル）

  メニューから句読点を切り替えられるようにします。「初期値 ⇔ 切替」をカンマ,で区切ります。不要なら```value=""```にして下さい。

  この機能はレジストリ ```HKEY_CURRENT_USER\Software\Microsoft\IME\15.0\IMEJP\MSIME "option1" ```を書き換えます。

  0：カンマ・ピリオド「，．」

  1：読点・句点「、。」

  2：読点・ピリオド「、．」

  3：カンマ・句点「，。」

  ```value="0,3"```

- SpaceWidth（トグル）

  メニューからスペースを切り替えられるようにします。「初期値 ⇔ 切替」をカンマ,で区切ります。不要なら```value=""```にして下さい。

  この機能はレジストリ ```HKEY_CURRENT_USERHKEY_CURRENT_USER\Software\Microsoft\IME\15.0\IMEJP\MSIME "InputSpace" ```を書き換えます。

  0：現在の入力モード

  1：常に全角

  2：常に半角

  ```value="2,1"```

- AppExit

  メニューの常駐の終了の位置を移動します。

  0：中ほど

  1：一番下

## 仕様

  常駐型。

  アクティブウインドウを監視してIMEの制御を行い、同時に連携するアプリを呼ぶことができます。

  タスクマネージャーには効きません。
