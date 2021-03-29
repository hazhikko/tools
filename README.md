# tools
ツールいろいろごった煮リポジトリ
## VBScript
### [makeLink.vbs](VBScript/makeLink.vbs)
ブログ：[相対パスでショートカットを作るコマンド作った](https://nekodeki.com/%e7%9b%b8%e5%af%be%e3%83%91%e3%82%b9%e3%81%a7%e3%82%b7%e3%83%a7%e3%83%bc%e3%83%88%e3%82%ab%e3%83%83%e3%83%88%e3%82%92%e4%bd%9c%e3%82%8b%e3%82%b3%e3%83%9e%e3%83%b3%e3%83%89%e4%bd%9c%e3%81%a3%e3%81%9f/)
#### 機能
ショートカットを相対パス指定で作成する
#### 使い方
1. 任意の場所に置く
2. 引数を指定して実行する  
    - 指定するパスは絶対/相対のどちらでも可
```shell
makeLink.vbs 対象ファイルのパス ショートカットを置く場所のパス
```
パスを通した場合はコマンド的な使い方ができる
```shell
makeLink 対象ファイルのパス ショートカットを置く場所のパス
```
#### 備考
ショートカットの対象となるファイルの拡張子がなんであってもショートカットのアイコンはフォルダになる  
※Windowsの仕様  
  
相対パスで作成しているため、作成したショートカットを移動すると動かなくなるので注意
## PowerShell
### [cursorCentering.ps1](PowerShell/cursorCentering.ps1)
ブログ：[PowerShellでマウスカーソルを画面中央に表示する](https://nekodeki.com/powershell%e3%81%a7%e3%83%9e%e3%82%a6%e3%82%b9%e3%82%ab%e3%83%bc%e3%82%bd%e3%83%ab%e3%82%92%e7%94%bb%e9%9d%a2%e4%b8%ad%e5%a4%ae%e3%81%ab%e8%a1%a8%e7%a4%ba%e3%81%99%e3%82%8b/)
#### 機能
ショートカットキーでマルチディスプレイの任意画面中央にカーソルを移動させる
#### 使い方
1. 任意の場所に置く
1. ショートカットキーツールに設定する
    - 引数にモニターの番号を指定する(2台あるなら1、2をそれぞれ指定したショートカットキーを作成)
    - [HotkeyPに設定する場合はこちら](https://nekodeki.com/powershell%e3%81%a7%e3%83%9e%e3%82%a6%e3%82%b9%e3%82%ab%e3%83%bc%e3%82%bd%e3%83%ab%e3%82%92%e7%94%bb%e9%9d%a2%e4%b8%ad%e5%a4%ae%e3%81%ab%e8%a1%a8%e7%a4%ba%e3%81%99%e3%82%8b/)
1. 実際にショートカットキーを押して動きを確認し、引数の数字を調整
#### 備考
PowerShellやGUIで引数となるモニターの番号を取得できないようなので、実際に動かして割り当てられている数字を確認する必要がある  
詳細はブログ参照