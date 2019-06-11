# Name  
「Sheetdate_postingTool」  
  
## Overview  
    ・社内システムにログインし、指定のExcelファイルをダウンロード  
    ・ダウンロードしたファイルのシート名を変更  
    ・社内システムにログインし、変更したシートをアップロード  
  
### Description  
・予め同フォルダ内の.iniファイルに下記情報を記載、それぞれの情報を読込み起動する  
　∟社内システム用ユーザーID  
　∟社内システム用ユーザーパスワード  
　∟社内システム用ユーザーネーム  
　∟ダウンロードファイル格納先フォルダパス  
　∟ダウンロードファイルバックアップ用フォルダパス  
　∟アップロードファイル格納先フォルダパス  
　∟アップロードファイルバックアップ用フォルダパス  
　∟エラーファイルフォルダパス  
　∟ChromeDriver格納先フォルダパス  
　∟Logファイル格納先フォルダパス  

・ダウンロードしたファイルは、アップロードで問題が発生した時のため、バックアップを作成  
・アップロードでエラーが発生した場合は、エラーファイルを取得  
・ワークフローの節目ごとに.txtベースのLogに実行時間と項目を追記  

### Usage  
・.exeでの起動も可能  
