# Excelガントチャートジェネレーター (VBA)

[![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/ja-jp/microsoft-365/excel)
[![VBA](https://img.shields.io/badge/VBA-777BB4?style=for-the-badge&logo=visual-basic-for-applications&logoColor=white)](https://docs.microsoft.com/ja-jp/office/vba/api/overview/excel)

**これは、VBAを用いて開発された、動的でインタラクティブな高機能ガントチャート生成ツールです。**

タスクリストをExcelシートに入力するだけで、本格的なガントチャートを自動で描画します。単なるテンプレートとは異なり、柔軟な設定と高度なプログラミング技術を駆使して、実用的なプロジェクト管理を実現します。

<img width="1789" height="657" alt="Image" src="https://github.com/user-attachments/assets/06ad8e86-3452-4b75-bcfc-b850ff9d5126" />

## 主な機能

*   **動的なガントチャート描画**: タスクシートのデータに基づき、自動でチャートを生成・更新します。
*   **インタラクティブな操作**: チャート上のタスクバーをクリックすることで、タスク詳細を表示します。

## GanttChartのダウンロード方法

このツールを使用するために、**`GanttChart.xlsm`というExcelファイルのみ**が必要です。以下の手順でダウンロードしてください。

1.  このGitHubページの上部にあるファイルリストから、`GanttChart.xlsm` をクリックします。
2.  次の画面で、右側にある「Download」ボタンをクリックします。
   <img width="1913" height="680" alt="Image" src="https://github.com/user-attachments/assets/c14e0265-d32a-47b6-8eec-f95c60e8fb9f" />

**【重要】**他のフォルダ（`vba-files`など）や設定ファイル（`.gitignore`など）は、開発用のファイルです。**ツールを利用するだけであれば、これらのファイルをダウンロードする必要はありません。**

## ツールの魅力・技術的な特徴

このプロジェクトは、Excel VBAの能力を最大限に引き出すための、モダンな開発アプローチを取り入れています。

*   **オブジェクト指向設計**:
    `GanttChart`, `Tasks`, `Settings` などの役割ごとにVBAクラスモジュールを設計し、処理をカプセル化。これにより、コードの可読性が向上し、機能追加や変更が容易な、スケールしやすいアーキテクチャを実現しています。

*   **イベント駆動プログラミング**:
    `M_ChartEvents` モジュールに見られるように、将来的にユーザーのチャート操作（クリックなど）をトリガーとして特定の処理を実行するための基盤が用意されています。これにより、静的なチャート表示に留まらない、インタラクティブなツールの構築が可能です。

*   **独立した描画エンジン**:
    チャートの描画ロジック (`M_Dlaw.bas`) を他のビジネスロジックから分離。これにより、デザインの変更や描画方法の最適化が、他の部分に影響を与えることなく行えます。

## 使い方

1.  `GanttChart.xlsm` ファイルを開きます。
2.  セキュリティ警告が表示された場合は、「コンテンツの有効化」をクリックしてマクロを有効にしてください。
3.  `Tasks` という名前のシートに、所定のフォーマットに従ってタスク情報を入力します。
    *   (例: B列にタスク名, ...)
      <img width="669" height="278" alt="Image" src="https://github.com/user-attachments/assets/2b6988ce-ff40-4e58-bcc1-33e09c2e447f" />
4.  「IDと終了日を入力」ボタンをクリックすると、`ID`と`終了日`が自動で入力され、罫線が更新され、列幅が自動調整されます。
      <img width="674" height="277" alt="Image" src="https://github.com/user-attachments/assets/95b22220-d126-4bf1-9e01-b0f813fdc84e" />
6.  `GanttChart` という名前のシートに移動し、「チャート更新」ボタンをクリックすると、最新のタスク情報に基づいたチャートが `GanttSheet` に描画されます。
7.  チャート上のタスクバーをクリックすることで、タスク詳細が表示されます。
      <img width="1789" height="657" alt="Image" src="https://github.com/user-attachments/assets/06ad8e86-3452-4b75-bcfc-b850ff9d5126" />

## 前提条件

*   Microsoft Excel(Windows版)
*   マクロの実行が許可されていること
