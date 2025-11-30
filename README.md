# embryology_data_analysis_vba_demo
Excel/VBA を用いて胚培養データを自動集計するためのデモ用ファイルです。
embryology_data_analysis（デモ版）

このリポジトリは、胚培養・ART データの集計を自動化するための Excel/VBA ツールのデモ版です。
院内情報・患者情報はすべて削除し、匿名化したデータのみを含みます。

■ 概要

Excel（Power Query / Power Pivot を含む）を利用した
胚培養関連データの自動集計システム

採卵数、受精率、胚盤胞率、妊娠関連指標など
主要 KPI の集計を自動化

院別・年月別のサマリーシートを ワンクリック生成

複数のピボットテーブルから値を取得し、指定フォーマットへ出力

本リポジトリは 動作イメージの公開・ポートフォリオ用途を目的としたデモ版です。

■ 主な機能（デモ版）

Power Query によるデータ取り込み

Power Pivot モデルによる統合集計

VBA によるレポート自動生成

複数クリニック・複数年度の横断集計処理

KPI（受精率・胚盤胞到達率・ET実施数・妊娠関連指標など）を
年月別・施設別に自動出力

■ 構成ファイル

embryology_data_analysis_demo.xlsm
→ デモ用 Excel（匿名化データ・VBA 含む）

src/
→ VBA モジュールを .bas 形式でエクスポートしたもの
（コード閲覧用）

■ 動作環境

Microsoft Excel

Power Query

Power Pivot

VBA（マクロ有効）

■ 注意事項

このリポジトリは デモ版であり、実データは一切含まれていません。

ファイルの再利用・再配布は許可していません。（無ライセンス）

内容はポートフォリオ目的で公開しており、業務仕様とは異なる場合があります。

This repository is for portfolio/demonstration purposes only.
Reuse or redistribution of the files is not permitted.
