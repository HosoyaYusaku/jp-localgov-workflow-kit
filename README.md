# jp-localgov-workflow-kit

A toolkit for Japanese local government officials to automate the generation of business process flowcharts and LLM-ready files from a simple Excel list. All processes are completed within Microsoft Excel using VBA.

This tool is designed to assist in visualizing, analyzing, and improving the day-to-day operations ("Gyomu") in Japanese municipalities.

## ✨ Features

-   **Automated Flowchart Generation:** Automatically creates a professional swimlane flowchart based on BPMN (Business Process Model and Notation) from a task list in Excel.
-   **LLM-Ready Export:** Generates Markdown files with a single click, perfectly formatted for creating business manuals or performing issue analysis with Large Language Models (LLMs).
-   **Tailored for Local Governments:** The swimlanes and parameters are designed with typical municipal roles (e.g., residents, full-time staff, contract staff) in mind.
-   **Easy to Implement & Customize:** Simply copy and paste the VBA code. Easily adjust the flowchart's appearance and LLM prompts by editing constants within the code.

## 🚀 How to Use

Follow these steps to set up and use the toolkit.

1.  **Download the Repository:** Download the project files as a ZIP or clone the repository.
2.  **Prepare the Excel File:**
    * Create a new macro-enabled Excel workbook (`.xlsm`).
    * Open the Visual Basic Editor (VBE) by pressing `Alt + F11`.
3.  **Import VBA Modules:**
    * In the VBE, right-click on the project explorer and select `Import File...`.
    * Import both `.bas` files located in the [`vba_modules`](./vba_modules/) folder:
        * `FlowchartGenerator.bas`
        * `LLMExporter.bas`
4.  **Prepare the Data Sheet:**
    * Copy the contents from the sample file (`sample_data/Business_Process_List.xlsx`).
    * Paste the data into a new sheet in your `.xlsm` file and **rename the sheet to `業務リスト`**.
5.  **Run the Macros:**
    * Press `Alt + F8` to open the Macro dialog.
    * Run `CreateFlowChart` to generate the business flowchart and a legend sheet.
    * Run `ExportMarkdownForLLM` to export two Markdown files to your desktop for LLM collaboration.

## 📊 Output Files

This toolkit generates the following files. You can see the actual output examples in the [`output_examples`](./output_examples/) folder.

### 1. Business Flowchart (`業務フロー図` Sheet)

A visual representation of your business process is automatically generated in a new sheet.

<img width="894" alt="flow" src="https://github.com/user-attachments/assets/bf65c744-8460-4adf-9bf3-b9bed5ee2b53" />

### 2. Markdown for LLM Collaboration

Generates files formatted with prompts for creating manuals and analyzing issues.

- See: [`1_generated_for_manual.md`](./output_examples/1_generated_for_manual.md) and [`2_generated_for_analysis.md`](./output_examples/2_generated_for_analysis.md)

### 3. Final Output by LLM

By providing the generated Markdown to an LLM, you can obtain high-quality documents like the ones below.

- See: [`3_llm_output_manual.txt`](./output_examples/3_llm_output_manual.txt) and [`4_llm_output_analysis.txt`](./output_examples/4_llm_output_analysis.txt)

## 🔧 Customization

You can easily customize the tool by editing the constants at the top of each VBA module.

-   **In `FlowchartGenerator.bas`:** Change sheet names, layout settings, colors, and connector styles.
-   **In `LLMExporter.bas`:** Modify the prompt text to tailor the LLM's output to your specific needs.

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---
<br>

# 自治体業務フロー自動化キット (jp-localgov-workflow-kit)

日本の自治体職員向けに、Excelのリストから業務フロー図とLLM連携用ファイルを自動生成するツールキットです。すべての処理はVBAを使用し、Microsoft Excel内で完結します。

このツールは、日本の自治体における日々の業務を可視化し、分析・改善するための一助となることを目指して開発されました。

## ✨ 主な機能

-   **フロー図の自動生成:** Excelのタスクリストから、BPMN（ビジネスプロセスモデリング表記法）を参考にした本格的な業務フロー図を自動で描画します。
-   **LLM連携ファイル出力:** ワンクリックで、業務マニュアル作成や課題分析に適した形式のMarkdownファイルを生成します。
-   **自治体業務に特化:** レーンの役割（例：利用者、常勤職員、会計年度任用職員）は、自治体特有の業務環境を想定して設計されています。
-   **簡単な導入とカスタマイズ:** VBAコードをコピー＆ペーストするだけで導入完了。コード内の定数を編集するだけで、フロー図の見た目やLLMへの指示を簡単に調整できます。

## 🚀 使い方

以下の手順でツールをセットアップして使用します。

1.  **リポジトリのダウンロード:** プロジェクトファイルをZIPでダウンロードするか、リポジトリをクローンします。
2.  **Excelファイルの準備:**
    * 新規にマクロ有効ブック（`.xlsm`形式）を作成します。
    * `Alt + F11`キーを押して、VBE（Visual Basic Editor）を開きます。
3.  **VBAモジュールのインポート:**
    * VBEのプロジェクトエクスプローラー上で右クリックし、「ファイルのインポート」を選択します。
    * [`vba_modules`](./vba_modules/)フォルダにある2つの`.bas`ファイルをインポートします。
        * `FlowchartGenerator.bas`
        * `LLMExporter.bas`
4.  **データシートの準備:**
    * サンプルファイル（`sample_data/Business_Process_List.xlsx`）の中身をコピーします。
    * 作成したマクロ有効ブックに新しいシートを追加し、データを貼り付けた後、**シート名を「業務リスト」に変更します。**
5.  **マクロの実行:**
    * `Alt + F8`キーを押し、マクロのダイアログを開きます。
    * `CreateFlowChart` を実行すると、業務フロー図と凡例シートが生成されます。
    * `ExportMarkdownForLLM` を実行すると、LLM連携用のMarkdownファイル2点がデスクトップに出力されます。

## 📊 生成されるファイル

本ツールキットは以下のファイルを生成します。実際の出力サンプルは[`output_examples`](./output_examples/)フォルダでご確認いただけます。

### 1. 業務フロー図（「業務フロー図」シート）

業務プロセスの流れを可視化した図が、新しいシートに自動で作成されます。

<img width="894" alt="flow" src="https://github.com/user-attachments/assets/bf65c744-8460-4adf-9bf3-b9bed5ee2b53" />

### 2. LLM連携用Markdown

マニュアル作成用と課題分析用のプロンプトが埋め込まれたファイルを生成します。

- 参照: [`1_generated_for_manual.md`](./output_examples/1_generated_for_manual.md) および [`2_generated_for_analysis.md`](./output_examples/2_generated_for_analysis.md)

### 3. LLMによる最終成果物

生成されたMarkdownをLLMに与えることで、以下のような質の高いドキュメントが得られます。

- 参照: [`3_llm_output_manual.txt`](./output_examples/3_llm_output_manual.txt) および [`4_llm_output_analysis.txt`](./output_examples/4_llm_output_analysis.txt)

## 🔧 カスタマイズ

各VBAモジュールの冒頭にある定数（`Const`）を編集することで、簡単にツールをカスタマイズできます。

-   **`FlowchartGenerator.bas`内:** シート名、レイアウト設定、色、コネクタの種類などを変更できます。
-   **`LLMExporter.bas`内:** プロンプト文章を自由に変更し、LLMの応答をあなたの目的に合わせて調整できます。

## 📄 ライセンス

このプロジェクトはMITライセンスです。詳細は[LICENSE](LICENSE)ファイルをご覧ください。
