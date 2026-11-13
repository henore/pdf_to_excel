import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import PyPDF2
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
import re
from datetime import datetime, time
import os
#PDFの内容から各種データを抽出し、データを基に氏名や出退勤日、
#時間をエクセルに自動で書き写すアプリ

#データ取得のPDFの選択と表示
def select_pdf():
    global pdf_filename
    pdf_filename = filedialog.askopenfilename(initialdir="./", title="Select a PDF File")
    pdf_label.config(text=f"選択されたPDF: {pdf_filename}")

#処理したいエクセルファイル格納フォルダ設定と表示
def select_excel_folder():
    global excel_folder
    excel_folder = filedialog.askdirectory(initialdir="./", title="Select Excel Folder")
    excel_label.config(text=f"選択されたフォルダ: {excel_folder}")

#実行しますか→はいを押下で処理スタート
def process_files():
    if not pdf_filename or not excel_folder:
        messagebox.showerror("エラー", "PDFファイルとエクセルフォルダを選択してください。")
        return

    if not messagebox.askyesno("確認", "実行しますか？"):
        return

    error_messages = []

    try:
        # PDFファイルを読み込む
        pdf_file = open(pdf_filename, 'rb')
        reader = PyPDF2.PdfReader(pdf_file)

        # ページ数を取得する
        page_count = len(reader.pages)

        # 進捗バーの設定
        progress_bar['maximum'] = page_count
        progress_bar['value'] = 0
        progress_label.config(text="処理中です...")

        # ページ毎にデータを取得する
        for page in range(page_count):

            # ページを取得する
            current_page = reader.pages[page]
            
            # ページ内のテキストデータを取得する
            content = current_page.extract_text()
            
            # 支援+10桁直後の2単語を抽出（氏名抽出）
            shien_pattern = re.compile(r'支援\d{10}\s+(\S+\s+\S+)')
            shien_matches = shien_pattern.findall(content)

            #氏名の後に数字や空白があれば削除処理
            shien_combined = ''.join(shien_matches[0].split())
            shien_combined = re.sub(r'\d', '', shien_combined)  
            
            # 年月分の単語の後にある数字最大2桁を抽出（月抽出）
            nengetsu_pattern = re.compile(r'年月分(\d{1,2})')
            nengetsu_matches = nengetsu_pattern.findall(content)
            
            # 枚中の単語の前にある数字最大2桁を抽出（年抽出）
            wareki_pattern = re.compile(r'(\d{1,2})\s*枚中')
            wareki_matches = wareki_pattern.findall(content)

            # 受給者証番号の10桁を抽出（"施設外\n支援"の直後に出現する）
            jukyusha_pattern = re.compile(r'支援(\d{10})')
            jukyusha_matches = jukyusha_pattern.findall(content)

            # 契約支給量を抽出（"はっぴぃたいむ"の前の文字列）
            keiyaku_pattern = re.compile(r'(就労継続支援\s+Ｂ型\s+.+?)\s+はっぴぃたいむ')
            keiyaku_matches = keiyaku_pattern.findall(content)

            # 日付・曜日・出退勤時間・送迎加算を一括抽出
            # 送迎加算ありの場合: 日 曜日 HH:MM HH:MM 1 1
            # 送迎加算なしの場合: 日 曜日 HH:MM HH:MM
            record_pattern = re.compile(
                r'(\d{1,2})\s*([月火水木金土])\s+'
                r'(\d{1,2}):(\d{1,2})\s+'
                r'(\d{1,2}):(\d{1,2})'
                r'(?:\s+(\d)\s+(\d))?'
            )
            records = record_pattern.findall(content)

            if not records:
                error_messages.append(f"{shien_combined}さんのpdf内に出退勤データが見つかりません")
                continue

            # PDFの和暦から西暦と月を算出
            seireki = int(wareki_matches[0]) + 2018
            month = int(nengetsu_matches[0])

            if month == 1:
                # 1月処理: まず当年ファイルがあるか確認（再処理の場合）
                excel_files = [f for f in os.listdir(excel_folder) if os.path.isfile(os.path.join(excel_folder, f)) and shien_combined in f and str(seireki) in f]

                if excel_files:
                    filename = excel_files[0]
                    filepath = os.path.join(excel_folder, filename)
                    wb = load_workbook(filepath)
                else:
                    # 前年のファイルを探す
                    prev_seireki = seireki - 1
                    prev_files = [f for f in os.listdir(excel_folder) if os.path.isfile(os.path.join(excel_folder, f)) and shien_combined in f and str(prev_seireki) in f]

                    if prev_files:
                        # 前年ファイルを新年にリネーム
                        old_filepath = os.path.join(excel_folder, prev_files[0])
                        filename = f'実績{seireki}_{shien_combined}.xlsx'
                        filepath = os.path.join(excel_folder, filename)
                        os.rename(old_filepath, filepath)
                        wb = load_workbook(filepath)

                        # 就労継続支援シート以外を全て削除
                        for sheet_name in wb.sheetnames:
                            if sheet_name != '就労継続支援':
                                del wb[sheet_name]
                    else:
                        # 前年ファイルもない新規の人
                        template_filename = os.path.join(excel_folder, '新規用.xlsx')
                        if not os.path.isfile(template_filename):
                            error_messages.append(f"{shien_combined}さんのエクセルが存在しません。新規用テンプレートも存在しません。")
                            continue
                        wb = load_workbook(template_filename)
                        wb['就労継続支援']['AD3'] = shien_combined
                        if jukyusha_matches:
                            for i, digit in enumerate(jukyusha_matches[0]):
                                wb['就労継続支援'].cell(row=3, column=9 + i, value=int(digit))
                        if keiyaku_matches:
                            wb['就労継続支援']['I5'] = keiyaku_matches[0]
                        filename = f'実績{seireki}_{shien_combined}.xlsx'
                        filepath = os.path.join(excel_folder, filename)
            else:
                # 2月～12月: 当年のファイルを探す
                excel_files = [f for f in os.listdir(excel_folder) if os.path.isfile(os.path.join(excel_folder, f)) and shien_combined in f and str(seireki) in f]

                if not excel_files:
                    # 当年ファイルがない新規の人
                    template_filename = os.path.join(excel_folder, '新規用.xlsx')
                    if not os.path.isfile(template_filename):
                        error_messages.append(f"{shien_combined}さんの{seireki}年のエクセルが存在しません。新規用テンプレートも存在しません。")
                        continue
                    wb = load_workbook(template_filename)
                    wb['就労継続支援']['AD3'] = shien_combined
                    if jukyusha_matches:
                        for i, digit in enumerate(jukyusha_matches[0]):
                            wb['就労継続支援'].cell(row=3, column=9 + i, value=int(digit))
                    if keiyaku_matches:
                        wb['就労継続支援']['I5'] = keiyaku_matches[0]
                    filename = f'実績{seireki}_{shien_combined}.xlsx'
                    filepath = os.path.join(excel_folder, filename)
                else:
                    filename = excel_files[0]
                    filepath = os.path.join(excel_folder, filename)
                    wb = load_workbook(filepath)

            # コピー元のシートを取得する
            source_sheet = wb['就労継続支援']

            # シート名の月を設定
            new_sheet_name = f"{nengetsu_matches[0]}月"

            # シートのコピーを作成する、シート名は〇月、強制的に追加
            new_sheet = wb.copy_worksheet(source_sheet)
            new_sheet.title = new_sheet_name

            # コピーしたシートを「就労継続支援」の左側に移動する
            wb._sheets.insert(wb._sheets.index(source_sheet), wb._sheets.pop())

            # 年月データを指定セルに書き込む
            new_sheet['D2'] = f"令和{wareki_matches[0]}年"
            new_sheet['L2'] = int(nengetsu_matches[0])

            # 勤務時間合計を計算するための変数
            total_working_hours = 0

            # BD2セルに値があるかチェック（休憩時間計算の例外チェック）
            has_break_exception = new_sheet['BD2'].value is not None

            # 各レコードのデータを書き込む（日付・曜日・出退勤時間・送迎加算）
            for idx, (day, weekday, sh, sm, eh, em, esc1, esc2) in enumerate(records):
                row = 11 + idx

                # 出勤日をC列、曜日をF列に書き込む
                new_sheet.cell(row=row, column=3, value=int(day))
                new_sheet.cell(row=row, column=6, value=weekday)

                # 出勤時間をN列に書き込む
                cell_n = new_sheet.cell(row=row, column=14)
                start_time = time(int(sh), int(sm))
                cell_n.value = start_time
                cell_n.number_format = 'HH:MM'

                # 退勤時間をS列に書き込む
                cell_s = new_sheet.cell(row=row, column=19)
                end_time = time(int(eh), int(em))
                cell_s.value = end_time
                cell_s.number_format = 'HH:MM'

                # 送迎加算があればX列とZ列に書き込む
                if esc1:
                    new_sheet.cell(row=row, column=24, value=int(esc1))
                if esc2:
                    new_sheet.cell(row=row, column=26, value=int(esc2))

                # 勤務時間を計算
                start_minutes = start_time.hour * 60 + start_time.minute
                end_minutes = end_time.hour * 60 + end_time.minute
                diff_minutes = end_minutes - start_minutes

                # 休憩時間の処理
                if not has_break_exception:
                    break_start = 11 * 60 + 30
                    break_end = 12 * 60 + 30
                    if start_minutes <= break_start and end_minutes >= break_end:
                        diff_minutes -= 60

                hours_diff = diff_minutes / 60
                quarter_hours = round(hours_diff * 4) / 4
                total_working_hours += quarter_hours

            # 合計勤務時間をR41に書き込む
            cell_total = new_sheet.cell(row=41, column=18)
            cell_total.value = total_working_hours
            cell_total.number_format = '0.00'
            
            # 書き込みが終わったらエクセルファイルを保存する
            wb.save(filepath)

            # 進捗バーを更新
            progress_bar['value'] += 1
            root.update_idletasks()

    #想定外エラー用
    except Exception as e:
        error_messages.append(f"エラーが発生しました: {e}")
    
    #全ての処理終了後PDFを閉じる
    finally:
        pdf_file.close()

    # 処理に失敗した人のエラー一覧表示
    if error_messages:
        for message in error_messages:
            print(message)
        messagebox.showerror("エラー", "\n".join(error_messages))
    else:
        messagebox.showinfo("完了", "処理が完了しました")
        progress_label.config(text="処理が完了しました")

# GUIの設定
root = tk.Tk()
root.title("pdf→エクセル実績表オート記入アプリ")

# 注意書き
notice = tk.Label(root, text="処理前には必ずバックアップを取り、コピーしたエクセルで処理をして下さい。")
notice.pack()
notice = tk.Label(root, text="処理をする際はpdfとエクセルの名前、両方を一致させて下さい。")
notice.pack()
notice = tk.Label(root, text="エクセル側のファイル名はスペースがある場合削除してください。")
notice.pack()
notice = tk.Label(root, text="pdf内の名前とエクセルファイルの名前に不一致やスペースがあったりすると、新規扱いになります")
notice.pack()

# PDF選択ボタン、pdf関数呼び出し
pdf_button = tk.Button(root, text="PDFファイルを選択", command=select_pdf)
pdf_button.pack()

# PDFファイルラベル
pdf_label = tk.Label(root, text="選択されたPDF: なし")
pdf_label.pack()

# エクセルフォルダ選択ボタン、エクセル用関数呼び出し
excel_button = tk.Button(root, text="エクセルフォルダを選択", command=select_excel_folder)
excel_button.pack()

# エクセルフォルダラベル
excel_label = tk.Label(root, text="選択されたフォルダ: なし")
excel_label.pack()

# 実行ボタン、実行関数呼び出して処理開始
execute_button = tk.Button(root, text="実行", command=process_files)
execute_button.pack()

# 進捗バー
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack()

# 進捗ラベル
progress_label = tk.Label(root, text="")
progress_label.pack()

# メインループ
root.mainloop()