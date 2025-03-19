import re
from docx import Document

# 数字を漢数字に変換するための関数（最大9999まで対応）
def convert_number_to_kanji(num_str):
    n = int(num_str)
    if n == 0:
        return "零"
    result = ""
    # 単位（1,10,100,1000）
    units = ["", "十", "百", "千"]
    digits = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    # 数字を逆順に処理（下位桁から上位桁へ）
    num_str_rev = num_str[::-1]
    for i, ch in enumerate(num_str_rev):
        digit = int(ch)
        if digit != 0:
            if i > 0 and digit == 1:
                # 十位以上で「1」は省略（例：11→「十一」となる）
                result = units[i] + result
            else:
                result = digits[digit] + units[i] + result
    return result

def convert_numbers_in_text(text):
    """テキスト中の数字（連続した数字列）を漢数字に変換する。"""
    return re.sub(r'\d+', lambda m: convert_number_to_kanji(m.group(0)), text)

def convert_ascii_to_fullwidth(text):
    """半角英文字を全角英文字に変換する。"""
    result = ""
    for char in text:
        if ('A' <= char <= 'Z') or ('a' <= char <= 'z'):
            result += chr(ord(char) + 0xFEE0)
        else:
            result += char
    return result

def insert_space_after_punctuation(text):
    """
    「！」または「？」の直後で、次の文字が全角スペース、閉じ丸括弧、または閉じかぎかっこでない場合、
    全角スペースを挿入する。
    """
    return re.sub(r'([！？])(?![　）」])', r'\1　', text)

# Wordファイルを読み込み（適宜パスを変更してください）
doc = Document("/Users/a104/Desktop/input.docx")

for paragraph in doc.paragraphs:
    # 段落の先頭が「「」または「（」で始まらない場合は、先頭に全角スペースを追加（字下げ）
    if not (paragraph.text.startswith("「") or paragraph.text.startswith("（")):
        paragraph.text = "　" + paragraph.text

    # 各Runごとにテキスト変換を実施
    for run in paragraph.runs:
        # 数字を漢数字に変換（複数桁対応）
        run.text = convert_numbers_in_text(run.text)
        # 「…」を「……」に置換
        run.text = run.text.replace("…", "……")
        # 半角英文字を全角英文字に変換
        run.text = convert_ascii_to_fullwidth(run.text)
        # 「！」、「？」の後に全角スペースを挿入
        run.text = insert_space_after_punctuation(run.text)
        # イタリックの場合は解除してボールドに変更
        if run.italic:
            run.italic = False
            run.bold = True

# 処理後のドキュメントを保存（適宜パスを変更してください）
doc.save("/Users/a104/Desktop/output.docx")
