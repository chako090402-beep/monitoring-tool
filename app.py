import streamlit as st
import pdfplumber
import pandas as pd
import re
from calendar import monthrange
from datetime import datetime
import io

# --- 判定ロジック関数（安定版の内容をそのまま移植） ---
def get_stats(text):
    if not text: return {"lines": 0, "chars": 0}
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    pure = re.sub(r'[^\u3400-\u9FFF\u3040-\u309F\u30A0-\u30FFa-zA-Z0-9]', '', text)
    return {"lines": len(lines), "chars": len(pure)}

def extract_date_and_check(text):
    date_match = re.search(r'令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日', text)
    if date_match:
        try:
            y, m, d = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
            last_day = monthrange(y + 2018, m)[1]
            if d != last_day: return False, f"{m}/{d}"
        except: pass
    return True, ""

def extract_name(text):
    try:
        match = re.search(r'氏名[:：\s]*(.*?)様', text)
        if match: return re.sub(r'[\s　]', '', match.group(1))
    except: pass
    return ""

def process_pdf(pdf_file):
    data = {}
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            name = extract_name(text) or f"未特定_P{i+1}"
            is_end_of_month, d_msg = extract_date_and_check(text)
            
            # 実施状況・評価抽出
            status_content, eval_content = "", ""
            if "達成状況" in text:
                try:
                    after_status = text.split("達成状況")[1]
                    for sw in ["署名", "評価／今後の対応", "（サービス提供事業者）", "今後の方針"]:
                        if sw in after_status:
                            after_status = after_status.split(sw)[0]
                            break
                    status_content = after_status
                except: pass
            if "評価／今後の対応" in text or "今後の方針" in text:
                try:
                    kw = "評価／今後の対応" if "評価／今後の対応" in text else "今後の方針"
                    after_eval = text.split(kw)[1]
                    eval_content = after_eval.split("サービスの実施状況")[0] if "サービスの実施状況" in after_eval else after_eval.split("署名")[0] if "署名" in after_eval else after_eval
                except: pass

            data[name] = {"タイプ": "評価表" if "評価表" in text else "モニタリング", "status_stats": get_stats(status_content), "eval_stats": get_stats(eval_content), "date_ok": is_end_of_month, "date_msg": d_msg}
    return data

# --- Streamlit 画面表示部分 ---
st.set_page_config(page_title="評価表・モニタリング自動チェッカー", layout="wide")
st.title("📄評価表・ モニタリング自動チェッカー")
st.write("先月と今月のPDFをアップロードして、変更漏れを自動チェックします。")

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("① 先月のPDFを選択", type="pdf")
with col2:
    new_file = st.file_uploader("② 今月のPDFを選択", type="pdf")

if old_file and new_file:
    if st.button("チェック開始！"):
        with st.spinner('解析中...'):
            old_docs = process_pdf(old_file)
            new_docs = process_pdf(new_file)
            
            final_results = []
            for name, new_d in new_docs.items():
                old_key = next((k for k in old_docs.keys() if name in k or k in name), None)
                errors = []
                old_type = "（新規）"
                if not new_d["date_ok"]: errors.append(f"作成日が{new_d['date_msg']}(月末以外)")
                if old_key:
                    old_d = old_docs[old_key]
                    old_type = old_d["タイプ"]
                    if old_d["status_stats"] == new_d["status_stats"] and new_d["status_stats"]["chars"] > 0:
                        errors.append("実施状況に変化なし")
                    if old_d["タイプ"] == "評価表" and new_d["タイプ"] == "モニタリング":
                        if old_d["eval_stats"] == new_d["eval_stats"] and new_d["eval_stats"]["chars"] > 0:
                            errors.append("評価欄に変化なし")
                else: errors.append("前月不在")

                status = "OK" if not errors else "⚠️ 要確認"
                if "前月不在" in errors: status = "新規"
                final_results.append({"氏名": name, "判定": status, "前月書類": old_type, "今月書類": new_d["タイプ"], "詳細": " / ".join(errors) if errors else "更新済み"})

            for old_name in old_docs.keys():
                if not any(old_name in n or n in old_name for n in new_docs.keys()):
                    final_results.append({"氏名": old_name, "判定": "❌ 不在", "前月書類": old_docs[old_name]["タイプ"], "今月書類": "-", "詳細": "今月消失"})

            df = pd.DataFrame(final_results)
            st.success("解析が完了しました！")
            st.dataframe(df.style.applymap(lambda x: 'background-color: #ffcccc' if x in ['⚠️ 要確認', '❌ 不在'] else '', subset=['判定']), use_container_width=True)
            
            # エクセルダウンロード
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button(label="📥 結果をExcelでダウンロード", data=output.getvalue(), file_name="チェック結果.xlsx")
