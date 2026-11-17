import streamlit as st
import openpyxl
import io
import os
import sys
import json
import hashlib
import urllib.request
import re
import time
import subprocess
import shutil
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="HOSPIT Orbit / Navigation", layout="wide")

USER_DB_FILE = "users.json"
CONFIG_FILE = "master_config.json"
TEMPLATE_FILE = "template.xlsx"
HISTORY_DIR = "history"
HISTORY_LOG_FILE = os.path.join(HISTORY_DIR, "history.json")

os.makedirs(HISTORY_DIR, exist_ok=True)

# ------------------------------------------
# Linux / Windows 共通: LibreOffice + openpyxl によるPDF出力関数
# ------------------------------------------
def find_libreoffice_exec() -> str:
    """OSに応じてLibreOffice(soffice)の実行ファイルパスを探す"""
    # 1. PATH環境変数内を検索 (Linux環境では通常ここで見つかる)
    for cmd in ["libreoffice", "soffice"]:
        path = shutil.which(cmd)
        if path:
            return path

    # 2. Windowsの標準インストール先を検索
    if sys.platform == "win32":
        win_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in win_paths:
            if os.path.exists(p):
                return p
    return None

def configure_openpyxl_page_setup(ws, print_area=None, orientation="portrait"):
    """
    openpyxl だけでシートの印刷範囲と「横1ページ収め」を設定する（Linux完全対応）
    orientation: "portrait" (縦) または "landscape" (横)
    """
    if print_area:
        ws.print_area = print_area

    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # 「シートを1ページ幅に収める（縦は自動）」を有効化
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

def export_multiple_sheets_to_pdf_linux(excel_source_path: str, target_sheet_names: list, output_pdf_path: str) -> bool:
    """
    win32com(Excel)を使わず、openpyxlで印刷設定＆不要シート非表示を行い、
    LibreOfficeのヘッドレスモードで数式再計算とPDF生成を行う
    """
    soffice_bin = find_libreoffice_exec()
    if not soffice_bin:
        print("エラー: LibreOffice が見つかりません。")
        return False

    # PDF変換専用の一時Excelファイルを作成（不要シートを隠すため）
    temp_xlsx = excel_source_path.replace(".xlsx", "_for_pdf.xlsx")
    try:
        wb = openpyxl.load_workbook(excel_source_path)

        for ws in wb.worksheets:
            name = ws.title
            if name not in target_sheet_names:
                # PDF化対象外のシートは非表示にする（LibreOfficeは非表示シートをPDFから除外する）
                ws.sheet_state = "hidden"
            else:
                ws.sheet_state = "visible"
                # シート別の印刷範囲＆向き設定
                if name == "要件確定シート":
                    configure_openpyxl_page_setup(ws, print_area=None, orientation="portrait")
                elif name == "試算案":
                    configure_openpyxl_page_setup(
                        ws,
                        print_area=["$A$1:$O$69", "$AC$1:$AQ$101", "$A$80:$O$105"],
                        orientation="portrait"
                    )
                elif name == "見積書シート案":
                    configure_openpyxl_page_setup(
                        ws,
                        print_area=["$A$1:$P$85", "$A$86:$P$164"],
                        orientation="portrait"
                    )
                elif name == "按分表シート":
                    configure_openpyxl_page_setup(
                        ws,
                        print_area=["$A$1:$AG$38", "$AJ$1:$BP$38"],
                        orientation="landscape"
                    )

        wb.save(temp_xlsx)
        wb.close()

        out_dir = os.path.dirname(os.path.abspath(output_pdf_path))
        abs_temp_xlsx = os.path.abspath(temp_xlsx)

        # LibreOffice をヘッドレス（画面なし）モードで起動してPDF変換＆数式再計算
        cmd = [
            soffice_bin,
            "--headless",
            "--convert-to", "pdf:calc_pdf_Export",
            "--outdir", out_dir,
            abs_temp_xlsx
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)

        # LibreOfficeが出力したPDFファイル名（_for_pdf.pdf）を目的のファイル名にリネーム
        generated_pdf = os.path.join(out_dir, os.path.basename(temp_xlsx).replace(".xlsx", ".pdf"))
        if os.path.exists(generated_pdf):
            if os.path.exists(output_pdf_path):
                os.remove(output_pdf_path)
            os.rename(generated_pdf, output_pdf_path)
            return True
        else:
            print(f"LibreOffice変換エラー: {result.stderr.decode('utf-8', errors='ignore')}")
            return False

    except Exception as e:
        print(f"Linux方式PDF出力エラー: {e}")
        return False
    finally:
        # 一時用Excelファイルを削除
        if os.path.exists(temp_xlsx):
            try:
                os.remove(temp_xlsx)
            except Exception:
                pass

# ------------------------------------------
# 0. マスタ設定 & 履歴管理関数
# ------------------------------------------
DEFAULT_CONFIG = {
    "install_team_size": 3,
    "install_units_per_day": 50,
    "worker_daily_wage": 35000,
    "shipping_unit_cost": 2500,
    "shipping_base_charge": 50000,
    "kitting_unit_cost": 8000,
    "hardware_unit_cost": 45000,
    "license_unit_cost": 15000,
    "server_onprem_cost": 500000,
    "server_cloud_cost": 200000,
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        return DEFAULT_CONFIG.copy()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            if k not in cfg:
                cfg[k] = v
        return cfg

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

config = load_config()

def load_history():
    if not os.path.exists(HISTORY_LOG_FILE):
        return []
    try:
        with open(HISTORY_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def add_history_record(record):
    records = load_history()
    records.insert(0, record)
    with open(HISTORY_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

# ------------------------------------------
# 1. ユーザー管理
# ------------------------------------------
def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def normalize_role(raw_role: str) -> str:
    if raw_role in ["admin", "internal_admin", "社内（管理者）"]:
        return "internal_admin"
    if raw_role in ["agency", "販売代理店"]:
        return "agency"
    return "internal_user"

def load_users() -> dict:
    default_users = {
        "admin@ezix.co.jp": {
            "user_id": "admin@ezix.co.jp",
            "user_name": "社内 管理者",
            "department": "管理",
            "role": "internal_admin",
            "password_hash": hash_password("admin1234")
        },
        "staff@ezix.co.jp": {
            "user_id": "staff@ezix.co.jp",
            "user_name": "社内 営業担当",
            "department": "営業",
            "role": "internal_user",
            "password_hash": hash_password("staff1234")
        },
        "agency@partner.co.jp": {
            "user_id": "agency@partner.co.jp",
            "user_name": "代理店 担当者",
            "department": "販売店",
            "role": "agency",
            "password_hash": hash_password("agency1234")
        }
    }
    
    if not os.path.exists(USER_DB_FILE) or os.path.getsize(USER_DB_FILE) == 0:
        with open(USER_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(default_users, f, ensure_ascii=False, indent=2)
        return default_users
        
    try:
        with open(USER_DB_FILE, "r", encoding="utf-8") as f:
            users = json.load(f)
    except json.JSONDecodeError:
        with open(USER_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(default_users, f, ensure_ascii=False, indent=2)
        return default_users

    updated = False
    if "admin@ezix.co.jp" in users:
        users["admin@ezix.co.jp"]["role"] = "internal_admin"
        updated = True
    for k, v in default_users.items():
        if k not in users:
            users[k] = v
            updated = True
    if updated:
        with open(USER_DB_FILE, "w", encoding="utf-8") as fw:
            json.dump(users, fw, ensure_ascii=False, indent=2)
    return users

def save_users(users: dict):
    with open(USER_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

users = load_users()

if "logged_in_user" not in st.session_state:
    st.session_state["logged_in_user"] = None
if "current_page" not in st.session_state:
    st.session_state["current_page"] = "menu"
if "corp_auto_address" not in st.session_state:
    st.session_state["corp_auto_address"] = ""
if "hosp_auto_address" not in st.session_state:
    st.session_state["hosp_auto_address"] = ""

# 郵便番号検索
def fetch_address_corp():
    raw_zip = st.session_state.get("corp_zip_input", "")
    zip_clean = re.sub(r"\D", "", raw_zip)
    if len(zip_clean) == 7:
        url = f"https://zipcloud.ibsnet.co.jp/api/search?zipcode={zip_clean}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as res:
                data = json.loads(res.read().decode("utf-8"))
                if data.get("results"):
                    item = data["results"][0]
                    st.session_state["corp_auto_address"] = f"{item['address1']}{item['address2']}{item['address3']}"
                else:
                    st.toast("該当する住所が見つかりませんでした。")
        except Exception:
            st.toast("住所の取得に失敗しました。")

def fetch_address_hosp():
    raw_zip = st.session_state.get("hosp_zip_input", "")
    zip_clean = re.sub(r"\D", "", raw_zip)
    if len(zip_clean) == 7:
        url = f"https://zipcloud.ibsnet.co.jp/api/search?zipcode={zip_clean}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as res:
                data = json.loads(res.read().decode("utf-8"))
                if data.get("results"):
                    item = data["results"][0]
                    st.session_state["hosp_auto_address"] = f"{item['address1']}{item['address2']}{item['address3']}"
                else:
                    st.toast("該当する住所が見つかりませんでした。")
        except Exception:
            st.toast("住所の取得に失敗しました。")

# ------------------------------------------
# 2. ログイン画面
# ------------------------------------------
if st.session_state["logged_in_user"] is None:
    st.title("HOSPIT Orbit / Navigation")
    st.markdown("##### Delivering Joy and Inspiration to Every User")
    st.caption("アカウント権限に応じて社内ポータルまたは販売代理店ポータルへ遷移します。")

    col_log, _ = st.columns([1, 1])
    with col_log:
        with st.form("login_form"):
            login_email = st.text_input("ID（メールアドレス）")
            login_password = st.text_input("パスワード", type="password")
            submit_login = st.form_submit_button("ログイン", type="primary")

        if submit_login:
            if login_email in users and users[login_email]["password_hash"] == hash_password(login_password):
                st.session_state["logged_in_user"] = users[login_email]
                st.session_state["current_page"] = "menu"
                st.rerun()
            else:
                st.error("IDまたはパスワードが正しくありません。")

    with st.expander("テスト用アカウント一覧"):
        st.write("1. 社内管理者: admin@ezix.co.jp / admin1234")
        st.write("2. 社内一般: staff@ezix.co.jp / staff1234")
        st.write("3. 販売代理店: agency@partner.co.jp / agency1234")
    st.stop()

# ------------------------------------------
# 3. 共通ヘッダー
# ------------------------------------------
user = st.session_state["logged_in_user"]
role = normalize_role(user.get("role", "internal_user"))

head_col1, head_col2 = st.columns([4, 1])
with head_col1:
    role_label = {
        "internal_admin": "【社内・管理者】",
        "internal_user": "【社内】",
        "agency": "【販売代理店】"
    }.get(role, "")
    st.write(f"{role_label} {user['user_name']} さん（{user.get('department', '')}）")
with head_col2:
    if st.button("ログアウト"):
        st.session_state["logged_in_user"] = None
        st.session_state["current_page"] = "menu"
        st.rerun()

st.divider()

# ------------------------------------------
# 4. 管理者専用サイドバー
# ------------------------------------------
if role == "internal_admin":
    with st.sidebar:
        st.header("管理者専用メニュー")
        
        with st.expander("アカウント発行・権限設定", expanded=False):
            with st.form("admin_user_reg_form"):
                new_name = st.text_input("氏名")
                new_dept = st.text_input("担当部署 / 販売代理店名")
                new_role_select = st.selectbox("権限区分", ["社内（一般）", "社内（管理者）", "販売代理店"])
                new_email = st.text_input("ID（メールアドレス）")
                new_pw = st.text_input("初期パスワード", type="password")
                submit_reg = st.form_submit_button("アカウントを発行", type="primary")

            if submit_reg:
                role_map = {"社内（管理者）": "internal_admin", "社内（一般）": "internal_user", "販売代理店": "agency"}
                if not new_name or not new_email or not new_pw:
                    st.warning("必須項目を入力してください。")
                elif new_email in users:
                    st.error("このIDは既に登録されています。")
                else:
                    users[new_email] = {
                        "user_id": new_email,
                        "user_name": new_name,
                        "department": new_dept,
                        "role": role_map[new_role_select],
                        "password_hash": hash_password(new_pw),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    save_users(users)
                    st.success(f"{new_name} さんのアカウントを発行しました。")
                    st.rerun()

            st.divider()
            st.caption("【登録済みアカウント】")
            for uid, udata in users.items():
                r_tag = normalize_role(udata.get("role", ""))
                st.text(f"・{udata['user_name']} [{r_tag}]\n  ID: {uid}")

        with st.expander("原価・算定条件マスタ", expanded=False):
            with st.form("config_edit_form"):
                c_team = st.number_input("標準作業員数 (名/組)", min_value=1, value=int(config["install_team_size"]))
                c_units = st.number_input("1日あたり設置能力 (台/日)", min_value=1, value=int(config["install_units_per_day"]))
                c_wage = st.number_input("作業員日当単価 (円/人日)", min_value=0, step=1000, value=int(config["worker_daily_wage"]))
                c_hw = st.number_input("機器本体単価 (円/台)", min_value=0, step=1000, value=int(config.get("hardware_unit_cost", 45000)))
                c_kit = st.number_input("キッティング作業単価 (円/台)", min_value=0, step=500, value=int(config["kitting_unit_cost"]))
                c_ship_unit = st.number_input("1台あたり配送運賃 (円/台)", min_value=0, step=100, value=int(config["shipping_unit_cost"]))
                c_ship_base = st.number_input("基本チャーター運賃 (円/式)", min_value=0, step=5000, value=int(config["shipping_base_charge"]))
                c_lic = st.number_input("ライセンス単価 (円/台)", min_value=0, step=1000, value=int(config["license_unit_cost"]))
                c_srv_onprem = st.number_input("オンプレサーバー費 (円)", min_value=0, step=10000, value=int(config["server_onprem_cost"]))
                c_srv_cloud = st.number_input("クラウド環境構築費 (円)", min_value=0, step=10000, value=int(config["server_cloud_cost"]))
                if st.form_submit_button("マスタ条件を保存"):
                    config.update({
                        "install_team_size": c_team, "install_units_per_day": c_units, "worker_daily_wage": c_wage,
                        "hardware_unit_cost": c_hw, "kitting_unit_cost": c_kit, "shipping_unit_cost": c_ship_unit,
                        "shipping_base_charge": c_ship_base, "license_unit_cost": c_lic,
                        "server_onprem_cost": c_srv_onprem, "server_cloud_cost": c_srv_cloud
                    })
                    save_config(config)
                    st.success("マスタ条件を更新しました。")
                    st.rerun()

# ------------------------------------------
# 5. 画面ルーティング
# ------------------------------------------
if role in ["internal_admin", "internal_user"]:
    if st.session_state["current_page"] == "menu":
        st.subheader("エジックス社内ポータル")
        st.write("実行したい業務を選択してください。")
        
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            with st.container(border=True):
                st.markdown("### 案件管理")
                st.write("申請された案件の確認、ステージ進捗管理、決定途プロセスの追跡")
                st.caption("未処理の新規申請: 2 件")
                if st.button("案件管理を開く", use_container_width=True):
                    st.session_state["current_page"] = "project_mgmt"
                    st.rerun()

        with m_col2:
            with st.container(border=True):
                st.markdown("### 試算書管理")
                st.write("病院向け・代理店向け試算書の作成、Excel/PDF帳票出力、履歴管理")
                st.caption("最新テンプレート連動（Linux/LibreOffice対応版）")
                if st.button("試算書管理を開く", type="primary", use_container_width=True):
                    st.session_state["current_page"] = "estimate_mgmt"
                    st.rerun()

    elif st.session_state["current_page"] == "project_mgmt":
        if st.button("← 社内メニューへ戻る"):
            st.session_state["current_page"] = "menu"
            st.rerun()
        st.title("案件管理")
        st.info("申請案件一覧、ステージ管理（決定途プロセス、課題調査など）を実装する画面です。")

    elif st.session_state["current_page"] == "estimate_mgmt":
        if st.button("← 社内メニューへ戻る"):
            st.session_state["current_page"] = "menu"
            st.rerun()
        
        st.title("試算書管理")
        tab_create, tab_history = st.tabs(["試算・見積作成", "過去の試算・提出履歴"])
        
        # ------------------------------------------
        # タブ1: 試算・見積作成
        # ------------------------------------------
        with tab_create:
            col_proj, col_agency = st.columns(2)
            with col_proj:
                st.subheader("1. 案件基本情報")
                today_str = datetime.now().strftime("%Y%m%d")
                project_code = st.text_input("案件番号", value=f"EST-{today_str}-01")
                created_date = st.date_input("作成日", value=datetime.today())

            with col_agency:
                st.subheader("2. 販売代理店情報")
                agency_name = st.selectbox("販売代理店名", ["会社A", "会社B", "会社C"])

            st.divider()
            st.subheader("3. 法人・施設情報")
            
            col_corp, col_fac = st.columns(2)
            with col_corp:
                st.markdown("##### 【法人情報】")
                corp_name = st.text_input("法人名", placeholder="例：医療法人 〇〇会")
                corp_zip = st.text_input("法人 郵便番号", placeholder="例：1000001", key="corp_zip_input", on_change=fetch_address_corp)
                corp_address = st.text_input("法人 住所", value=st.session_state["corp_auto_address"])

            with col_fac:
                st.markdown("##### 【施設情報】")
                facility_name = st.text_input("施設名", placeholder="例：〇〇病院")
                facility_zip = st.text_input("施設 郵便番号", placeholder="例：1000001", key="hosp_zip_input", on_change=fetch_address_hosp)
                facility_address = st.text_input("施設 住所", value=st.session_state["hosp_auto_address"])

            st.divider()
            st.subheader("4. 試算条件")
            col_cond1, col_cond2 = st.columns(2)
            with col_cond1:
                scheme_type = st.radio("販売スキーム", ["新スキーム（代理店機器調達・キッティング役務請求）", "従来スキーム（機器本体込み販売）"])
                bed_count = st.number_input("病床数", min_value=1, value=100, step=1)
            with col_cond2:
                server_type = st.selectbox("サーバータイプ", ["オンプレ", "クラウド"])
                contract_years = st.selectbox("契約期間", ["6年", "7年"])

            # ------------------------------------------
            # 試算実行 & ファイル出力
            # ------------------------------------------
            st.divider()
            submitted = st.button("試算を実行してExcel・PDFを作成", type="primary")

            if submitted:
                if not corp_name or not facility_name:
                    st.warning("法人名と施設名は必須です。")
                elif not os.path.exists(TEMPLATE_FILE):
                    st.error(f"{TEMPLATE_FILE} が見つかりません。")
                else:
                    try:
                        with st.status("試算処理を実行中...", expanded=True) as status:
                            status.write("1/3 テンプレート読み込み & データ入力中...")
                            with open(TEMPLATE_FILE, "rb") as f_tpl:
                                template_bytes = io.BytesIO(f_tpl.read())
                            
                            wb = openpyxl.load_workbook(template_bytes)

                            if "要件確定シート" in wb.sheetnames:
                                ws_input = wb["要件確定シート"]
                            else:
                                ws_input = wb.active

                            # セル代入（確定番地）
                            ws_input["F4"] = corp_name
                            ws_input["W4"] = corp_zip
                            ws_input["V5"] = corp_address
                            ws_input["F6"] = facility_name
                            ws_input["W6"] = facility_zip
                            ws_input["V7"] = facility_address
                            ws_input["V8"] = bed_count
                            ws_input["H10"] = server_type
                            ws_input["AL4"] = contract_years

                            safe_name = facility_name if facility_name else "試算書"
                            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                            file_base_name = f"{project_code}_{safe_name}_{timestamp_str}"
                            output_xlsx_path = os.path.join(HISTORY_DIR, f"{file_base_name}.xlsx")
                            output_pdf_path = os.path.join(HISTORY_DIR, f"{file_base_name}.pdf")

                            status.write("2/3 Excelファイルを生成中...")
                            wb.save(output_xlsx_path)
                            wb.close()

                            status.write("3/3 LibreOffice(Linux互換モード)でPDF変換を実行中...")
                            target_sheets = ["要件確定シート", "試算案", "見積書シート案", "按分表シート"]
                            pdf_ok = export_multiple_sheets_to_pdf_linux(output_xlsx_path, target_sheets, output_pdf_path)

                            new_record = {
                                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "project_code": project_code,
                                "corp_name": corp_name,
                                "facility_name": facility_name,
                                "bed_count": bed_count,
                                "scheme_type": scheme_type,
                                "server_type": server_type,
                                "contract_years": contract_years,
                                "author": user["user_name"],
                                "xlsx_file": f"{file_base_name}.xlsx",
                                "pdf_file": f"{file_base_name}.pdf" if pdf_ok else None
                            }
                            add_history_record(new_record)

                            status.update(label="処理が完了しました！", state="complete", expanded=False)

                        with open(output_xlsx_path, "rb") as f_xl:
                            excel_bytes = f_xl.read()

                        pdf_bytes = None
                        if pdf_ok and os.path.exists(output_pdf_path):
                            with open(output_pdf_path, "rb") as f_pdf:
                                pdf_bytes = f_pdf.read()

                        dl_col1, dl_col2 = st.columns(2)
                        with dl_col1:
                            st.download_button(
                                label="Excelをダウンロード",
                                data=excel_bytes,
                                file_name=f"{project_code}_{safe_name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        with dl_col2:
                            if pdf_bytes:
                                st.download_button(
                                    label="PDFをダウンロード（Linux互換・4シート）",
                                    data=pdf_bytes,
                                    file_name=f"{project_code}_{safe_name}.pdf",
                                    mime="application/pdf"
                                )
                            else:
                                st.warning("PDF生成に失敗しました。LibreOfficeがインストールされているか確認してください。")

                    except PermissionError as pe:
                        st.error(f"権限エラー: {pe}")
                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")

        # ------------------------------------------
        # タブ2: 過去の試算・提出履歴
        # ------------------------------------------
        with tab_history:
            st.subheader("過去の試算・提出履歴一覧")
            records = load_history()

            if not records:
                st.info("まだ試算履歴がありません。「試算・見積作成」タブで試算を実行するとここに一覧表示されます。")
            else:
                df_data = []
                for idx, r in enumerate(records):
                    df_data.append({
                        "作成日時": r.get("timestamp", ""),
                        "案件番号": r.get("project_code", ""),
                        "法人名": r.get("corp_name", ""),
                        "施設名": r.get("facility_name", ""),
                        "病床数": r.get("bed_count", ""),
                        "スキーム": r.get("scheme_type", ""),
                        "期間": r.get("contract_years", ""),
                        "作成者": r.get("author", "")
                    })
                
                df = pd.DataFrame(df_data)
                st.dataframe(df, use_container_width=True, hide_index=True)

                st.divider()
                st.markdown("#### 過去ファイルの個別確認・再ダウンロード")
                
                options = [f"[{r.get('timestamp')}] {r.get('project_code')} - {r.get('facility_name')}" for r in records]
                selected_opt = st.selectbox("ダウンロードしたい履歴を選択してください", options)
                selected_idx = options.index(selected_opt)
                target_rec = records[selected_idx]

                h_col1, h_col2 = st.columns(2)
                xlsx_path = os.path.join(HISTORY_DIR, target_rec.get("xlsx_file", ""))
                pdf_path = os.path.join(HISTORY_DIR, target_rec.get("pdf_file", "")) if target_rec.get("pdf_file") else None

                with h_col1:
                    if os.path.exists(xlsx_path):
                        with open(xlsx_path, "rb") as f:
                            st.download_button(
                                label=f"Excel 再ダウンロード ({target_rec.get('project_code')})",
                                data=f.read(),
                                file_name=f"{target_rec.get('project_code')}_{target_rec.get('facility_name')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"hist_xl_{selected_idx}"
                            )
                    else:
                        st.caption("Excelファイルが見つかりません。")

                with h_col2:
                    if pdf_path and os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            st.download_button(
                                label=f"PDF 再ダウンロード ({target_rec.get('project_code')})",
                                data=f.read(),
                                file_name=f"{target_rec.get('project_code')}_{target_rec.get('facility_name')}.pdf",
                                mime="application/pdf",
                                key=f"hist_pdf_{selected_idx}"
                            )
                    else:
                        st.caption("PDFファイルはありません。")

# B. 販売代理店ポータル
elif role == "agency":
    if st.session_state["current_page"] == "menu":
        st.subheader("Partner Room（販売代理店ポータル）")
        st.write("新規の案件申請、またはこれまでに申請した案件の確認を行えます。")
        
        a_col1, a_col2 = st.columns(2)
        with a_col1:
            with st.container(border=True):
                st.markdown("### 案件申請")
                st.write("病院情報、イベント種別、決定キーマン、予定時期などを登録してエジックスへ申請")
                if st.button("新規案件を申請する", type="primary", use_container_width=True):
                    st.session_state["current_page"] = "agency_apply"
                    st.rerun()

        with a_col2:
            with st.container(border=True):
                st.markdown("### 案件一覧")
                st.write("過去に申請した案件の状況確認、試算結果の受領")
                if st.button("申請済み案件一覧を見る", use_container_width=True):
                    st.session_state["current_page"] = "agency_list"
                    st.rerun()

    elif st.session_state["current_page"] == "agency_apply":
        if st.button("← ポータルへ戻る"):
            st.session_state["current_page"] = "menu"
            st.rerun()
        st.title("案件申請")
        st.info("申請項目（病院情報、キーマン、予定時期など）の入力画面です。")

    elif st.session_state["current_page"] == "agency_list":
        if st.button("← ポータルへ戻る"):
            st.session_state["current_page"] = "menu"
            st.rerun()
        st.title("案件一覧")
        st.info("申請済み案件の状況確認画面です。")