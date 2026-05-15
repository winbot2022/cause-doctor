import io
import json
import re
from datetime import datetime
from typing import Dict, List, Any

import streamlit as st
from PIL import Image as PILImage

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


st.set_page_config(
    page_title="作業手順書作成アプリ",
    layout="wide",
)


# =========================
# Excel書式
# =========================
title_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
section_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")
header_fill = PatternFill(fill_type="solid", fgColor="FCE4D6")
thin_gray_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")

center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
top_left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

thin_border = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

# =========================
# 工程別 標準手順テンプレート
# =========================

PROCESS_TEMPLATES = {
    "汎用テンプレート": [
        "作業前に安全を確認する",
        "作業対象を確認する",
        "必要な工具・治具を準備する",
        "作業場所を整理し、異物がないことを確認する",
        "作業対象を所定の位置にセットする",
        "作業を実施する",
        "作業結果を確認する",
        "良品を所定の置き場に置く",
        "異常があれば作業を止め、責任者へ報告する",
    ],

    "マシニング加工": [
        "作業前に設備周辺の安全を確認する",
        "図面・加工指示書・品番を確認する",
        "使用する治具・工具・測定具を準備する",
        "材料またはワークの品番・数量を確認する",
        "治具・バイス・クランプ部を清掃する",
        "ワークを所定の位置にセットする",
        "クランプ状態を確認する",
        "加工プログラム番号を確認する",
        "工具番号・工具状態を確認する",
        "原点・加工開始位置を確認する",
        "扉を閉め、安全を確認する",
        "起動ボタンを押し加工を開始する",
        "加工中の異音・振動・切削液の状態を確認する",
        "加工完了後、ワークを取り出す",
        "加工面・キズ・バリの有無を確認する",
        "必要寸法を測定する",
        "良品を所定の置き場に置く",
        "異常があれば作業を止め、責任者へ報告する",
    ],

    "NC旋盤加工": [
        "作業前に設備周辺の安全を確認する",
        "図面・加工指示書・品番を確認する",
        "使用するチャック・爪・工具・測定具を準備する",
        "材料またはワークの品番・数量を確認する",
        "チャック・爪・ワーク接触面を清掃する",
        "ワークをチャックにセットする",
        "チャックの締付状態を確認する",
        "突き出し長さを確認する",
        "加工プログラム番号を確認する",
        "工具番号・工具摩耗状態を確認する",
        "扉を閉め、安全を確認する",
        "起動ボタンを押し加工を開始する",
        "加工中の異音・振動・切削液の状態を確認する",
        "加工完了後、ワークを取り出す",
        "加工面・キズ・バリの有無を確認する",
        "外径・内径・長さなど必要寸法を測定する",
        "良品を所定の置き場に置く",
        "異常があれば作業を止め、責任者へ報告する",
    ],

    "プレス加工": [
        "作業前に設備周辺の安全を確認する",
        "加工指示書・品番・材料を確認する",
        "金型・治具の状態を確認する",
        "材料の品番・数量・向きを確認する",
        "作業台・金型周辺に異物がないことを確認する",
        "材料を所定の位置にセットする",
        "手や指が危険範囲にないことを確認する",
        "両手操作・安全装置の状態を確認する",
        "起動ボタンを押して加工する",
        "加工後の製品を取り出す",
        "変形・キズ・打痕・バリの有無を確認する",
        "必要に応じて寸法または形状を確認する",
        "良品を所定の置き場に置く",
        "不良品は識別して所定の場所に置く",
        "異常があれば作業を止め、責任者へ報告する",
    ],

    "バリ取り": [
        "作業前に対象品の品番・数量を確認する",
        "図面・作業指示書・バリ取り箇所を確認する",
        "使用する工具・治具・保護具を準備する",
        "作業台を清掃し、異物がないことを確認する",
        "対象品を安定した位置に置く",
        "指定箇所のバリを除去する",
        "エッジ部・穴周辺・加工端面を確認する",
        "必要に応じて面取り状態を確認する",
        "キズ・打痕を付けていないか確認する",
        "バリ残りがないか目視確認する",
        "必要に応じて手触りまたは測定で確認する",
        "良品を所定の置き場に置く",
        "不良または判断に迷う品は識別する",
        "異常があれば責任者へ報告する",
    ],

    "外観検査": [
        "検査対象品の品番・数量を確認する",
        "図面・検査基準・限度見本を確認する",
        "検査場所の照明・作業環境を確認する",
        "対象品を検査しやすい位置に置く",
        "外観にキズ・打痕・汚れがないか確認する",
        "バリ・欠け・変形がないか確認する",
        "端面・穴・曲げ部・加工面を確認する",
        "必要に応じて拡大鏡・照明を使用する",
        "判定結果を記録する",
        "良品と不良品を区分する",
        "不良品は識別し、所定の場所に置く",
        "同じ不良が続く場合は責任者へ報告する",
    ],

    "寸法測定": [
        "測定対象品の品番・数量を確認する",
        "図面・測定箇所・公差を確認する",
        "使用する測定器を準備する",
        "測定器のゼロ点・校正状態を確認する",
        "測定面の汚れ・切粉・油を除去する",
        "対象品を安定した状態で保持する",
        "指定された箇所を測定する",
        "測定値を記録する",
        "公差内であることを確認する",
        "測定結果に異常があれば再測定する",
        "規格外の場合は識別し、責任者へ報告する",
        "測定器を清掃し、所定の場所に戻す",
    ],

    "組立": [
        "作業前に作業指示書・図面・品番を確認する",
        "使用する部品・数量を確認する",
        "使用する工具・治具を準備する",
        "作業台を清掃し、異物がないことを確認する",
        "部品にキズ・変形・異物がないか確認する",
        "部品の向き・表裏・左右を確認する",
        "部品を所定の位置に組み付ける",
        "ボルト・ナット・部品を仮締めする",
        "指定トルクまたは指定方法で締め付ける",
        "組付け状態・ガタ・ズレを確認する",
        "必要に応じて動作確認を行う",
        "外観・キズ・汚れを確認する",
        "完成品を所定の置き場に置く",
        "異常があれば責任者へ報告する",
    ],

    "洗浄": [
        "洗浄対象品の品番・数量を確認する",
        "洗浄方法・洗浄条件を確認する",
        "使用する洗浄液・設備・治具を確認する",
        "対象品に大きな異物・切粉がないか確認する",
        "対象品を洗浄治具または所定位置にセットする",
        "洗浄条件を確認して洗浄を開始する",
        "洗浄中に異常音・漏れ・停止がないか確認する",
        "洗浄後、対象品を取り出す",
        "水分・洗浄液・汚れ残りがないか確認する",
        "必要に応じてエアブローまたは乾燥を行う",
        "キズ・変色・異物残りがないか確認する",
        "良品を所定の置き場に置く",
        "異常があれば責任者へ報告する",
    ],

    "梱包": [
        "梱包対象品の品番・数量を確認する",
        "梱包仕様・出荷指示を確認する",
        "使用する箱・緩衝材・ラベルを準備する",
        "製品にキズ・汚れ・異物がないか確認する",
        "製品を指定数量ごとに並べる",
        "緩衝材または仕切りを所定の位置に入れる",
        "製品を指定の向き・位置で箱に入れる",
        "数量を再確認する",
        "ラベル・品番・納入先を確認する",
        "箱を閉じ、指定方法で封をする",
        "梱包状態に破損・ズレがないか確認する",
        "所定の出荷場所に置く",
        "異常があれば責任者へ報告する",
    ],

    "日常点検": [
        "点検対象設備を確認する",
        "点検表または点検項目を確認する",
        "設備周辺の安全状態を確認する",
        "電源・非常停止・安全カバーの状態を確認する",
        "油量・エア圧・冷却水・切削液を確認する",
        "異音・異臭・振動がないか確認する",
        "配管・ホース・ケーブルに異常がないか確認する",
        "治具・工具・可動部に異常がないか確認する",
        "清掃が必要な箇所を清掃する",
        "点検結果を記録する",
        "異常がある場合は使用を止め、責任者へ報告する",
        "点検後、設備を通常状態に戻す",
    ],
}


# =========================
# ポイント候補
# =========================

POINT_OPTIONS = [
    "奥まで確実に入れる",
    "指定位置に合わせる",
    "向きを合わせる",
    "表裏を確認する",
    "左右を確認する",
    "品番を照合する",
    "数量を確認する",
    "図面と照合する",
    "ラベルと現品を照合する",
    "固定状態を確認する",
    "ガタつきがないことを確認する",
    "締付状態を確認する",
    "指定トルクで締め付ける",
    "異物を除去する",
    "切粉を除去する",
    "油・汚れを除去する",
    "キズを付けないように扱う",
    "落下しないように保持する",
    "加工面に触れないようにする",
    "基準面を確認する",
    "測定面を清掃する",
    "ゼロ点を確認する",
    "公差を確認する",
    "良否判定基準を確認する",
    "限度見本と比較する",
    "異音・振動を確認する",
    "切削液の状態を確認する",
    "安全装置の状態を確認する",
    "作業後に所定位置へ戻す",
]


# =========================
# 注意事項候補
# =========================

CAUTION_OPTIONS = [
    "指はさみに注意",
    "巻き込まれに注意",
    "切粉に注意",
    "高温部に注意",
    "刃物に注意",
    "落下に注意",
    "重量物の取り扱いに注意",
    "キズに注意",
    "打痕に注意",
    "変形に注意",
    "異品混入に注意",
    "数量違いに注意",
    "品番違いに注意",
    "向き違いに注意",
    "表裏違いに注意",
    "左右違いに注意",
    "締付不足に注意",
    "締め忘れに注意",
    "入れ忘れに注意",
    "取り忘れに注意",
    "測定忘れに注意",
    "記録忘れに注意",
    "油汚れに注意",
    "異物混入に注意",
    "保護具を着用する",
    "保護メガネを着用する",
    "手袋を着用する",
    "設備停止を確認する",
    "非常停止位置を確認する",
    "周囲の安全を確認する",
    "不明点は責任者に確認する",
]


# =========================
# 確認項目候補
# =========================

CHECK_OPTIONS = [
    "品番OK",
    "数量OK",
    "図面番号OK",
    "材料OK",
    "向きOK",
    "表裏OK",
    "左右OK",
    "位置OK",
    "固定OK",
    "締付OK",
    "トルクOK",
    "プログラム番号OK",
    "工具番号OK",
    "治具OK",
    "原点OK",
    "ゼロ点OK",
    "寸法OK",
    "公差内",
    "外観OK",
    "キズなし",
    "打痕なし",
    "バリなし",
    "欠けなし",
    "変形なし",
    "汚れなし",
    "異物なし",
    "異音なし",
    "振動なし",
    "漏れなし",
    "安全装置OK",
    "保護具OK",
    "記録OK",
    "良品置き場OK",
    "不良品識別OK",
]
# =========================
# セッション初期化
# =========================
def init_session_state() -> None:
    defaults = {
        "manual_title": "",
        "process_type": "汎用テンプレート",
        "equipment_name": "",
        "product_name": "",
        "part_number": "",
        "drawing_number": "",
        "program_number": "",
        "jig_name": "",
        "tool_name": "",
        "author": "",
        "revision": "Rev.0",
        "selected_steps": [],
        "extra_step_input": "",
        "step_details": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session_state()


def safe_value(value: str) -> str:
    value = str(value).strip()
    return value if value else "未設定"

def safe_filename(text: str) -> str:
    text = str(text or "").strip()
    if not text:
        text = "未選択"
    # Windowsファイル名に使えない文字を除去
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    return text

def move_step_up(index: int) -> None:
    """選択済み手順を1つ上へ移動する"""
    if index <= 0:
        return

    steps = st.session_state.selected_steps
    steps[index - 1], steps[index] = steps[index], steps[index - 1]
    st.session_state.selected_steps = steps


def move_step_down(index: int) -> None:
    """選択済み手順を1つ下へ移動する"""
    steps = st.session_state.selected_steps

    if index >= len(steps) - 1:
        return

    steps[index], steps[index + 1] = steps[index + 1], steps[index]
    st.session_state.selected_steps = steps

def build_template_json_data() -> Dict[str, Any]:
    """
    現在の入力内容をJSON保存用に整形する。
    写真データは保存しない。
    """
    step_details_for_json = {}

    for step in st.session_state.selected_steps:
        detail = st.session_state.step_details.get(step, {})

        step_details_for_json[step] = {
            "point": detail.get("point", []),
            "caution": detail.get("caution", []),
            "check": detail.get("check", []),
            "free_point": detail.get("free_point", ""),
            "free_caution": detail.get("free_caution", ""),
            "free_check": detail.get("free_check", ""),
        }

    data = {
        "manual_title": st.session_state.manual_title,
        "process_type": st.session_state.process_type,
        "equipment_name": st.session_state.equipment_name,
        "product_name": st.session_state.product_name,
        "part_number": st.session_state.part_number,
        "drawing_number": st.session_state.drawing_number,
        "program_number": st.session_state.program_number,
        "jig_name": st.session_state.jig_name,
        "tool_name": st.session_state.tool_name,
        "author": st.session_state.author,
        "revision": st.session_state.revision,
        "selected_steps": st.session_state.selected_steps,
        "step_details": step_details_for_json,
    }

    return data


def apply_template_json_data(data: Dict[str, Any]) -> None:
    """
    JSONから読み込んだ内容をsession_stateへ反映する。
    写真欄は空で初期化する。
    """
    st.session_state.manual_title = data.get("manual_title", "")
    st.session_state.process_type = data.get("process_type", "汎用テンプレート")
    st.session_state.equipment_name = data.get("equipment_name", "")
    st.session_state.product_name = data.get("product_name", "")
    st.session_state.part_number = data.get("part_number", "")
    st.session_state.drawing_number = data.get("drawing_number", "")
    st.session_state.program_number = data.get("program_number", "")
    st.session_state.jig_name = data.get("jig_name", "")
    st.session_state.tool_name = data.get("tool_name", "")
    st.session_state.author = data.get("author", "")
    st.session_state.revision = data.get("revision", "Rev.0")

    selected_steps = data.get("selected_steps", [])
    if not isinstance(selected_steps, list):
        selected_steps = []

    st.session_state.selected_steps = selected_steps

    loaded_step_details = data.get("step_details", {})
    if not isinstance(loaded_step_details, dict):
        loaded_step_details = {}

    new_step_details = {}

    for step in st.session_state.selected_steps:
        detail = loaded_step_details.get(step, {})

        new_step_details[step] = {
            "point": detail.get("point", []),
            "caution": detail.get("caution", []),
            "check": detail.get("check", []),
            "free_point": detail.get("free_point", ""),
            "free_caution": detail.get("free_caution", ""),
            "free_check": detail.get("free_check", ""),
            "image": None,
            "image_name": "",
        }

    st.session_state.step_details = new_step_details

# =========================
# 画面UI
# =========================

st.title("作業手順書作成アプリ")
st.caption("工程を選び、写真と注意点を追加するだけで、Excel作業手順書のたたき台を作成します。")

st.info("分かる範囲で入力してください。未入力でも作成できます。")

with st.expander("基本情報", expanded=True):
    col1, col2 = st.columns(2)

# =========================
# 定型フォーム保存・読込
# =========================

with st.expander("定型フォームの保存・読込", expanded=False):
    st.caption("写真を除いた入力内容をJSON形式で保存・再利用できます。同じ工程の手順書を作るときに便利です。")

    col_json1, col_json2 = st.columns(2)

    with col_json1:
        st.markdown("#### 現在の内容を保存")

        template_data = build_template_json_data()
        json_bytes = json.dumps(
            template_data,
            ensure_ascii=False,
            indent=2,
        ).encode("utf-8")

        process_name = safe_filename(st.session_state.process_type)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        json_file_name = f"{process_name}-{timestamp}.json"

        st.download_button(
            label="定型フォームをJSON保存",
            data=json_bytes,
            file_name=json_file_name,
            mime="application/json",
            use_container_width=True,
            disabled=not st.session_state.selected_steps,
        )

        if not st.session_state.selected_steps:
            st.caption("※手順を読み込む、または追加すると保存できます。")

    with col_json2:
        st.markdown("#### 定型フォームを読込")

        uploaded_json = st.file_uploader(
            "保存済みJSONを選択",
            type=["json"],
            key="template_json_uploader",
        )

        if uploaded_json is not None:
            if st.button("JSONを読み込む", use_container_width=True):
                try:
                    data = json.loads(uploaded_json.getvalue().decode("utf-8"))
                    apply_template_json_data(data)
                    st.success("定型フォームを読み込みました。")
                    st.rerun()
                except Exception as e:
                    st.error(f"JSONの読み込みに失敗しました: {e}")

    with col1:
        st.session_state.manual_title = st.text_input(
            "手順書名",
            value=st.session_state.manual_title,
            placeholder="例：製品A マシニング加工手順",
        )

        process_options = list(PROCESS_TEMPLATES.keys())
        st.session_state.process_type = st.selectbox(
            "工程分類",
            options=process_options,
            index=process_options.index(st.session_state.process_type)
            if st.session_state.process_type in process_options
            else 0,
        )

        st.session_state.equipment_name = st.text_input(
            "対象設備",
            value=st.session_state.equipment_name,
            placeholder="例：マシニングセンタ1号機",
        )

        st.session_state.product_name = st.text_input(
            "製品名",
            value=st.session_state.product_name,
            placeholder="例：ブラケットA",
        )

        st.session_state.part_number = st.text_input(
            "品番",
            value=st.session_state.part_number,
            placeholder="例：ABC-123",
        )

    with col2:
        st.session_state.drawing_number = st.text_input(
            "図面番号",
            value=st.session_state.drawing_number,
            placeholder="例：D-4567",
        )

        st.session_state.program_number = st.text_input(
            "加工プログラム番号",
            value=st.session_state.program_number,
            placeholder="例：O1234",
        )

        st.session_state.jig_name = st.text_input(
            "使用治具",
            value=st.session_state.jig_name,
            placeholder="例：JIG-01",
        )

        st.session_state.tool_name = st.text_input(
            "使用工具",
            value=st.session_state.tool_name,
            placeholder="例：T01, T02, ノギス",
        )

        st.session_state.author = st.text_input(
            "作成者",
            value=st.session_state.author,
            placeholder="例：山田太郎",
        )

        st.session_state.revision = st.text_input(
            "改訂番号",
            value=st.session_state.revision,
            placeholder="例：Rev.0",
        )

st.markdown("---")
st.subheader("1. 標準手順テンプレートの選択")

template_steps = PROCESS_TEMPLATES.get(
    st.session_state.process_type,
    PROCESS_TEMPLATES["汎用テンプレート"],
)

col_load1, col_load2 = st.columns([1, 2])

with col_load1:
    if st.button("テンプレートを読込", use_container_width=True):
        st.session_state.selected_steps = template_steps.copy()

        # 各手順の詳細情報を初期化
        st.session_state.step_details = {}
        for step in st.session_state.selected_steps:
            st.session_state.step_details[step] = {
                "point": [],
                "caution": [],
                "check": [],
                "free_point": "",
                "free_caution": "",
                "free_check": "",
                "image": None,
                "image_name": "",
            }

        st.success("テンプレートを読み込みました。")

with col_load2:
    st.caption("工程分類を選んでから「テンプレートを読込」を押してください。不要な手順は後で外せます。")

if st.session_state.selected_steps:
    st.markdown("---")
    st.subheader("2. 使用する手順を選択")

    st.caption("不要な手順はチェックを外してください。")

    # 現在の選択順を優先する
    # これにより「上へ」「下へ」で並べ替えた順番を維持する
    all_step_options = st.session_state.selected_steps + [
        x for x in template_steps if x not in st.session_state.selected_steps
    ]
    
    new_selected_steps = []
    
    for idx, step in enumerate(all_step_options, start=1):
        checked = step in st.session_state.selected_steps
    
        if st.checkbox(
            f"{idx}. {step}",
            value=checked,
            key=f"step_select_{safe_filename(step)}",
        ):
            new_selected_steps.append(step)
    
    st.session_state.selected_steps = new_selected_steps

    # =========================
    # 手順の並べ替え
    # =========================
    if st.session_state.selected_steps:
        st.markdown("### 手順の並べ替え")
        st.caption("追加した手順や標準手順の順番を変更できます。Excel出力もこの順番になります。")

        for idx, step in enumerate(st.session_state.selected_steps):
            col_no, col_step, col_up, col_down = st.columns([0.5, 5, 1, 1])

            with col_no:
                st.write(idx + 1)

            with col_step:
                st.write(step)

            with col_up:
                if st.button(
                    "上へ",
                    key=f"move_up_{idx}_{safe_filename(step)}",
                    use_container_width=True,
                    disabled=(idx == 0),
                ):
                    move_step_up(idx)
                    st.rerun()

            with col_down:
                if st.button(
                    "下へ",
                    key=f"move_down_{idx}_{safe_filename(step)}",
                    use_container_width=True,
                    disabled=(idx == len(st.session_state.selected_steps) - 1),
                ):
                    move_step_down(idx)
                    st.rerun()

    # step_details に存在しない手順があれば追加
    for step in st.session_state.selected_steps:
        if step not in st.session_state.step_details:
            st.session_state.step_details[step] = {
                "point": [],
                "caution": [],
                "check": [],
                "free_point": "",
                "free_caution": "",
                "free_check": "",
                "image": None,
                "image_name": "",
            }

else:
    st.warning("まず工程分類を選び、「テンプレートを読込」を押してください。")

st.markdown("### 手順を追加")

add_col1, add_col2 = st.columns([3, 1])

with add_col1:
    st.session_state.extra_step_input = st.text_input(
        "候補にない手順を追加",
        value=st.session_state.extra_step_input,
        placeholder="例：加工前にエアブローで切粉を除去する",
    )

with add_col2:
    st.write("")
    st.write("")
    if st.button("追加", use_container_width=True):
        new_step = st.session_state.extra_step_input.strip()

        if new_step:
            if new_step not in st.session_state.selected_steps:
                st.session_state.selected_steps.append(new_step)
                st.session_state.step_details[new_step] = {
                    "point": [],
                    "caution": [],
                    "check": [],
                    "free_point": "",
                    "free_caution": "",
                    "free_check": "",
                    "image": None,
                    "image_name": "",
                }
                st.session_state.extra_step_input = ""
                st.success("手順を追加しました。")
                st.rerun()
            else:
                st.info("すでに追加されています。")
        else:
            st.warning("追加する手順を入力してください。")

# =========================
# 各手順の詳細入力
# =========================

if st.session_state.selected_steps:
    st.markdown("---")
    st.subheader("3. 各手順の詳細入力")

    st.caption("写真、ポイント、注意事項、確認項目を必要に応じて設定してください。空欄でもExcel出力できます。")

    for idx, step in enumerate(st.session_state.selected_steps, start=1):
        if step not in st.session_state.step_details:
            st.session_state.step_details[step] = {
                "point": [],
                "caution": [],
                "check": [],
                "free_point": "",
                "free_caution": "",
                "free_check": "",
                "image": None,
                "image_name": "",
            }

        detail = st.session_state.step_details[step]

        with st.expander(f"手順{idx}：{step}", expanded=False):
            uploaded_file = st.file_uploader(
                "写真を選択または撮影",
                type=["jpg", "jpeg", "png"],
                key=f"image_{idx}_{step}",
            )

            if uploaded_file is not None:
                detail["image"] = uploaded_file.getvalue()
                detail["image_name"] = uploaded_file.name

                try:
                    preview_image = PILImage.open(io.BytesIO(detail["image"]))
                    st.image(preview_image, caption="登録写真", use_container_width=True)
                except Exception:
                    st.warning("写真のプレビューに失敗しました。")

            detail["point"] = st.multiselect(
                "ポイント（選択）",
                options=POINT_OPTIONS,
                default=detail.get("point", []),
                key=f"point_{idx}_{step}",
            )

            detail["free_point"] = st.text_input(
                "ポイント（自由入力）",
                value=detail.get("free_point", ""),
                placeholder="例：赤丸部を基準面に当てる",
                key=f"free_point_{idx}_{step}",
            )

            detail["caution"] = st.multiselect(
                "注意事項（選択）",
                options=CAUTION_OPTIONS,
                default=detail.get("caution", []),
                key=f"caution_{idx}_{step}",
            )

            detail["free_caution"] = st.text_input(
                "注意事項（自由入力）",
                value=detail.get("free_caution", ""),
                placeholder="例：A面はキズ禁止",
                key=f"free_caution_{idx}_{step}",
            )

            detail["check"] = st.multiselect(
                "確認項目（選択）",
                options=CHECK_OPTIONS,
                default=detail.get("check", []),
                key=f"check_{idx}_{step}",
            )

            detail["free_check"] = st.text_input(
                "確認項目（自由入力）",
                value=detail.get("free_check", ""),
                placeholder="例：クランプピンに確実に当たっていること",
                key=f"free_check_{idx}_{step}",
            )

            st.session_state.step_details[step] = detail

# =========================
# Excel出力
# =========================

def join_items(selected_items: List[str], free_text: str) -> str:
    items = list(selected_items) if selected_items else []
    if free_text and free_text.strip():
        items.append(free_text.strip())
    return "、".join(items)


def add_cell(ws, row: int, col: int, value, fill=None, font=None, align=None, border=True):
    cell = ws.cell(row=row, column=col, value=value)
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    if border:
        cell.border = thin_border
    return cell


def create_excel_bytes() -> bytes:
    wb = Workbook()

    # =========================
    # 罫線スタイル
    # =========================
    thin_side = Side(style="thin", color="000000")
    medium_side = Side(style="medium", color="000000")
    thick_side = Side(style="thick", color="000000")

    def apply_box_border(ws, start_row, start_col, end_row, end_col, outer_side, inner_side):
        """
        指定範囲に罫線を付ける。
        外枠は outer_side、内側は inner_side。
        """
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)

                left = outer_side if c == start_col else inner_side
                right = outer_side if c == end_col else inner_side
                top = outer_side if r == start_row else inner_side
                bottom = outer_side if r == end_row else inner_side

                cell.border = Border(
                    left=left,
                    right=right,
                    top=top,
                    bottom=bottom,
                )

    # =========================
    # シート1：作業手順書
    # =========================
    ws = wb.active
    ws.title = "作業手順書"

    # ページ設定
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 列幅
    column_widths = {
        "A": 5,
        "B": 14,
        "C": 34,
        "D": 28,
        "E": 28,
        "F": 24,
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # タイトル
    ws.merge_cells("A1:F1")
    title = safe_value(st.session_state.manual_title)
    add_cell(
        ws,
        1,
        1,
        f"作業手順書：{title}",
        fill=title_fill,
        font=Font(bold=True, size=16),
        align=center_align,
    )
    ws.row_dimensions[1].height = 28

    # 基本情報
    info_rows = [
        ("工程", safe_value(st.session_state.process_type), "対象設備", safe_value(st.session_state.equipment_name)),
        ("製品名", safe_value(st.session_state.product_name), "品番", safe_value(st.session_state.part_number)),
        ("図面番号", safe_value(st.session_state.drawing_number), "プログラム番号", safe_value(st.session_state.program_number)),
        ("使用治具", safe_value(st.session_state.jig_name), "使用工具", safe_value(st.session_state.tool_name)),
        ("作成者", safe_value(st.session_state.author), "改訂番号", safe_value(st.session_state.revision)),
        ("作成日", datetime.now().strftime("%Y-%m-%d"), "承認", ""),
    ]
    
    start_row = 3
    
    for idx, (k1, v1, k2, v2) in enumerate(info_rows, start=start_row):
        # 左側：項目名 A:B
        ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=2)
        add_cell(
            ws,
            idx,
            1,
            k1,
            fill=thin_gray_fill,
            font=Font(bold=True),
            align=center_align,
        )
    
        # 左側：値 C
        add_cell(
            ws,
            idx,
            3,
            v1,
            align=left_align,
        )
    
        # 右側：項目名 D
        add_cell(
            ws,
            idx,
            4,
            k2,
            fill=thin_gray_fill,
            font=Font(bold=True),
            align=center_align,
        )
    
        # 右側：値 E:F
        ws.merge_cells(start_row=idx, start_column=5, end_row=idx, end_column=6)
        add_cell(
            ws,
            idx,
            5,
            v2,
            align=left_align,
        )

    # 手順表ヘッダー
    header_row = 10
    headers = ["No", "写真", "作業手順", "ポイント", "注意事項", "確認項目"]

    for col_idx, header in enumerate(headers, start=1):
        add_cell(
            ws,
            header_row,
            col_idx,
            header,
            fill=header_fill,
            font=Font(bold=True),
            align=center_align,
        )

    # 手順データ
    row = header_row + 1
    image_files_for_photo_sheet = []

    max_main_steps = 8  # 本体1ページ想定。多い場合は下に続く。

    for idx, step in enumerate(st.session_state.selected_steps, start=1):
        detail = st.session_state.step_details.get(step, {})

        point_text = join_items(detail.get("point", []), detail.get("free_point", ""))
        caution_text = join_items(detail.get("caution", []), detail.get("free_caution", ""))
        check_text = join_items(detail.get("check", []), detail.get("free_check", ""))

        add_cell(ws, row, 1, idx, align=center_align)
        add_cell(ws, row, 2, "", align=center_align)
        add_cell(ws, row, 3, step, align=top_left_align)
        add_cell(ws, row, 4, point_text, align=top_left_align)
        add_cell(ws, row, 5, caution_text, align=top_left_align)
        add_cell(ws, row, 6, check_text, align=top_left_align)

        ws.row_dimensions[row].height = 72

        image_bytes = detail.get("image")
        image_name = detail.get("image_name", "")

        if image_bytes:
            try:
                img_for_excel = XLImage(io.BytesIO(image_bytes))
                img_for_excel.width = 90
                img_for_excel.height = 60
                ws.add_image(img_for_excel, f"B{row}")

                image_files_for_photo_sheet.append(
                    {
                        "no": idx,
                        "step": step,
                        "image": image_bytes,
                        "image_name": image_name,
                        "point": point_text,
                        "caution": caution_text,
                        "check": check_text,
                    }
                )
            except Exception:
                add_cell(ws, row, 2, "画像読込不可", align=center_align)

        row += 1

    # =========================
    # 罫線の見栄え調整
    # 外枠：太線
    # 見出し・大区切り：中太線
    # 内部：細線
    # =========================

    last_row = max(row - 1, header_row + 1)

    # タイトル行：外枠太線
    apply_box_border(
        ws,
        start_row=1,
        start_col=1,
        end_row=1,
        end_col=6,
        outer_side=thick_side,
        inner_side=thin_side,
    )

    # 基本情報ブロック：外枠太線、内部細線
    apply_box_border(
        ws,
        start_row=3,
        start_col=1,
        end_row=8,
        end_col=6,
        outer_side=thick_side,
        inner_side=thin_side,
    )

    # 手順表全体：外枠太線、内部細線
    apply_box_border(
        ws,
        start_row=header_row,
        start_col=1,
        end_row=last_row,
        end_col=6,
        outer_side=thick_side,
        inner_side=thin_side,
    )

    # 手順表ヘッダー行：中太線
    apply_box_border(
        ws,
        start_row=header_row,
        start_col=1,
        end_row=header_row,
        end_col=6,
        outer_side=medium_side,
        inner_side=medium_side,
    )
    
    # 印刷範囲
    ws.print_area = f"A1:F{last_row}"

    # =========================
    # シート2：写真集
    # =========================
    ws_photo = wb.create_sheet("写真集")

    ws_photo.page_setup.orientation = "landscape"
    ws_photo.page_setup.paperSize = ws_photo.PAPERSIZE_A4
    ws_photo.page_setup.fitToWidth = 1
    ws_photo.page_setup.fitToHeight = 0
    ws_photo.sheet_properties.pageSetUpPr.fitToPage = True

    ws_photo.column_dimensions["A"].width = 8
    ws_photo.column_dimensions["B"].width = 48
    ws_photo.column_dimensions["C"].width = 60

    ws_photo.merge_cells("A1:C1")
    add_cell(
        ws_photo,
        1,
        1,
        f"写真集：{title}",
        fill=title_fill,
        font=Font(bold=True, size=16),
        align=center_align,
    )
    ws_photo.row_dimensions[1].height = 28

    photo_header_row = 3
    for col_idx, header in enumerate(["手順No", "写真", "説明・注意点"], start=1):
        add_cell(
            ws_photo,
            photo_header_row,
            col_idx,
            header,
            fill=header_fill,
            font=Font(bold=True),
            align=center_align,
        )

    photo_row = photo_header_row + 1

    if image_files_for_photo_sheet:
        for item in image_files_for_photo_sheet:
            ws_photo.row_dimensions[photo_row].height = 180

            add_cell(ws_photo, photo_row, 1, item["no"], align=center_align)
            add_cell(ws_photo, photo_row, 2, "", align=center_align)

            explanation = (
                f"作業手順：{item['step']}\n"
                f"ポイント：{item['point']}\n"
                f"注意事項：{item['caution']}\n"
                f"確認項目：{item['check']}"
            )
            add_cell(ws_photo, photo_row, 3, explanation, align=top_left_align)

            try:
                img_large = XLImage(io.BytesIO(item["image"]))
                img_large.width = 320
                img_large.height = 220
                ws_photo.add_image(img_large, f"B{photo_row}")
            except Exception:
                add_cell(ws_photo, photo_row, 2, "画像読込不可", align=center_align)

            photo_row += 1
    else:
        add_cell(ws_photo, photo_row, 1, "写真は登録されていません。", align=left_align)
        ws_photo.merge_cells(start_row=photo_row, start_column=1, end_row=photo_row, end_column=3)

    # =========================
    # シート3：入力データ
    # =========================
    ws_data = wb.create_sheet("入力データ")

    data_headers = [
        "No",
        "作業手順",
        "ポイント",
        "注意事項",
        "確認項目",
        "画像ファイル名",
    ]

    for col_idx, header in enumerate(data_headers, start=1):
        add_cell(
            ws_data,
            1,
            col_idx,
            header,
            fill=header_fill,
            font=Font(bold=True),
            align=center_align,
        )

    for idx, step in enumerate(st.session_state.selected_steps, start=1):
        detail = st.session_state.step_details.get(step, {})
        point_text = join_items(detail.get("point", []), detail.get("free_point", ""))
        caution_text = join_items(detail.get("caution", []), detail.get("free_caution", ""))
        check_text = join_items(detail.get("check", []), detail.get("free_check", ""))

        values = [
            idx,
            step,
            point_text,
            caution_text,
            check_text,
            detail.get("image_name", ""),
        ]

        for col_idx, value in enumerate(values, start=1):
            add_cell(ws_data, idx + 1, col_idx, value, align=top_left_align)

    for col_idx in range(1, len(data_headers) + 1):
        ws_data.column_dimensions[get_column_letter(col_idx)].width = 24

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# =========================
# 出力
# =========================

st.markdown("---")
st.header("4. Excel出力")
st.info("ここからExcel作業手順書を出力できます。")

if not st.session_state.selected_steps:
    st.info("手順を選択するとExcel出力できます。")
else:
    excel_bytes = create_excel_bytes()

    file_title = safe_value(st.session_state.manual_title).replace(" ", "_")
    file_name = f"{datetime.now().strftime('%Y-%m-%d_%H%M')}_{file_title}.xlsx"

    st.download_button(
        label="Excel作業手順書を出力する",
        data=excel_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_excel_bottom",
    )
