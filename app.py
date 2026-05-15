import os
import json
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
import html
import streamlit.components.v1 as components

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

APP_TITLE = "真因問診AI"
APP_SUBTITLE = "AIが8割の一般要因を整理し、現場の2割の違和感を引き出す"


# =========================
# AI呼び出し
# =========================
def call_ai(prompt: str) -> str:
    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt,
    )
    return response.output_text


# =========================
# JSON安全読み込み
# =========================
def parse_json_safely(text: str) -> dict:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(text[start:end + 1])
        raise


# =========================
# 初回プロンプト
# =========================
def build_first_prompt(problem: str, context: str) -> str:
    return f"""
あなたは、製造現場やビジネスの課題解決を知り尽くした「熟練の診断医」です。
ユーザーが入力した課題に対し、特性要因図の下書きと、真因を引き出す鋭い問いを1つだけ作成してください。

# 課題
{problem}

# 現場・業種の補足
{context}

# 重要方針
このアプリの目的は、特性要因図を完成させることではありません。
AIが一般的な要因を整理し、その後の問診によって、現場固有の違和感・差・例外を引き出すことです。

# 出力条件
- 回答は日本語。
- 一般論的な要因を4Mで各6〜8個出す。
- 4Mは以下。
  - Man（人）
  - Machine（機械・設備）
  - Material（材料・情報・対象物）
  - Method（方法・手順・条件）
- その後、「名医の問診」として質問を1つだけ出す。
- 質問は、次の3つのどれかの型から最も有効なものを選ぶ。
  1. 比較：「〇〇の時は起きるが、△△の時は起きない」という差はありますか？
  2. 逆説：「もしわざとこの問題を再現させるとしたら、何を操作しますか？」
  3. 類推：「この現象を料理や自然現象に例えると、何が起きているように見えますか？」
- 質問は1つだけ。
- 質問は、現場の人が「あっ」と思いやすいものにする。

# 出力形式
必ず次のJSON形式だけで返してください。説明文は不要です。

{{
  "problem": "課題名",
  "fishbone": {{
    "Man": ["要因1", "要因2", "要因3"],
    "Machine": ["要因1", "要因2", "要因3"],
    "Material": ["要因1", "要因2"],
    "Method": ["要因1", "要因2", "要因3"]
  }},
  "doctor_question": {{
    "type": "比較 または 逆説 または 類推",
    "question": "鋭い質問を1つ"
  }}
}}
"""


# =========================
# 継続問診プロンプト
# =========================
def build_followup_prompt(
    problem: str,
    context: str,
    first_result: dict,
    interview_history: list,
) -> str:
    return f"""
あなたは、製造現場やビジネスの課題解決を知り尽くした「熟練の診断医」です。
以下の課題、特性要因図の下書き、これまでの問診履歴をもとに、現時点の真因候補と、次の問診を1つだけ出してください。

# 課題
{problem}

# 現場・業種の補足
{context}

# 特性要因図の下書き
{json.dumps(first_result, ensure_ascii=False, indent=2)}

# これまでの問診履歴
{json.dumps(interview_history, ensure_ascii=False, indent=2)}

# 最重要方針
このアプリの価値は、1回の回答で真因を決め打ちすることではありません。
質問と回答を繰り返し、現場固有の「差」「例外」「再現条件」「違和感」を少しずつ掘り出すことにあります。

# 判断ルール
- まだ真因が確定できない場合は、「真因候補」として仮説を2〜3個出す。
- 無理に断定しない。
- ユーザー回答から見える現場固有の情報を最重視する。
- 次の問診は、これまでの回答で最も曖昧な点、または真因に近づきそうな点を突く。
- 次の問診は1つだけ。
- 似た質問を繰り返さない。
- 質問は、以下の3つの型から最も有効なものを選ぶ。
  1. 比較：「〇〇の時は起きるが、△△の時は起きない」という差はありますか？
  2. 逆説：「もしわざとこの問題を再現させるとしたら、何を操作しますか？」
  3. 類推：「この現象を料理や自然現象に例えると、何が起きているように見えますか？」

# 出力条件
- 真因候補は2〜3個。
- 各候補について、疑わしい理由を短く書く。
- 次に確認すべきことを、現場で実行できる観察・データ確認に落とし込む。
- 最後に、次の問診を1つだけ出す。

# 出力形式
必ず次のJSON形式だけで返してください。説明文は不要です。

{{
  "root_cause_candidates": [
    {{
      "candidate": "真因候補",
      "reason": "疑わしい理由",
      "next_check": "次に確認すべきこと"
    }}
  ],
  "next_doctor_question": {{
    "type": "比較 または 逆説 または 類推",
    "question": "次の鋭い質問を1つ"
  }}
}}
"""
def build_report_prompt(
    problem: str,
    context: str,
    first_result: dict,
    interview_history: list,
    latest_analysis: dict,
) -> str:
    return f"""
あなたは、製造現場やビジネスの課題解決を知り尽くした「熟練の診断医」です。
以下の情報をもとに、現時点での真因探索レポートを作成してください。

# 課題
{problem}

# 現場・業種の補足
{context}

# 特性要因図の下書き
{json.dumps(first_result, ensure_ascii=False, indent=2)}

# 問診履歴
{json.dumps(interview_history, ensure_ascii=False, indent=2)}

# 最新の真因候補
{json.dumps(latest_analysis, ensure_ascii=False, indent=2)}

# 作成方針
- 断定しすぎない。
- ただし、現時点で最も疑わしい真因候補を1〜3個に絞る。
- 「なぜそう考えるか」を、問診履歴に基づいて説明する。
- 次に現場で確認すべきことを、具体的な行動にする。
- 現場リーダーがそのまま使える簡潔な診断書にする。

# 出力形式
Markdownで、以下の見出しに従って出力してください。

## 真因探索レポート

### 1. 解決したい課題

### 2. 現時点で最も疑わしい真因候補

### 3. そう考える理由

### 4. まだ未確認の点

### 5. 最優先で確認すべきこと

### 6. 次の現場アクション
"""

def build_proposal_export_prompt(
    problem: str,
    context: str,
    first_result: dict,
    interview_history: list,
    latest_analysis: dict,
    final_report: str,
) -> str:
    return f"""
あなたは、製造業の現場改善と改善提案書作成に詳しいコンサルタントです。
以下の情報をもとに、「改善提案書作成AIへ渡すための連携データ」を作成してください。

# 課題
{problem}

# 現場・業種の補足
{context}

# 特性要因図の下書き
{json.dumps(first_result, ensure_ascii=False, indent=2)}

# 問診履歴
{json.dumps(interview_history, ensure_ascii=False, indent=2)}

# 最新の真因候補
{json.dumps(latest_analysis, ensure_ascii=False, indent=2)}

# 真因探索レポート
{final_report}

# 作成方針
- 改善提案書作成AIに渡す前提で、短く、整理された情報にする。
- そのままコピーして貼り付けられる形式にする。
- 現場改善の提案書につながるように、課題、現状、真因候補、確認事項、改善の方向性を整理する。
- 断定しすぎず、「真因候補」「改善余地」として表現する。
- 専門用語は使ってよいが、現場リーダーや工場長が読んで分かる言葉にする。

# 出力形式
必ず次のJSON形式だけで返してください。説明文は不要です。

{{
  "source_app": "真因問診AI",
  "improvement_target": "改善対象・課題名",
  "process": "対象工程・業務",
  "current_state": [
    "現状把握1",
    "現状把握2",
    "現状把握3"
  ],
  "problems": [
    "主な問題点1",
    "主な問題点2",
    "主な問題点3"
  ],
  "root_cause_candidates": [
    "真因候補・改善余地1",
    "真因候補・改善余地2",
    "真因候補・改善余地3"
  ],
  "recommended_checks": [
    "次に確認すべきこと1",
    "次に確認すべきこと2",
    "次に確認すべきこと3"
  ],
  "proposal_keywords": [
    "不良低減",
    "再発防止",
    "標準化"
  ],
  "one_line_summary": "改善提案書に使える1行要約"
}}
"""

def format_proposal_export_text(export_data: dict) -> str:
    def lines(items):
        if not items:
            return "・未整理"
        return "\n".join([f"・{item}" for item in items])

    return f"""【出力元アプリ】
{export_data.get("source_app", "真因問診AI")}

【改善対象】
{export_data.get("improvement_target", "")}

【対象工程・業務】
{export_data.get("process", "")}

【現状把握】
{lines(export_data.get("current_state", []))}

【主な問題点】
{lines(export_data.get("problems", []))}

【真因候補・改善余地】
{lines(export_data.get("root_cause_candidates", []))}

【次に確認すべきこと】
{lines(export_data.get("recommended_checks", []))}

【改善提案書で強調したい点】
{", ".join(export_data.get("proposal_keywords", []))}

【1行要約】
{export_data.get("one_line_summary", "")}
"""
def display_proposal_export_block(export_text: str, export_json: dict):
    st.divider()
    st.subheader("7. 【次のステップ：改善提案書作成】")

    st.caption("下のテキストをコピーして、改善提案書作成AIに貼り付けてください。")

    st.text_area(
        "提案書用連携テキスト",
        value=export_text,
        height=320,
    )

    st.download_button(
        label="提案書用JSONをダウンロード",
        data=json.dumps(export_json, ensure_ascii=False, indent=2),
        file_name="proposal_input_from_cause_doctor.json",
        mime="application/json",
    )

# =========================
# 表示：4M
# =========================

def svg_text(x, y, text, size=14, weight="normal", anchor="start", color="#222", max_chars=14, line_gap=18):
    text = text or ""
    lines = [text[i:i+max_chars] for i in range(0, len(text), max_chars)] or [""]
    tspans = []
    for i, line in enumerate(lines):
        dy = 0 if i == 0 else line_gap
        tspans.append(
            f'<tspan x="{x}" dy="{dy}">{html.escape(line)}</tspan>'
        )
    return f'''
    <text x="{x}" y="{y}" font-size="{size}" font-weight="{weight}"
          text-anchor="{anchor}" fill="{color}" font-family="'Yu Gothic', 'Meiryo', sans-serif">
        {''.join(tspans)}
    </text>
    '''


def render_fishbone_svg(problem: str, fishbone: dict):
    MAX_ITEMS_IN_DIAGRAM = 8

    width = 1650
    height = 1080
    spine_y = 540

    def esc(s):
        return html.escape(str(s or ""))

    def draw_multiline_text(
        x, y, text,
        size=16,
        weight="normal",
        color="#111827",
        anchor="start",
        max_chars=18,
        line_gap=20
    ):
        text = str(text or "")
        lines = [text[i:i + max_chars] for i in range(0, len(text), max_chars)] or [""]
        tspans = []
        for i, line in enumerate(lines):
            dy = 0 if i == 0 else line_gap
            tspans.append(f'<tspan x="{x}" dy="{dy}">{esc(line)}</tspan>')

        return f'''
        <text x="{x}" y="{y}" font-size="{size}" font-weight="{weight}"
              fill="{color}" text-anchor="{anchor}"
              font-family="'Yu Gothic', 'Meiryo', sans-serif">
            {''.join(tspans)}
        </text>
        '''

    def interp(p1, p2, t):
        x = p1[0] + (p2[0] - p1[0]) * t
        y = p1[1] + (p2[1] - p1[1]) * t
        return x, y

    svg = []
    svg.append(f'''
    <svg width="100%" height="{height}" viewBox="0 0 {width} {height}"
         xmlns="http://www.w3.org/2000/svg">
      <style>
        .spine {{
            stroke:#333;
            stroke-width:4;
            fill:none;
        }}
        .mainbone {{
            stroke:#444;
            stroke-width:3;
            fill:none;
        }}
        .subbone {{
            stroke:#777;
            stroke-width:2;
            fill:none;
        }}
        .head {{
            fill:#f8fafc;
            stroke:#333;
            stroke-width:3;
        }}
        .label {{
            font-size:24px;
            font-weight:700;
            fill:#1f2937;
            font-family:'Yu Gothic','Meiryo',sans-serif;
        }}
        .sub {{
            font-size:14px;
            font-weight:700;
            fill:#6b7280;
            font-family:'Yu Gothic','Meiryo',sans-serif;
        }}
      </style>
    ''')

    # 背骨
    svg.append(f'<line class="spine" x1="220" y1="{spine_y}" x2="1280" y2="{spine_y}" />')
    svg.append(f'<polygon points="1280,{spine_y} 1245,{spine_y-18} 1245,{spine_y+18}" fill="#333" />')

    # 魚の頭
    svg.append(f'<ellipse class="head" cx="1440" cy="{spine_y}" rx="140" ry="78" />')
    svg.append(draw_multiline_text(
        1440, spine_y - 12, problem,
        size=22, weight="bold", anchor="middle",
        max_chars=10, line_gap=24
    ))

    # 4M設定
    # 左右をやや中央へ寄せ、高さを確保
    sections = {
        "Man": {
            "jp": "人",
            "root": (520, spine_y),
            "tip": (320, 90),
            "label": (70, 80),
            "text_width_end": 300,
            "side": "left",
            "vertical": "up",
        },
        "Machine": {
            "jp": "機械・設備",
            "root": (670, spine_y),
            "tip": (400, 980),
            "label": (100, 1010),
            "text_width_end": 330,
            "side": "left",
            "vertical": "down",
            # 必要に応じて個別調整可
            # 例: 0番目の小骨だけ短くする
            # "item_bx": {0: 360}
        },
        "Material": {
            "jp": "材料・情報・対象物",
            "root": (820, spine_y),
            "tip": (620, 90),
            "label": (950, 80),
            "text_width_end": 920,
            "side": "right",
            "vertical": "up",
        },
        "Method": {
            "jp": "方法・手順・条件",
            "root": (980, spine_y),
            "tip": (760, 980),
            "label": (1110, 1010),
            "text_width_end": 1080,
            "side": "right",
            "vertical": "down",
            # 必要に応じて個別調整可
            # "item_bx": {0: 1120},
            # "item_t": [0.20, 0.32, 0.44, 0.56, 0.68, 0.78, 0.86, 0.92]
        },
    }

    for key, cfg in sections.items():
        root = cfg["root"]
        tip = cfg["tip"]
        items = fishbone.get(key, [])[:MAX_ITEMS_IN_DIAGRAM]

        # 大骨
        svg.append(
            f'<line class="mainbone" x1="{root[0]}" y1="{root[1]}" x2="{tip[0]}" y2="{tip[1]}" />'
        )

        # ラベル
        lx, ly = cfg["label"]
        svg.append(f'<text class="label" x="{lx}" y="{ly}">{key}</text>')
        svg.append(f'<text class="sub" x="{lx}" y="{ly + 22}">{cfg["jp"]}</text>')

        if not items:
            continue

        n = len(items)

        # 項目数に応じて文字サイズ等を自動調整
        if n <= 4:
            font_size = 17
            max_chars = 18
            line_gap = 20
        elif n <= 6:
            font_size = 16
            max_chars = 17
            line_gap = 19
        else:
            font_size = 15
            max_chars = 16
            line_gap = 18

        # 小骨の取付位置（大骨上）
        if "item_t" in cfg:
            ts = cfg["item_t"][:n]
        else:
            if n == 1:
                ts = [0.5]
            else:
                # 端に寄りすぎないよう少し内側に均等配置
                start_t = 0.16
                end_t = 0.90
                step = (end_t - start_t) / (n - 1)
                ts = [start_t + i * step for i in range(n)]

            # 上側は、上に表示したい要因ほど骨の先端側へ
            if cfg["vertical"] == "up":
                ts = list(reversed(ts))

        # 小骨の個別長さ調整
        item_bx = cfg.get("item_bx", {})

        for i, (item, t) in enumerate(zip(items, ts)):
            ax, ay = interp(root, tip, t)

            # 小骨を少し下げる
            bone_y = ay + 8

            # デフォルト終点
            bx = cfg["text_width_end"]
            if i in item_bx:
                bx = item_bx[i]

            # 文字は小骨の外側に配置
            if cfg["side"] == "left":
                tx = bx - 10
                anchor = "end"
            else:
                tx = bx + 10
                anchor = "start"

            # 小骨（水平）
            svg.append(
                f'<line class="subbone" x1="{ax}" y1="{bone_y}" x2="{bx}" y2="{bone_y}" />'
            )

            # 要因テキスト
            svg.append(draw_multiline_text(
                tx, bone_y - 6, f'・{item}',
                size=font_size,
                weight="bold",
                anchor=anchor,
                max_chars=max_chars,
                line_gap=line_gap
            ))

    svg.append("</svg>")

    components.html("".join(svg), height=height + 20, scrolling=False)

def display_fishbone(problem: str, fishbone: dict):
    st.subheader("1. 【特性要因図】")
    render_fishbone_svg(problem, fishbone)

    with st.expander("4Mの要因一覧を見る"):
        cols = st.columns(4)
        labels = [
            ("Man", "人"),
            ("Machine", "機械・設備"),
            ("Material", "材料・情報・対象物"),
            ("Method", "方法・手順・条件"),
        ]

        for col, (key, label) in zip(cols, labels):
            with col:
                st.markdown(f"### {key}")
                st.caption(label)
                for item in fishbone.get(key, []):
                    st.markdown(f"- {item}")


# =========================
# 表示：問診
# =========================
def display_current_question(question_data: dict):
    st.subheader("2. 【名医の問診】")
    q_type = question_data.get("type", "")
    question = question_data.get("question", "")
    st.info(f"**{q_type}の問い**\n\n{question}")


# =========================
# 表示：真因候補
# =========================
def display_root_causes(result: dict):
    st.subheader("3. 【現時点の真因候補】")

    candidates = result.get("root_cause_candidates", [])

    if not candidates:
        st.caption("まだ真因候補はありません。")
        return

    for i, item in enumerate(candidates, start=1):
        st.markdown(f"### 候補{i}：{item.get('candidate', '')}")
        st.markdown(f"**疑わしい理由**：{item.get('reason', '')}")
        st.markdown(f"**次に確認すべきこと**：{item.get('next_check', '')}")


# =========================
# 表示：問診履歴
# =========================
def display_interview_history(history: list):
    if not history:
        return

    st.subheader("4. 【問診履歴】")

    for i, item in enumerate(history, start=1):
        st.markdown(f"#### 問診{i}")
        st.markdown(f"**質問**：{item.get('question', '')}")
        st.markdown(f"**回答**：{item.get('answer', '')}")


# =========================
# セッション初期化
# =========================
def init_session_state():
    if "first_result" not in st.session_state:
        st.session_state.first_result = None

    if "current_question" not in st.session_state:
        st.session_state.current_question = None

    if "latest_analysis" not in st.session_state:
        st.session_state.latest_analysis = None

    if "interview_history" not in st.session_state:
        st.session_state.interview_history = []

    if "round_no" not in st.session_state:
        st.session_state.round_no = 1

    if "final_report" not in st.session_state:
        st.session_state.final_report = None

    if "proposal_export_data" not in st.session_state:
        st.session_state.proposal_export_data = None

    if "proposal_export_text" not in st.session_state:
        st.session_state.proposal_export_text = None

# =========================
# リセット
# =========================
def reset_all():
    st.session_state.first_result = None
    st.session_state.current_question = None
    st.session_state.latest_analysis = None
    st.session_state.interview_history = []
    st.session_state.round_no = 1
    st.session_state.final_report = None
    st.session_state.proposal_export_data = None
    st.session_state.proposal_export_text = None

# =========================
# メイン
# =========================
def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🩺",
        layout="wide",
    )

    init_session_state()

    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    st.markdown("""
このアプリは、特性要因図をきれいに描くためのものではありません。  
AIが一般的な要因を素早く整理し、現場の人だけが知っている  
**「違和感」「差」「例外」「再現条件」**を引き出すための問診型ツールです。
""")

    st.divider()

    problem = st.text_input(
        "解決したい課題（特性）",
        placeholder="例：加工後のバリが多い、検査工程で手戻りが多い、納期遅れが頻発している",
    )

    context = st.text_area(
        "現場・業種の補足（任意）",
        placeholder="例：航空機部品の機械加工。少量多品種で、作業者ごとの経験差が大きい。",
        height=100,
    )

    col1, col2 = st.columns([1, 1])

    with col1:
        generate_button = st.button("特性要因図を作成する", type="primary")

    with col2:
        reset_button = st.button("最初からやり直す")

    if reset_button:
        reset_all()
        st.rerun()

    # 初回生成
    if generate_button:
        if not problem.strip():
            st.error("解決したい課題を入力してください。")
            return

        with st.spinner("AI診断医が、4Mで要因を整理しています..."):
            prompt = build_first_prompt(problem, context)
            raw = call_ai(prompt)
            first_result = parse_json_safely(raw)

            st.session_state.first_result = first_result
            st.session_state.current_question = first_result.get("doctor_question", {})
            st.session_state.latest_analysis = None
            st.session_state.interview_history = []
            st.session_state.round_no = 1

    # 初回結果表示
    if st.session_state.first_result:
        first_result = st.session_state.first_result

        st.success("特性要因図の下書きを作成しました。")

        display_fishbone(problem, first_result.get("fishbone", {}))

        # 最新の真因候補を表示
        if st.session_state.latest_analysis:
            st.divider()
            display_root_causes(st.session_state.latest_analysis)

        # 問診履歴を表示
        if st.session_state.interview_history:
            st.divider()
            display_interview_history(st.session_state.interview_history)

        st.divider()

        # 現在の問診を表示
        current_question = st.session_state.current_question
        if current_question:
            display_current_question(current_question)

            answer_key = f"answer_round_{st.session_state.round_no}"

            user_answer = st.text_area(
                "この問診への回答",
                key=answer_key,
                placeholder="例：A号機では多いがB号機では少ない。特に夜勤で多い。材料ロットが変わった後から増えた気がする。",
                height=120,
            )

            col_a, col_b = st.columns([1, 1])

            with col_a:
                continue_button = st.button("回答して、次の問診へ進む", type="primary")

            with col_b:
                report_button = st.button("ここまでで診断書を作成する")

            if continue_button:
                if not user_answer.strip():
                    st.error("問診への回答を入力してください。")
                    return

                # 現在のQ&Aを履歴に追加
                st.session_state.interview_history.append(
                    {
                        "round": st.session_state.round_no,
                        "type": current_question.get("type", ""),
                        "question": current_question.get("question", ""),
                        "answer": user_answer,
                    }
                )

                with st.spinner("回答を踏まえて、次の問診を考えています..."):
                    prompt = build_followup_prompt(
                        problem=problem,
                        context=context,
                        first_result=first_result,
                        interview_history=st.session_state.interview_history,
                    )

                    raw = call_ai(prompt)
                    analysis = parse_json_safely(raw)

                    st.session_state.latest_analysis = analysis
                    st.session_state.current_question = analysis.get("next_doctor_question", {})
                    st.session_state.round_no += 1

                st.rerun()

            if report_button:
                if not st.session_state.interview_history:
                    st.error("少なくとも1回は問診に回答してから診断書を作成してください。")
                    return

                with st.spinner("ここまでの問診履歴から、診断書を作成しています..."):
                    prompt = build_report_prompt(
                        problem=problem,
                        context=context,
                        first_result=first_result,
                        interview_history=st.session_state.interview_history,
                        latest_analysis=st.session_state.latest_analysis,
                    )

                    report = call_ai(prompt)
                    st.session_state.final_report = report

                    # 改善提案書用の連携データも同時に作成
                    export_prompt = build_proposal_export_prompt(
                        problem=problem,
                        context=context,
                        first_result=first_result,
                        interview_history=st.session_state.interview_history,
                        latest_analysis=st.session_state.latest_analysis,
                        final_report=report,
                    )

                    export_raw = call_ai(export_prompt)
                    export_data = parse_json_safely(export_raw)

                    st.session_state.proposal_export_data = export_data
                    st.session_state.proposal_export_text = format_proposal_export_text(export_data)

                st.rerun()

        if st.session_state.final_report:
            st.divider()
            st.subheader("5. 【真因探索レポート】")
            st.markdown(st.session_state.final_report)

        if st.session_state.proposal_export_text and st.session_state.proposal_export_data:
            display_proposal_export_block(
                st.session_state.proposal_export_text,
                st.session_state.proposal_export_data,
            )

if __name__ == "__main__":
    main()
