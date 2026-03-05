import streamlit as st
import anthropic
import json
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

st.set_page_config(
    page_title="GMP DocWriter | XI",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "lang" not in st.session_state:
    st.session_state.lang = "EN"

def t(en, zh):
    return zh if st.session_state.lang == "ZH" else en

DEFAULTS = {
    "step": 1,
    "doc_type": None,
    "basic": {},
    "event": {},
    "impact": {},
    "rca": {},
    "capa": {},
    "risk": {},
    "ai_suggestions": {},
    "chat_history": [],
    "ai_capa_proposal": "",
    "report_generated": False,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v if not isinstance(v, (dict, list)) else (v.copy() if isinstance(v, dict) else [])

@st.cache_resource
def get_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

def ask_claude(system_prompt, user_message):
    client = get_client()
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )
    return msg.content[0].text

SYSTEM_ENHANCE = """You are a GMP technical writing expert specialising in radiopharmaceutical manufacturing.
Your job is to improve deviation/investigation reports to be audit-ready and standalone-readable.
Key requirements:
- Every document reference must include document number AND version
- Every claim must be supported by specific data/numbers
- Timelines must be explicit (dates and times)
- Root cause must be justified with evidence
- CAPA must distinguish corrective vs preventive
- Risk scores must include justification for probability AND severity
- Product impact must address: process / product quality / patient safety separately
Improve the provided text. Return only the improved text, no commentary."""

SYSTEM_RCA = """You are a GMP root cause analysis expert for pharmaceutical manufacturing.
Guide the user through a rigorous 6M root cause analysis (Man, Method, Machine, Materials, Mother Nature, Measurement).
For each M, ask probing questions to either confirm or exclude it as a contributing factor.
Do NOT accept "not applicable" without justification.
If the user says "Man error", probe deeper: Was it lack of training? Lack of procedure? Lack of independent verification? Workload?
Always ask: "If this is the root cause, what specifically would the CAPA need to address to prevent recurrence?"
Be concise. Ask one focused question at a time."""

SYSTEM_RISK = """You are a GMP risk assessment expert.
Challenge risk scores that seem too low.
Ask the user to justify their probability and severity scores with evidence.
Be concise and direct."""

SYSTEM_CAPA = """You are a GMP CAPA expert for pharmaceutical manufacturing.
Based on the root cause analysis conversation provided, propose:
1. A Corrective Action (CA) - specifically addressing what happened in THIS event
2. A Preventive Action (PA) - specifically addressing the root cause to prevent recurrence
3. An Effectiveness Check method - how to verify the CAPA worked
Rules:
- CA and PA must be different
- PA must directly link to the identified root cause
- Each action must have a clear owner type (QA, Production, Training, etc.)
- Effectiveness check must be measurable and time-bound
Format your response clearly with CA, PA, and Effectiveness Check as separate sections."""

with st.sidebar:
    col1, col2 = st.columns(2)
    with col1:
        en_type = "primary" if st.session_state.lang == "EN" else "secondary"
        if st.button("English", use_container_width=True, type=en_type, key="lang_en"):
            st.session_state.lang = "EN"
    with col2:
        zh_type = "primary" if st.session_state.lang == "ZH" else "secondary"
        if st.button("中文", use_container_width=True, type=zh_type, key="lang_zh"):
            st.session_state.lang = "ZH"

    st.divider()
    st.markdown(f"### {t('Progress', '填寫進度')}")
    steps = [
        t("Doc Type", "文件類型"),
        t("Basic Info", "基本資訊"),
        t("Event Detail", "事件描述"),
        t("Impact", "影響評估"),
        t("Root Cause", "根本原因"),
        t("CAPA & Risk", "CAPA 與風險"),
    ]
    for i, s in enumerate(steps, 1):
        icon = "✅" if st.session_state.step
elif st.session_state.step == 3:
    st.subheader(f"Step 3 / 6 — {t('Event Description', '事件描述')}")
    e = st.session_state.event
    st.info(t(
        "📌 Every sentence must be independently verifiable. Include exact document references (number+version), specific values, names, exact timestamps.",
        "📌 每句話必須可以獨立驗證。包含精確的文件引用（編號+版本）、具體數值、相關人員姓名、精確時間戳。"
    ))
    e["description"] = st.text_area(
        t("Description of the Deviation (What / Where / When / Who / Extent)", "偏差描述"),
        value=e.get("description", ""), height=180)
    e["immediate_action"] = st.text_area(
        t("Immediate Actions Taken (with timestamp and person responsible)", "已採取的立即措施"),
        value=e.get("immediate_action", ""), height=100)
    e["extent"] = st.text_area(
        t("Scope / Extent of Impact", "影響範圍"),
        value=e.get("extent", ""), height=80)
    if st.button(t("🤖 AI: Improve My Description", "🤖 AI：改善我的描述"), key="enhance_desc"):
        if e.get("description"):
            with st.spinner(t("Improving...", "改善中...")):
                improved = ask_claude(SYSTEM_ENHANCE,
                    f"Improve this GMP deviation description. Context: {st.session_state.basic}\n\nText to improve:\n{e['description']}")
                e["description_ai"] = improved
                st.session_state.event = e
        else:
            st.warning(t("Please write something first.", "請先填寫內容。"))
    if e.get("description_ai"):
        st.success(t("✅ AI Suggestion:", "✅ AI 建議："))
        st.code(e["description_ai"], language=None)
    st.session_state.event = e
    col_back, col_next = st.columns(2)
    with col_back:
        if st.button(t("← Back", "← 上一步"), use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with col_next:
        if st.button(t("Next →", "下一步 →"), type="primary", use_container_width=True):
            if not e.get("description"):
                st.error(t("Please fill in the deviation description.", "請填寫偏差描述。"))
            else:
                st.session_state.step = 4
                st.rerun()

elif st.session_state.step == 4:
    st.subheader(f"Step 4 / 6 — {t('Impact Assessment', '影響評估')}")
    imp = st.session_state.impact
    st.warning(t(
        "⚠️ Common error: Writing only 'no impact to product quality' without explanation. You must separately address all three levels below.",
        "⚠️ 常見錯誤：只寫「對產品品質無影響」而不加解釋。您必須分別說明以下三個層面。"
    ))
    imp["process_impact"] = st.text_area(
        t("1. Process Impact (what happened to the manufacturing process?)", "1. 製程影響"),
        value=imp.get("process_impact", ""), height=100)
    imp["quality_impact"] = st.text_area(
        t("2. Product Quality Impact (were any release specifications affected?)", "2. 產品品質影響"),
        value=imp.get("quality_impact", ""), height=100)
    imp["patient_impact"] = st.text_area(
        t("3. Patient Safety Impact (was any released product affected? How many patients?)", "3. 病患安全影響"),
        value=imp.get("patient_impact", ""), height=100)
    if st.session_state.doc_type == "PRDI":
        imp["disposition"] = st.selectbox(
            t("Batch Disposition Decision", "批次處置決定"),
            [t("Release — all specs met", "放行 — 所有規格符合"),
             t("Reject — spec failure", "拒絕 — 規格不符"),
             t("Conditional release with QA justification", "條件放行")])
        imp["disposition_rationale"] = st.text_area(t("Disposition Rationale", "處置理由"), value=imp.get("disposition_rationale", ""), height=80)
    if st.button(t("🤖 AI: Challenge My Impact Assessment", "🤖 AI：挑戰我的影響評估"), key="challenge_impact"):
        text = f"Process: {imp.get('process_impact','')}\nQuality: {imp.get('quality_impact','')}\nPatient: {imp.get('patient_impact','')}"
        if text.strip():
            with st.spinner(t("Analysing...", "分析中...")):
                challenge = ask_claude(SYSTEM_RISK,
                    f"Review this impact assessment and ask probing questions if anything is vague or missing:\n{text}")
                st.info(f"🤖 {challenge}")
        else:
            st.warning(t("Please fill in the impact fields first.", "請先填寫影響評估欄位。"))
    st.session_state.impact = imp
    col_back, col_next = st.columns(2)
    with col_back:
        if st.button(t("← Back", "← 上一步"), use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    with col_next:
        if st.button(t("Next →", "下一步 →"), type="primary", use_container_width=True):
            if not imp.get("process_impact"):
                st.error(t("Please complete the impact assessment.", "請完成影響評估。"))
            else:
                st.session_state.step = 5
                st.rerun()

elif st.session_state.step == 5:
    st.subheader(f"Step 5 / 6 — {t('Root Cause Analysis', '根本原因分析')}")
    rca = st.session_state.rca
    st.info(t(
        "📌 RCA Rule: You must EITHER confirm OR explicitly exclude each 6M category with evidence.",
        "📌 RCA 原則：您必須對每個 6M 類別提供確認或明確排除的依據。"
    ))
    ms_en = ["Man", "Method", "Machine", "Materials", "Mother Nature", "Measurement"]
    ms_zh = ["人員", "方法", "機器", "物料", "環境", "測量"]
    for i, (m_en, m_zh) in enumerate(zip(ms_en, ms_zh)):
        with st.expander(f"{'✅' if rca.get(f'm_{m_en}_status') else '⚪'} {m_en} {m_zh}", expanded=(i == 0)):
            status_options = [
                t("Confirmed root cause", "確認為根本原因"),
                t("Contributing factor", "促成因素"),
                t("Excluded — with evidence", "已排除（附依據）"),
                t("Under investigation", "調查中")]
            rca[f"m_{m_en}_status"] = st.selectbox(
                t("Status", "狀態"), status_options, key=f"status_{m_en}",
                index=status_options.index(rca.get(f"m_{m_en}_status", t("Excluded — with evidence", "已排除（附依據）"))))
            rca[f"m_{m_en}_detail"] = st.text_area(
                t(f"Evidence / Justification for {m_en}", f"{m_zh} — 依據／理由"),
                value=rca.get(f"m_{m_en}_detail", ""), height=80, key=f"detail_{m_en}")
    rca["root_cause_statement"] = st.text_area(
        t("Final Root Cause Statement (link evidence → cause → CAPA)", "最終根本原因陳述"),
        value=rca.get("root_cause_statement", ""), height=120)

    st.divider()
    st.markdown(f"#### 🤖 {t('AI RCA Coach', 'AI 根本原因分析輔導')}")
    user_msg = st.text_input(t("Describe your root cause or ask AI to probe further:", "描述您的根本原因或請 AI 深入探討："), key="rca_chat_input")
    if st.button(t("Ask AI", "詢問 AI"), key="ask_rca"):
        if user_msg:
            context = f"Event: {st.session_state.basic.get('title','')}\nDescription: {st.session_state.event.get('description','')}\nUser says: {user_msg}"
            with st.spinner(t("Thinking...", "思考中...")):
                response = ask_claude(SYSTEM_RCA, context)
                st.session_state.chat_history.append(("user", user_msg))
                st.session_state.chat_history.append(("ai", response))
    for role, msg in st.session_state.chat_history[-8:]:
        if role == "ai":
            st.info(f"🤖 {msg}")
        else:
            st.chat_message("user").write(msg)

    if len(st.session_state.chat_history) >= 2:
        st.divider()
        if st.button(t("🤖 AI: Propose CAPA based on this RCA", "🤖 AI：根據 RCA 提議 CAPA"), key="propose_capa", type="primary"):
            chat_summary = "\n".join([f"{role.upper()}: {msg}" for role, msg in st.session_state.chat_history])
            context = f"Event Title: {st.session_state.basic.get('title','')}\nRoot Cause Statement: {rca.get('root_cause_statement','')}\nRCA Conversation:\n{chat_summary}"
            with st.spinner(t("Generating CAPA proposal...", "生成 CAPA 建議中...")):
                proposal = ask_claude(SYSTEM_CAPA, context)
                st.session_state.ai_capa_proposal = proposal
        if st.session_state.ai_capa_proposal:
            st.success(f"🤖 **{t('AI CAPA Proposal — copy to Step 6', 'AI CAPA 建議 — 複製到第 6 步')}**")
            st.markdown(st.session_state.ai_capa_proposal)

    st.session_state.rca = rca
    col_back, col_next = st.columns(2)
    with col_back:
        if st.button(t("← Back", "← 上一步"), use_container_width=True):
            st.session_state.step = 4
            st.rerun()
    with col_next:
        if st.button(t("Next →", "下一步 →"), type="primary", use_container_width=True):
            if not rca.get("root_cause_statement"):
                st.error(t("Please complete the root cause statement.", "請完成根本原因陳述。"))
            else:
                st.session_state.step = 6
                st.rerun()

elif st.session_state.step == 6:
    st.subheader(f"Step 6 / 6 — {t('CAPA & Risk Assessment', 'CAPA 與風險評估')}")
    capa = st.session_state.capa
    risk = st.session_state.risk

    if st.session_state.ai_capa_proposal:
        with st.expander(t("💡 View AI CAPA Proposal from Step 5", "💡 查看第 5 步的 AI CAPA 建議"), expanded=True):
            st.markdown(st.session_state.ai_capa_proposal)

    st.markdown(f"### {t('Corrective Action (CA) — Fix THIS event', '糾正措施 (CA) — 解決本次事件')}")
    capa["ca_description"] = st.text_area(t("CA Description", "CA 描述"), value=capa.get("ca_description", ""), height=80)
    col1, col2 = st.columns(2)
    with col1:
        capa["ca_owner"] = st.text_input(t("CA Owner / Dept", "CA 負責人"), value=capa.get("ca_owner", ""))
    with col2:
        capa["ca_due"] = st.text_input(t("CA Due Date (yyyy-Mon-dd)", "CA 完成日期"), value=capa.get("ca_due", ""))
    st.markdown(f"### {t('Preventive Action (PA) — Stop RECURRENCE', '預防措施 (PA) — 防止再次發生')}")
    capa["pa_description"] = st.text_area(t("PA Description (must link to root cause)", "PA 描述（必須連結至根本原因）"), value=capa.get("pa_description", ""), height=80)
    col1, col2 = st.columns(2)
    with col1:
        capa["pa_owner"] = st.text_input(t("PA Owner / Dept", "PA 負責人"), value=capa.get("pa_owner", ""))
    with col2:
        capa["pa_due"] = st.text_input(t("PA Due Date (yyyy-Mon-dd)", "PA 完成日期"), value=capa.get("pa_due", ""))
    capa["effectiveness_check"] = st.text_area(
        t("Effectiveness Check Method", "有效性驗證方法"),
        value=capa.get("effectiveness_check", ""), height=80)
    st.divider()
    st.markdown(f"### {t('Risk Assessment', '風險評估')}")
    risk["hazard"] = st.text_area(t("Identified Hazard / Risk", "已識別的危害"), value=risk.get("hazard", ""), height=60)
    col1, col2 = st.columns(2)
    prob_options = ["Frequent (5)", "Probable (4)", "Occasional (3)", "Remote (2)", "Unlikely (1)"]
    sev_options = ["Catastrophic (4)", "Critical (3)", "Marginal (2)", "Negligible (1)"]
    with col1:
        risk["probability"] = st.selectbox(t("Probability", "發生概率"), prob_options,
            index=prob_options.index(risk.get("probability", "Unlikely (1)")))
    with col2:
        risk["severity"] = st.selectbox(t("Severity", "嚴重程度"), sev_options,
            index=sev_options.index(risk.get("severity", "Negligible (1)")))
    risk["risk_justification"] = st.text_area(
        t("Risk Justification (justify BOTH probability AND severity)", "風險理由"),
        value=risk.get("risk_justification", ""), height=100)
    if st.button(t("🤖 AI: Challenge My Risk Score", "🤖 AI：挑戰我的風險評分"), key="challenge_risk"):
        text = f"Hazard: {risk.get('hazard','')}\nProbability: {risk.get('probability','')}\nSeverity: {risk.get('severity','')}\nJustification: {risk.get('risk_justification','')}"
        if text.strip():
            with st.spinner(t("Analysing...", "分析中...")):
                challenge = ask_claude(SYSTEM_RISK, f"Review this risk assessment:\n{text}\nEvent: {st.session_state.basic.get('title','')}")
                st.warning(f"🤖 {challenge}")
    st.session_state.capa = capa
    st.session_state.risk = risk
    col_back, col_gen = st.columns(2)
    with col_back:
        if st.button(t("← Back", "← 上一步"), use_container_width=True):
            st.session_state.step = 5
            st.rerun()
    with col_gen:
        if st.button(t("📄 Generate Report", "📄 產生報告"), type="primary", use_container_width=True):
            missing = []
            if not capa.get("ca_description"): missing.append("CA")
            if not capa.get("pa_description"): missing.append("PA")
            if not capa.get("ca_due"): missing.append("CA Due Date")
            if not capa.get("pa_due"): missing.append("PA Due Date")
            if not capa.get("effectiveness_check"): missing.append("Effectiveness Check")
            if not risk.get("risk_justification"): missing.append("Risk Justification")
            if missing:
                st.error(t(f"Please complete: {', '.join(missing)}", f"請完成：{', '.join(missing)}"))
            else:
                st.session_state.report_generated = True
                st.rerun()

if st.session_state.report_generated:
    b = st.session_state.basic
    e = st.session_state.event
    imp = st.session_state.impact
    rca = st.session_state.rca
    capa = st.session_state.capa
    risk = st.session_state.risk
    st.divider()
    st.success(t("✅ Report Complete — Preview below", "✅ 報告完成 — 預覽如下"))
    with st.expander(t("📄 Report Preview", "📄 報告預覽"), expanded=True):
        st.markdown(f"## {st.session_state.doc_type}: {b.get('dev_number', '')}")
        st.markdown(f"**{t('Title','標題')}:** {b.get('title','')}")
        st.markdown(f"**{t('Product','產品')}:** {b.get('product','')} | **Lot:** {b.get('lot_number','')} | **Site:** {b.get('site','')}")
        st.markdown(f"**{t('Source Document','來源文件')}:** {b.get('source_doc','')}")
        st.divider()
        st.markdown(f"### {t('Description','偏差描述')}")
        ai_desc = e.get("description_ai", "")
        st.write(ai_desc if ai_desc else e.get("description",""))
        st.markdown(f"**{t('Immediate Actions','立即措施')}:** {e.get('immediate_action','')}")
        st.divider()
        st.markdown(f"### {t('Impact Assessment','影響評估')}")
        st.markdown(f"**Process:** {imp.get('process_impact','')}")
        st.markdown(f"**Quality:** {imp.get('quality_impact','')}")
        st.markdown(f"**Patient Safety:** {imp.get('patient_impact','')}")
        st.divider()
        st.markdown(f"### {t('Root Cause','根本原因')}")
        for m_en, m_zh in zip(["Man","Method","Machine","Materials","Mother Nature","Measurement"],
                               ["人員","方法","機器","物料","環境","測量"]):
            st.markdown(f"**{m_en}:** {rca.get(f'm_{m_en}_status','')} — {rca.get(f'm_{m_en}_detail','')}")
        st.markdown(f"**Root Cause:** {rca.get('root_cause_statement','')}")
        st.divider()
        st.markdown(f"### CAPA")
        st.markdown(f"**CA:** {capa.get('ca_description','')} | Owner: {capa.get('ca_owner','')} | Due: {capa.get('ca_due','')}")
        st.markdown(f"**PA:** {capa.get('pa_description','')} | Owner: {capa.get('pa_owner','')} | Due: {capa.get('pa_due','')}")
        st.markdown(f"**Effectiveness Check:** {capa.get('effectiveness_check','')}")
        st.divider()
        st.markdown(f"### {t('Risk Assessment','風險評估')}")
        st.markdown(f"**Probability:** {risk.get('probability','')} | **Severity:** {risk.get('severity','')}")
        st.markdown(f"**Justification:** {risk.get('risk_justification','')}")

    def generate_word():
        doc = Document()
        doc.add_heading("Investigation, Corrective Action, Preventive Action Form", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Document #: {b.get('dev_number','')}  |  Site: {b.get('site','')}  |  Business Area: {b.get('business_area','')}")
        doc.add_paragraph(f"Date of Occurrence: {b.get('date_occurrence','')}  |  Date of Discovery: {b.get('date_discovery','')}")
        doc.add_paragraph(f"Type of Event: {b.get('event_type','')}  |  Person Involved: {b.get('operator','')}")
        doc.add_paragraph(f"Source Documentation: {b.get('source_doc','')}")
        doc.add_heading("Title", 2); doc.add_paragraph(b.get("title",""))
        doc.add_heading("Description of the Deviation", 2)
        ai_desc = e.get("description_ai", "")
        doc.add_paragraph(ai_desc if ai_desc else e.get("description",""))
        doc.add_paragraph(f"Immediate Actions: {e.get('immediate_action','')}")
        doc.add_paragraph(f"Scope: {e.get('extent','')}")
        doc.add_heading("Impact Assessment", 2)
        doc.add_paragraph(f"Process Impact: {imp.get('process_impact','')}")
        doc.add_paragraph(f"Product Quality Impact: {imp.get('quality_impact','')}")
        doc.add_paragraph(f"Patient Safety Impact: {imp.get('patient_impact','')}")
        doc.add_heading("Root Cause Analysis — 6M", 2)
        for m_en in ["Man","Method","Machine","Materials","Mother Nature","Measurement"]:
            doc.add_paragraph(f"{m_en}: {rca.get(f'm_{m_en}_status','')} — {rca.get(f'm_{m_en}_detail','')}")
        doc.add_paragraph(f"Root Cause: {rca.get('root_cause_statement','')}")
        doc.add_heading("CAPA", 2)
        tbl = doc.add_table(rows=3, cols=3); tbl.style = "Table Grid"
        tbl.rows[0].cells[0].text="Type"; tbl.rows[0].cells[1].text="Description"; tbl.rows[0].cells[2].text="Owner / Due Date"
        tbl.rows[1].cells[0].text="Corrective Action"; tbl.rows[1].cells[1].text=capa.get("ca_description",""); tbl.rows[1].cells[2].text=f"{capa.get('ca_owner','')} / {capa.get('ca_due','')}"
        tbl.rows[2].cells[0].text="Preventive Action"; tbl.rows[2].cells[1].text=capa.get("pa_description",""); tbl.rows[2].cells[2].text=f"{capa.get('pa_owner','')} / {capa.get('pa_due','')}"
        doc.add_paragraph(f"Effectiveness Check: {capa.get('effectiveness_check','')}")
        doc.add_heading("Risk Assessment", 2)
        rt = doc.add_table(rows=2, cols=4); rt.style = "Table Grid"
        rt.rows[0].cells[0].text="Hazard"; rt.rows[0].cells[1].text="Probability"; rt.rows[0].cells[2].text="Severity"; rt.rows[0].cells[3].text="Risk Level"
        prob_val = risk.get("probability","Unlikely (1)")[-2]; sev_val = risk.get("severity","Negligible (1)")[-2]
        try:
            score = int(prob_val) * int(sev_val)
            level = "High" if score >= 12 else ("Medium" if score >= 6 else "Low")
        except:
            score = "N/A"; level = "N/A"
        rt.rows[1].cells[0].text=risk.get("hazard",""); rt.rows[1].cells[1].text=risk.get("probability",""); rt.rows[1].cells[2].text=risk.get("severity",""); rt.rows[1].cells[3].text=f"{level} ({score})"
        doc.add_paragraph(f"Justification: {risk.get('risk_justification','')}")
        doc.add_heading("Approval", 2)
        at = doc.add_table(rows=4, cols=2); at.style = "Table Grid"
        at.rows[0].cells[0].text="Function"; at.rows[0].cells[1].text="Signature / Date"
        for i, role in enumerate(["Author/Initiator","Department Approval","Quality Assurance Approval"]):
            at.rows[i+1].cells[0].text=role

        if st.session_state.chat_history:
            doc.add_page_break()
            doc.add_heading("Appendix A — AI-Assisted RCA Conversation Log", 1)
            doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%b-%d %H:%M')} | Tool: GMP DocWriter XI")
            doc.add_paragraph("This appendix documents the iterative root cause analysis coaching session.")
            for role, msg in st.session_state.chat_history:
                p = doc.add_paragraph()
                if role == "user":
                    run = p.add_run(f"[Investigator]: {msg}")
                    run.bold = True
                else:
                    run = p.add_run(f"[AI RCA Coach]: {msg}")
                    run.italic = True
                p.paragraph_format.space_after = Pt(6)

        if st.session_state.ai_capa_proposal:
            doc.add_page_break()
            doc.add_heading("Appendix B — AI CAPA Proposal", 1)
            doc.add_paragraph(st.session_state.ai_capa_proposal)

        doc.add_paragraph("\nSOP References: SOP-ISOP-023 | SOP-ISOP-024 | SOP-QA-005")
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    word_file = generate_word()
    filename = f"{b.get('dev_number','report')}_{b.get('title','')[:30].replace(' ','-')}.docx"
    st.download_button(
        label=t("⬇️ Download Word Report (.docx)", "⬇️ 下載 Word 報告 (.docx)"),
        data=word_file, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", use_container_width=True)
