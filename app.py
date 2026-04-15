import streamlit as st
import os
import io
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Thay đổi lớn nhất: Dùng LLM từ lõi của CrewAI
from crewai import Agent, Task, Crew, Process, LLM

# ===== 1. QUẢN LÝ TRẠNG THÁI =====
if 'srs_content' not in st.session_state: st.session_state.srs_content = None
if 'api_key' not in st.session_state: st.session_state.api_key = ""

# Danh sách Model Text của Groq (Lọc bỏ Whisper)
GROQ_MODELS = [
    "llama-3.3-70b-versatile", 
    "llama-3.1-8b-instant", 
    "openai/gpt-oss-120b", 
    "openai/gpt-oss-20b"
]

# ===== 2. ĐỊNH DẠNG DOCX =====
def set_ieee_format(doc, m_left, m_right, m_top, m_bottom):
    section = doc.sections[0]
    section.left_margin = Cm(m_left); section.right_margin = Cm(m_right)
    section.top_margin = Cm(m_top); section.bottom_margin = Cm(m_bottom)
    section.different_first_page_header_footer = True
    
def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instr = OxmlElement('w:instrText'); instr.text = "PAGE"; run._r.append(instr)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar2)

def add_toc(doc):
    doc.add_page_break()
    p = doc.add_paragraph("MỤC LỤC", style='Heading 1'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph = doc.add_paragraph(); run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate'); run._r.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar3)
    doc.add_page_break()

# ===== 3. GIAO DIỆN STREAMLIT =====
st.set_page_config(page_title="Master SRS CrewAI", layout="wide")
st.title("🤖 Multi-Model SRS Architect (Groq Powered)")

with st.sidebar:
    st.header("⚙️ Cấu hình Hệ thống")
    # Đọc key từ secrets (nếu có cấu hình trên Streamlit Cloud), nếu không thì để trống cho người dùng nhập
    default_key = st.secrets.get("GROQ_API_KEY", "") if hasattr(st, "secrets") else ""
    st.session_state.api_key = st.text_input("Groq API Key", value=default_key, type="password")
    
    st.divider()
    st.subheader("📏 Page Setup (cm)")
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        m_left = st.number_input("Trái", value=3.0)
        m_top = st.number_input("Trên", value=2.0)
    with c_m2:
        m_right = st.number_input("Phải", value=2.0)
        m_bottom = st.number_input("Dưới", value=2.0)
    
    line_sp = st.number_input("Giãn dòng", value=1.15)
    font_sz = st.number_input("Cỡ chữ", value=12)

# --- GIAO DIỆN PHÂN VAI & CHỌN MODEL ---
st.subheader("👥 Cấu hình Đội ngũ AI & Bộ não")
with st.expander("Tùy chỉnh vai trò và Model cho từng Agent", expanded=True):
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("### ✍️ Người viết (Writer)")
        model_1 = st.selectbox("🧠 Não của Writer", GROQ_MODELS, index=1)
        role_1 = st.text_input("Role 1", "Senior Solution Architect")
        goal_1 = st.text_area("Goal 1", "Lập bản nháp SRS chuẩn IEEE 830 với đầy đủ 5 chương cốt lõi dựa trên ý tưởng.", height=100)
        backstory_1 = st.text_area("Backstory 1", "Bạn có 30 năm kinh nghiệm phân tích hệ thống, viết rất nhanh và chuẩn format.", height=80)
    
    with c2:
        st.markdown("### 🕵️ Người phản biện (Reviewer)")
        model_2 = st.selectbox("🧠 Não của Reviewer", GROQ_MODELS, index=2)
        role_2 = st.text_input("Role 2", "Cybersecurity & QA Lead")
        goal_2 = st.text_area("Goal 2", "Phản biện gắt gao bản nháp SRS về tính bảo mật, khả năng mở rộng và các luồng lỗi.", height=100)
        backstory_2 = st.text_area("Backstory 2", "Bạn là chuyên gia tìm lỗ hổng logic. Không bao giờ hài lòng với bản nháp đầu tiên.", height=80)
    
    with c3:
        st.markdown("### 👨‍⚖️ Người chốt duyệt (Approver)")
        model_3 = st.selectbox("🧠 Não của Approver", GROQ_MODELS, index=0)
        role_3 = st.text_input("Role 3", "Product Master")
        goal_3 = st.text_area("Goal 3", "Tổng hợp bản nháp và các phản biện để xuất ra bản SRS hoàn chỉnh, chi tiết và chuyên nghiệp nhất.", height=100)
        backstory_3 = st.text_area("Backstory 3", "Bạn là Giám đốc Sản phẩm. Đảm bảo tài liệu sẵn sàng giao cho team Dev mà không bị hoa mỹ markdown.", height=80)

st.divider()
user_idea = st.text_area("🚀 Ý tưởng phần mềm / Yêu cầu nghiệp vụ:", height=100)

# ===== 4. FLOW GEN SRS VỚI MULTI-MODEL =====
if st.button("📝 Bắt đầu Khởi tạo SRS", type="primary"):
    if not st.session_state.api_key or not user_idea: 
        st.error("Bác quên nhập API Key hoặc Ý tưởng rồi kìa!"); st.stop()
    
    # Bơm key vào hệ thống cho CrewAI tự nhận
    os.environ["GROQ_API_KEY"] = st.session_state.api_key

    with st.status("🚀 Đội ngũ đang họp... (Xem log chi tiết ở Terminal)", expanded=True) as status:
        st.write("Đang cấp phát 'Bộ não' (LLM) riêng biệt cho từng Agent...")
        
        # Khởi tạo 3 bộ LLM chuẩn Native của CrewAI (Đã fix lỗi LiteLLM)
        llm_writer = LLM(model=f"groq/{model_1}", temperature=0.3)
        llm_reviewer = LLM(model=f"groq/{model_2}", temperature=0.1) 
        llm_approver = LLM(model=f"groq/{model_3}", temperature=0.2)

        st.write("Đang thức tỉnh Agents...")
        agent_writer = Agent(role=role_1, goal=goal_1, backstory=backstory_1, llm=llm_writer, verbose=True)
        agent_reviewer = Agent(role=role_2, goal=goal_2, backstory=backstory_2, llm=llm_reviewer, verbose=True)
        agent_approver = Agent(role=role_3, goal=goal_3, backstory=backstory_3, llm=llm_approver, verbose=True)

        st.write("Đang giao Task...")
        task_draft = Task(
            description=f"Viết bản nháp SRS cho ý tưởng sau: {user_idea}",
            expected_output="Bản nháp SRS chuẩn IEEE 830 đầy đủ các phần.",
            agent=agent_writer
        )
        task_review = Task(
            description="Đọc bản nháp SRS vừa tạo. Chỉ ra các thiếu sót về logic, bảo mật. Đưa ra đề xuất sửa chữa cụ thể.",
            expected_output="Danh sách các điểm hổng logic cần sửa chữa.",
            agent=agent_reviewer
        )
        task_finalize = Task(
            description="Dùng bản nháp ban đầu và nhận xét của Reviewer để viết lại bản SRS cuối cùng. Trình bày chuẩn format văn bản hành chính, KHÔNG dùng Markdown (** hay #).",
            expected_output="Bản tài liệu SRS hoàn chỉnh cuối cùng.",
            agent=agent_approver
        )

        srs_crew = Crew(
            agents=[agent_writer, agent_reviewer, agent_approver],
            tasks=[task_draft, task_review, task_finalize],
            process=Process.sequential,
            verbose=True
        )

        st.write(f"🔥 Bắt đầu suy luận luân phiên ({model_1} -> {model_2} -> {model_3})...")
        try:
            result = srs_crew.kickoff()
            st.session_state.srs_content = str(result)
            status.update(label="✅ Quá trình suy luận Multi-Model đã hoàn tất!", state="complete")
        except Exception as e:
            st.error(f"Lỗi hệ thống: {str(e)}")
            status.update(label="❌ Có lỗi xảy ra!", state="error")
            st.stop()

# ===== 5. XUẤT FILE DOCX =====
if st.session_state.srs_content:
    st.divider()
    st.subheader("📄 Kết quả SRS")
    st.text_area("Bản thảo SRS Cuối cùng (Plain Text Preview)", st.session_state.srs_content, height=400)
    
    def create_srs_docx():
        doc = Document()
        set_ieee_format(doc, m_left, m_right, m_top, m_bottom)
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"\n\n\nSOFTWARE REQUIREMENTS SPECIFICATION\n(IEEE 830 Standard)\n\nProject: {user_idea[:50].upper()}...")
        r.bold = True; r.font.size = Pt(20); r.font.name = "Times New Roman"
        add_toc(doc)
        
        for line in st.session_state.srs_content.split('\n'):
            if line.strip():
                if re.match(r'^\d+\.', line) or re.match(r'^\d+\.\d+', line):
                    level = line.count('.') if line.count('.') <= 3 else 3
                    h = doc.add_heading(line, level=level)
                    for run in h.runs: 
                        run.font.name = "Times New Roman"; run.font.color.rgb = RGBColor(0,0,0)
                else:
                    p = doc.add_paragraph(line); p.paragraph_format.line_spacing = line_sp
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    run = p.runs[0] if p.runs else p.add_run(line)
                    run.font.name = "Times New Roman"; run.font.size = Pt(font_sz)

        add_page_number(doc.sections[0].footer.paragraphs[0])
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    st.download_button("📥 Tải File Word (Chuẩn IEEE 830)", create_srs_docx(), "SRS_MultiModel_Master.docx", type="primary")