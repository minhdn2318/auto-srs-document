import streamlit as st
import os
import io
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from crewai import Agent, Task, Crew, Process, LLM

# ===== 1. QUẢN LÝ TRẠNG THÁI TỪNG CHƯƠNG =====
if 'srs_chapters' not in st.session_state: st.session_state.srs_chapters = {}
if 'test_cases' not in st.session_state: st.session_state.test_cases = ""
if 'api_key' not in st.session_state: st.session_state.api_key = ""

GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "openai/gpt-oss-120b", "openai/gpt-oss-20b"]

# ĐỊNH NGHĨA CÁC CHƯƠNG BẮT BUỘC (CHIA ĐỂ TRỊ)
CHAPTERS_PLAN = [
    {
        "id": "CH1",
        "title": "CHƯƠNG 1: TỔNG QUAN & PHÂN QUYỀN",
        "desc": "Viết Tổng quan dự án, Mục tiêu, Phạm vi. Xác định rõ các Tác nhân (Actors) và Quyền hạn tương ứng."
    },
    {
        "id": "CH2",
        "title": "CHƯƠNG 2: ĐẶC TẢ USE CASE (CHỨC NĂNG CỐT LÕI)",
        "desc": "Mổ xẻ 3-5 chức năng xương sống nhất. BẮT BUỘC format: Tên UC, Tiền điều kiện, Luồng chính (từng bước), Luồng ngoại lệ, Hậu điều kiện, Business Rules."
    },
    {
        "id": "CH3",
        "title": "CHƯƠNG 3: YÊU CẦU DỮ LIỆU & PHI CHỨC NĂNG",
        "desc": "Xác định các Thực thể dữ liệu chính. Yêu cầu Hiệu năng, Bảo mật, Mở rộng."
    }
]

# ===== 2. HÀM HỖ TRỢ DOCX =====
def set_ieee_format(doc, m_left, m_right, m_top, m_bottom):
    section = doc.sections[0]
    section.left_margin = Cm(m_left); section.right_margin = Cm(m_right)
    section.top_margin = Cm(m_top); section.bottom_margin = Cm(m_bottom)

def add_toc(doc):
    doc.add_page_break()
    p = doc.add_paragraph("MỤC LỤC TÀI LIỆU", style='Heading 1'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = doc.add_paragraph().add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin'); run._r.append(fldChar)
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate'); run._r.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end'); run._r.append(fldChar3)
    doc.add_page_break()

# ===== 3. GIAO DIỆN CHÍNH =====
st.set_page_config(page_title="Deep Expert SRS", layout="wide")
st.title("🧠 Deep Expert: Lập trình viên tài liệu (Chapter-by-Chapter)")

with st.sidebar:
    st.header("🔑 Cấu hình Hệ thống")
    st.session_state.api_key = st.text_input("Groq API Key", value=st.secrets.get("GROQ_API_KEY", ""), type="password")
    
    st.divider()
    st.subheader("⚙️ Cấu hình Model")
    m1 = st.selectbox("BA Model (Writer)", GROQ_MODELS, index=0)
    m2 = st.selectbox("QA Model (Critic)", GROQ_MODELS, index=2)
    m3 = st.selectbox("TPM Model (Approver)", GROQ_MODELS, index=0)

user_idea = st.text_area("🚀 Ý tưởng nghiệp vụ (Chi tiết):", height=120)

col1, col2 = st.columns(2)

# ===== 4. VÒNG LẶP GEN SRS (TỪNG CHƯƠNG) =====
if col1.button("🔥 BẮT ĐẦU PHÂN TÍCH TỪNG CHƯƠNG", type="primary"):
    if not st.session_state.api_key or not user_idea: 
        st.error("Thiếu thông tin API Key hoặc Yêu cầu!"); st.stop()
    
    os.environ["GROQ_API_KEY"] = st.session_state.api_key
    st.session_state.srs_chapters = {} # Reset
    st.session_state.test_cases = ""

    # Khởi tạo Não
    ba_llm = LLM(model=f"groq/{m1}", temperature=0.3)
    qa_llm = LLM(model=f"groq/{m2}", temperature=0.1)
    tpm_llm = LLM(model=f"groq/{m3}", temperature=0.1)

    ba_agent = Agent(role="Lead BA", goal="Viết bản nháp siêu chi tiết cho TỪNG CHƯƠNG được giao.", backstory="Bạn là BA 15 năm kinh nghiệm.", llm=ba_llm)
    qa_agent = Agent(role="Principal QA", goal="Phản biện gắt gao bản nháp của BA, tìm ra lỗ hổng logic, case ngoại lệ.", backstory="Bạn là nỗi ám ảnh của BA, chuyên bắt lỗi.", llm=qa_llm)
    tpm_agent = Agent(role="Tech Product Manager", goal="Tổng hợp nháp của BA và phản biện của QA thành bản chốt cuối cùng.", backstory="Bạn là người chốt hạ tài liệu, chuyên nghiệp, xúc tích.", llm=tpm_llm)

    context_memory = "" # Biến lưu trữ các chương trước để làm ngữ cảnh

    for chap in CHAPTERS_PLAN:
        with st.container():
            st.markdown(f"### ⚙️ Đang xử lý: {chap['title']}")
            
            # Cấu hình Task cho từng chương
            t_draft = Task(
                description=f"Ý tưởng gốc: {user_idea}\n\nCác chương đã chốt trước đó (Ngữ cảnh): {context_memory}\n\nNHIỆM VỤ: Hãy viết bản nháp cho: {chap['title']}. Nội dung yêu cầu: {chap['desc']}",
                expected_output="Bản nháp sâu sắc, chi tiết.",
                agent=ba_agent
            )
            t_review = Task(
                description=f"Hãy đọc bản nháp của BA vừa tạo cho {chap['title']}. Chỉ trích và vạch ra các thiếu sót (Edge cases, validation, logic lỗi).",
                expected_output="Danh sách các điểm cần sửa/bổ sung.",
                agent=qa_agent
            )
            t_final = Task(
                description=f"Dựa vào bản nháp và phản biện, hãy viết lại {chap['title']} hoàn chỉnh nhất. Xóa bỏ các râu ria, chỉ giữ lại nội dung tài liệu chuẩn.",
                expected_output="Bản chốt của chương.",
                agent=tpm_agent
            )

            crew = Crew(agents=[ba_agent, qa_agent, tpm_agent], tasks=[t_draft, t_review, t_final], process=Process.sequential)
            
            with st.spinner(f"Agents đang tranh luận về {chap['title']}..."):
                crew.kickoff()
                
                # Lấy dữ liệu thô (raw output) từ từng Task để làm Log
                draft_log = getattr(t_draft.output, 'raw', str(t_draft.output))
                review_log = getattr(t_review.output, 'raw', str(t_review.output))
                final_text = getattr(t_final.output, 'raw', str(t_final.output))

                # Ghi nhận vào state và context
                st.session_state.srs_chapters[chap['id']] = f"{chap['title']}\n{final_text}"
                context_memory += f"\n\n--- TÓM TẮT {chap['title']} ---\n{final_text[:500]}..." # Dùng 500 ký tự đầu làm ngữ cảnh để tránh lố token

            # Hiển thị UI Log minh bạch cho người dùng xem
            with st.expander(f"👁️ Xem Log Tranh luận: {chap['title']}", expanded=False):
                st.markdown("#### 🧑‍💻 Bản nháp của Lead BA")
                st.info(draft_log)
                st.markdown("#### 🕵️ Lời phản biện của Principal QA")
                st.warning(review_log)
                st.markdown("#### 👨‍⚖️ Bản chốt của TPM")
                st.success(final_text)

    st.success("🎉 Đã hoàn thành toàn bộ SRS!")

# ===== 5. MODULE TEST CASE ĐỘC LẬP =====
if st.session_state.srs_chapters:
    if col2.button("🧪 TẠO TEST CASE TỪ SRS NÀY"):
        os.environ["GROQ_API_KEY"] = st.session_state.api_key
        qa_llm = LLM(model=f"groq/{m2}", temperature=0.2)
        
        test_agent = Agent(
            role="Senior Automation QA", 
            goal="Dựa trên tài liệu SRS đã chốt, thiết kế bộ Test Suite chi tiết đến từng bước click chuột.", 
            backstory="Bạn là trùm test hệ thống, tư duy logic tuyệt đỉnh.", 
            llm=qa_llm
        )
        
        full_srs_text = "\n".join(st.session_state.srs_chapters.values())
        
        test_task = Task(
            description=f"Đọc kỹ bản SRS sau:\n{full_srs_text}\n\nHãy tạo Test Case cho các Use Case trong CHƯƠNG 2. Trình bày theo dạng: Mã TC, Tên, Các bước (Steps), Input Data, Kết quả mong đợi (Expected). Bắt buộc phải có cả Positive và Negative case.",
            expected_output="Danh sách Test Case chuyên nghiệp.",
            agent=test_agent
        )
        
        with st.status("🔬 QA đang nhâm nhi cafe và thiết kế Test Case...", expanded=True):
            test_crew = Crew(agents=[test_agent], tasks=[test_task])
            result = test_crew.kickoff()
            st.session_state.test_cases = getattr(result, 'raw', str(result))
            st.success("Đã sinh Test Case xong!")

# ===== 6. HIỂN THỊ TỔNG HỢP & XUẤT FILE =====
if st.session_state.srs_chapters:
    st.divider()
    tab_srs, tab_tc = st.tabs(["📄 TOÀN BỘ TÀI LIỆU SRS", "🧪 BỘ TEST CASE"])
    
    full_doc = "\n\n".join(st.session_state.srs_chapters.values())
    
    with tab_srs:
        st.text_area("SRS Master Document", full_doc, height=500)
    
    with tab_tc:
        if st.session_state.test_cases:
            st.text_area("Test Suite", st.session_state.test_cases, height=500)
        else:
            st.info("Hãy bấm nút 'Tạo Test Case' ở trên để sinh kịch bản kiểm thử.")

    def create_export_docx():
        doc = Document()
        set_ieee_format(doc, 3.0, 2.0, 2.0, 2.0)
        
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"\n\n\nENTERPRISE SRS & TEST PLAN\n\n")
        run.bold = True; run.font.size = Pt(20)
        add_toc(doc)
        
        # Combine SRS and TC
        final_text = full_doc
        if st.session_state.test_cases:
            final_text += "\n\nCHƯƠNG 4: TEST SUITE (KỊCH BẢN KIỂM THỬ)\n" + st.session_state.test_cases

        for line in final_text.split('\n'):
            line = line.strip()
            if not line: continue
            if re.match(r'^CHƯƠNG \d+', line):
                h = doc.add_heading(line, level=1)
                for r in h.runs: r.font.color.rgb = RGBColor(0,0,0)
            else:
                p = doc.add_paragraph(line)
                if p.runs: p.runs[0].font.name = "Times New Roman"
        
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    st.download_button("📥 TẢI XUỐNG FILE WORD TỔNG HỢP", create_export_docx(), "Enterprise_Project_Plan.docx", type="primary", use_container_width=True)