import streamlit as st
import os
import io
import re
import time
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from crewai import Agent, Task, Crew, Process, LLM

# ===== 1. QUẢN LÝ TRẠNG THÁI =====
if 'final_srs' not in st.session_state: st.session_state.final_srs = ""
if 'final_testcases' not in st.session_state: st.session_state.final_testcases = ""
if 'is_running' not in st.session_state: st.session_state.is_running = False

GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "openai/gpt-oss-120b", "openai/gpt-oss-20b"]

# ===== 2. GIAO DIỆN CHÍNH =====
st.set_page_config(page_title="Autonomous BA Agent", layout="wide")
st.title("🤖 AUTONOMOUS BA: Cỗ Máy Sinh Tài Liệu Tự Trị")
st.markdown("*Nhập ý tưởng 1 lần - Hệ thống tự lập dàn ý, tự bóc tách Use Case, tự tranh luận và sinh Test Case.*")

with st.sidebar:
    st.header("🔑 Cấu hình Hệ thống")
    st.session_state.api_key = st.text_input("Groq API Key", value=st.secrets.get("GROQ_API_KEY", ""), type="password")
    
    st.divider()
    st.subheader("🧠 Cấu hình Não Bộ (LLM)")
    m_ba = st.selectbox("Lead BA (Viết chính)", GROQ_MODELS, index=0)
    m_qa = st.selectbox("Principal QA (Bắt lỗi)", GROQ_MODELS, index=2)
    m_tpm = st.selectbox("Tech Lead (Chốt sổ)", GROQ_MODELS, index=0)

user_idea = st.text_area("🚀 Nhập ý tưởng/Yêu cầu nghiệp vụ (Càng chi tiết càng tốt):", height=100)
auto_testcase = st.checkbox("🧪 Tự động sinh Test Case sau khi chốt SRS", value=True)

# ===== 3. CORE LOGIC: AUTONOMOUS WORKFLOW =====
if st.button("🔥 KHỞI ĐỘNG CỖ MÁY (1-CLICK AUTO)", type="primary"):
    if not st.session_state.api_key or not user_idea: 
        st.error("Bác nhập thiếu Key hoặc Ý tưởng rồi!"); st.stop()
    
    os.environ["GROQ_API_KEY"] = st.session_state.api_key
    st.session_state.final_srs = ""
    st.session_state.final_testcases = ""

    # Khởi tạo 3 bộ LLM
    llm_ba = LLM(model=f"groq/{m_ba}", temperature=0.3)
    llm_qa = LLM(model=f"groq/{m_qa}", temperature=0.1)
    llm_tpm = LLM(model=f"groq/{m_tpm}", temperature=0.1)

    # Khởi tạo Agents (Chỉ định nghĩa 1 lần để tái sử dụng)
    agent_ba = Agent(role="Lead BA", goal="Phân tích hệ thống và viết đặc tả cực kỳ sâu sắc.", backstory="Chuyên gia BA 15 năm kinh nghiệm, luôn tư duy hệ thống.", llm=llm_ba)
    agent_qa = Agent(role="Principal QA", goal="Tìm mọi lỗ hổng logic, case ngoại lệ trong tài liệu.", backstory="Nỗi ác mộng của BA, tư duy 'user luôn làm sai'.", llm=llm_qa)
    agent_tpm = Agent(role="Tech Lead", goal="Tổng hợp, chốt hạ tài liệu chuẩn format hành chính kỹ thuật.", backstory="Người ra quyết định cuối cùng, ghét sự lan man.", llm=llm_tpm)

    with st.status("🚀 Cỗ máy Autonomous đang vận hành...", expanded=True) as status:
        full_document = ""
        
        # --- BƯỚC 1: LẬP OUTLINE VÀ TỔNG QUAN (CHƯƠNG 1) ---
        st.write("### ⚙️ BƯỚC 1: Đang lập Outline & Tổng quan dự án...")
        t_ch1_draft = Task(description=f"Dựa vào ý tưởng: '{user_idea}'. Viết CHƯƠNG 1: TỔNG QUAN. Gồm: Mục tiêu, Phạm vi, Đối tượng sử dụng.", expected_output="Bản nháp Chương 1", agent=agent_ba)
        t_ch1_final = Task(description="Review và chuẩn hóa bản nháp Chương 1. Output không dùng markdown phức tạp.", expected_output="Bản chốt Chương 1", agent=agent_tpm)
        
        crew_ch1 = Crew(agents=[agent_ba, agent_tpm], tasks=[t_ch1_draft, t_ch1_final], verbose=False)
        ch1_result = getattr(crew_ch1.kickoff(), 'raw', str(crew_ch1.kickoff()))
        full_document += "CHƯƠNG 1: TỔNG QUAN DỰ ÁN\n" + ch1_result + "\n\n"
        
        with st.expander("👁️ Xem log Chương 1 (Tổng quan)"): st.write(ch1_result)

        # --- BƯỚC 2: AUTO-EXTRACT USE CASES (BÓC TÁCH TỰ ĐỘNG) ---
        st.write("### ⚙️ BƯỚC 2: AI đang tự động nhận diện các Use Case...")
        t_extract_uc = Task(
            description=f"Dựa vào ý tưởng: '{user_idea}'. Hãy liệt kê TÊN của 3 đến 5 Use Case cốt lõi nhất của hệ thống. CHỈ TRẢ VỀ MỘT DANH SÁCH có gạch đầu dòng, không giải thích gì thêm. VD: \n- Đăng nhập hệ thống\n- Thanh toán giỏ hàng",
            expected_output="Danh sách gạch đầu dòng các Use Case",
            agent=agent_ba
        )
        crew_extract = Crew(agents=[agent_ba], tasks=[t_extract_uc], verbose=False)
        uc_raw_list = getattr(crew_extract.kickoff(), 'raw', str(crew_extract.kickoff()))
        
        # Parse danh sách Use Case bằng Python
        use_cases = [line.replace('-', '').strip() for line in uc_raw_list.split('\n') if '-' in line and len(line) > 5]
        if not use_cases: use_cases = ["Quản lý nghiệp vụ chính", "Xử lý ngoại lệ hệ thống"] # Fallback
        
        st.info(f"🎯 AI đã nhận diện được {len(use_cases)} Use Case cốt lõi: {', '.join(use_cases)}")
        full_document += "CHƯƠNG 2: ĐẶC TẢ USE CASE CHI TIẾT\n"

        # --- BƯỚC 3: AUTONOMOUS LOOP TỪNG USE CASE ---
        st.write("### ⚙️ BƯỚC 3: Bắt đầu vòng lặp tranh luận cho TỪNG Use Case...")
        
        for idx, uc_name in enumerate(use_cases):
            st.write(f"🔄 Đang mổ xẻ UC {idx+1}/{len(use_cases)}: **{uc_name}**")
            
            t_uc_draft = Task(
                description=f"Viết đặc tả siêu chi tiết cho Use Case: '{uc_name}' (thuộc hệ thống: {user_idea}). Bắt buộc format: Tên UC, Tiền điều kiện, Luồng chính (đánh số từng bước), Luồng phụ/Ngoại lệ.",
                expected_output="Bản nháp Use Case", agent=agent_ba
            )
            t_uc_critic = Task(
                description=f"Đọc bản nháp Use Case '{uc_name}'. Tìm ra các lỗ hổng logic, các trường hợp người dùng nhập sai, rớt mạng, lỗi DB mà BA chưa tính tới.",
                expected_output="Danh sách điểm yếu của Use Case", agent=agent_qa
            )
            t_uc_final = Task(
                description=f"Tổng hợp bản nháp của BA và lời chê của QA. Chốt lại thành bản đặc tả Use Case '{uc_name}' hoàn hảo nhất. Ghi rõ các Business Rules.",
                expected_output="Bản chốt Use Case", agent=agent_tpm
            )
            
            crew_uc = Crew(agents=[agent_ba, agent_qa, agent_tpm], tasks=[t_uc_draft, t_uc_critic, t_uc_final], verbose=False)
            uc_result = getattr(crew_uc.kickoff(), 'raw', str(crew_uc.kickoff()))
            full_document += f"\n2.{idx+1}. Use Case: {uc_name}\n" + uc_result + "\n"
            
            with st.expander(f"👁️ Xem log mổ xẻ UC: {uc_name}"):
                st.markdown("#### 🕵️ Lời chỉ trích của QA:")
                st.warning(getattr(t_uc_critic.output, 'raw', "QA đồng ý với BA"))
                st.markdown("#### 👨‍⚖️ Bản chốt của TPM:")
                st.success(uc_result)
                
            time.sleep(1) # Nghỉ 1s tránh hit Rate Limit của Groq

        # --- BƯỚC 4: YÊU CẦU PHI CHỨC NĂNG ---
        st.write("### ⚙️ BƯỚC 4: Viết yêu cầu phi chức năng (Bảo mật, Hiệu năng)...")
        t_ch3_draft = Task(description=f"Viết CHƯƠNG 3: YÊU CẦU PHI CHỨC NĂNG cho dự án: {user_idea}. Tập trung vào Data, Hiệu năng, Bảo mật.", expected_output="Bản nháp", agent=agent_ba)
        t_ch3_final = Task(description="Chuẩn hóa bản nháp Chương 3", expected_output="Bản chốt", agent=agent_tpm)
        crew_ch3 = Crew(agents=[agent_ba, agent_tpm], tasks=[t_ch3_draft, t_ch3_final], verbose=False)
        ch3_result = getattr(crew_ch3.kickoff(), 'raw', str(crew_ch3.kickoff()))
        full_document += "\nCHƯƠNG 3: YÊU CẦU PHI CHỨC NĂNG & DỮ LIỆU\n" + ch3_result
        
        st.session_state.final_srs = full_document

        # --- BƯỚC 5: AUTO TEST CASE (Nếu chọn) ---
        if auto_testcase:
            st.write("### ⚙️ BƯỚC 5: Tự động đẻ Test Case từ SRS vừa chốt...")
            t_testcase = Task(
                description=f"Đọc toàn bộ SRS sau:\n{full_document[:3000]}...\n\nHãy tạo BẢNG TEST CASE CHI TIẾT (ID, Tên, Steps, Input, Expected) cho các Use Case ở Chương 2. Bao phủ cả Positive và Negative.",
                expected_output="Danh sách Test Case hoàn chỉnh", agent=agent_qa
            )
            crew_tc = Crew(agents=[agent_qa], tasks=[t_testcase], verbose=False)
            st.session_state.final_testcases = getattr(crew_tc.kickoff(), 'raw', str(crew_tc.kickoff()))

        status.update(label="✅ TOÀN BỘ QUY TRÌNH AUTONOMOUS ĐÃ HOÀN TẤT!", state="complete")

# ===== 4. HIỂN THỊ VÀ XUẤT DOCX =====
if st.session_state.final_srs:
    st.divider()
    tab1, tab2 = st.tabs(["📄 TÀI LIỆU SRS (MASTER)", "🧪 BỘ TEST CASE (QA)"])
    
    with tab1: st.text_area("Nội dung SRS", st.session_state.final_srs, height=500)
    with tab2: 
        if st.session_state.final_testcases: st.text_area("Nội dung Test Case", st.session_state.final_testcases, height=500)
        else: st.info("Không có Test Case (Do chưa tích chọn lúc khởi chạy).")

    def create_export_docx():
        doc = Document()
        doc.sections[0].left_margin = Cm(3.0); doc.sections[0].right_margin = Cm(2.0)
        
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"\n\n\nAUTONOMOUS SRS & TEST PLAN\n\n")
        run.bold = True; run.font.size = Pt(20)
        
        final_text = st.session_state.final_srs
        if st.session_state.final_testcases:
            final_text += "\n\nCHƯƠNG 4: KỊCH BẢN KIỂM THỬ (TEST SUITE)\n" + st.session_state.final_testcases

        for line in final_text.split('\n'):
            line = line.strip()
            if not line: continue
            if re.match(r'^CHƯƠNG \d+', line) or re.match(r'^2\.\d+\. Use Case', line):
                h = doc.add_heading(line, level=1 if 'CHƯƠNG' in line else 2)
                for r in h.runs: r.font.color.rgb = RGBColor(0,0,0)
            else:
                p = doc.add_paragraph(line)
                if p.runs: p.runs[0].font.name = "Times New Roman"
        
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    st.download_button("📥 TẢI XUỐNG FILE WORD TỔNG HỢP (SRS + TESTCASE)", create_export_docx(), "Auto_SRS_Master.docx", type="primary", use_container_width=True)