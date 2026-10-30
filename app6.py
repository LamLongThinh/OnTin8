import streamlit as st
from docx import Document
from datetime import datetime
import datetime as dt
import re, json, pandas as pd, os, random, time
from io import BytesIO
import base64
import requests

# === THÊM PAGE CONFIG ĐỂ TỐI ƯU HIỂN THỊ ===
st.set_page_config(layout="wide", page_title="MyHoaQuiz - Tin học 8", initial_sidebar_state="expanded") 
# ==========================================

# ====== Cấu hình cơ bản ======
QUIZ_FILE = "questions.json"
ESSAY_FILE = "essays.json"
SCORES_FILE = "scores.xlsx"
ADMIN_PASSWORD = "admin123"
EXPECTED_COLUMNS = ["Tên Học Sinh", "Lớp", "Điểm Trắc Nghiệm", "Điểm Tự Luận", "Tổng Điểm", "Chi Tiết Tự Luận", "Thời Gian Nộp Bài"]
DEFAULT_TIME_LIMIT = 45
LOGO_PATH = "LOGO.png"

# GitHub Repository cấu hình
GITHUB_REPO_API = "https://api.github.com/repos/LamLongThinh/OnTin8/contents"

# Thêm logo và tiêu đề (KHU VỰC CHÍNH)
col1, col2, col3 = st.columns([5, 4, 1])

if os.path.exists(LOGO_PATH):
    with col3:
        st.image(LOGO_PATH, width=100) 
    
with col2:
    st.markdown(
        """
        <h1 style='text-align: center; font-weight: 800;'>   
            MyHoaQUIZ
        </h1>
        <h2 style='text-align: center; font-weight: 800;'>    
            📝 ÔN TẬP TRẮC NGHIỆM & TỰ LUẬN – TIN HỌC 8
        </h2>
        <h6 style='text-align: center; color: gray; font-weight: 700; margin-top: -10px;'> 
            KIẾN THỨC TRỌNG TÂM HỌC KÌ
        </h6>
        """,
        unsafe_allow_html=True
    )

# ====== Khởi tạo file cục bộ ======
def init_local_files():
    if not os.path.exists(SCORES_FILE):
        pd.DataFrame(columns=EXPECTED_COLUMNS).to_excel(SCORES_FILE, index=False)
init_local_files()

# ====== Đồng bộ file từ GitHub ======
def fetch_docx_from_github():
    """Tự động tìm và tải file .docx từ GitHub repository OnTin8"""
    try:
        response = requests.get(GITHUB_REPO_API, timeout=10)
        if response.status_code == 200:
            files = response.json()
            docx_file = next((f for f in files if f['name'].endswith('.docx')), None)
            if docx_file and 'download_url' in docx_file:
                doc_resp = requests.get(docx_file['download_url'])
                if doc_resp.status_code == 200:
                    return BytesIO(doc_resp.content)
    except Exception as e:
        st.sidebar.warning(f"⚠️ Không thể tự động đồng bộ từ GitHub: {e}")
    return None

# ====== Hàm xử lý file Word (40 Trắc nghiệm + 2 Tự luận) ======
def parse_quiz_from_word(file):
    doc = Document(file)
    text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    parts = re.split(r"(Tự\s*luận|Câu\s*hỏi\s*tự\s*luận)", text, flags=re.IGNORECASE)
    
    mcq_text = parts[0]
    essay_text = "".join(parts[1:]) if len(parts) > 1 else ""
    
    # 1. Xử lý Trắc nghiệm (40 câu)
    blocks = re.split(r"(Câu\s*\d+[.:])", mcq_text)
    quiz, content_blocks = [], []
    for i in range(1, len(blocks), 2):
        if i + 1 < len(blocks):
            content_blocks.append(blocks[i] + blocks[i + 1])
            
    for block in content_blocks:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if not lines: continue
        q_text = re.sub(r"^Câu\s*\d+[.:]\s*", "", lines[0]).strip()
        options, correct = {}, None
        for line in lines[1:]:
            if re.match(r"^[A-D][\.\)]", line): 
                letter, content = line.split('.', 1) if '.' in line else line.split(')', 1)
                options[letter.strip().upper()] = content.strip()
            elif re.search(r"đáp\s*án", line, flags=re.IGNORECASE):
                correct = line.split(":")[-1].strip().upper()
        if len(options) >= 4 and correct in options:
            quiz.append({
                "question": q_text, 
                "options": [options[k] for k in ["A","B","C","D"] if k in options], 
                "answer": options[correct], 
                "image_base64": None
            })

    # 2. Xử lý Tự luận (2 câu)
    essay_blocks = re.split(r"((?:Bài|Câu)\s*(?:tự\s*luận)?\s*\d+[.:])", essay_text, flags=re.IGNORECASE)
    essays = []
    essay_contents = []
    for i in range(1, len(essay_blocks), 2):
        if i + 1 < len(essay_blocks):
            essay_contents.append(essay_blocks[i] + essay_blocks[i + 1])
            
    for block in essay_contents:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if not lines: continue
        e_text = re.sub(r"^.*?(?:tự\s*luận)?\s*\d+[.:]\s*", "", lines[0]).strip()
        if not e_text and len(lines) > 1:
            e_text = lines[1]
        essays.append({"question": e_text if e_text else lines[0]})
        
    if not essays:
        essays = [
            {"question": "Câu hỏi tự luận 1 (Giáo viên cập nhật trong nội dung file Word)."},
            {"question": "Câu hỏi tự luận 2 (Giáo viên cập nhật trong nội dung file Word)."}
        ]

    return quiz[:40], essays[:2]

def save_quiz_data(quiz, essays):
    with open(QUIZ_FILE, "w", encoding="utf-8") as f:
        json.dump(quiz, f, ensure_ascii=False, indent=4)
    with open(ESSAY_FILE, "w", encoding="utf-8") as f:
        json.dump(essays, f, ensure_ascii=False, indent=4)
    st.success(f"✅ Đã lưu {len(quiz)} câu trắc nghiệm và {len(essays)} câu tự luận!")

def load_quiz_data():
    quiz, essays = [], []
    if os.path.exists(QUIZ_FILE):
        with open(QUIZ_FILE, "r", encoding="utf-8") as f:
            quiz = json.load(f)
    if os.path.exists(ESSAY_FILE):
        with open(ESSAY_FILE, "r", encoding="utf-8") as f:
            essays = json.load(f)
    return quiz, essays

# ====== Tích hợp Supabase ======
def save_to_supabase(result_data):
    supabase_url = st.session_state.get("supabase_url", os.getenv("SUPABASE_URL", ""))
    supabase_key = st.session_state.get("supabase_key", os.getenv("SUPABASE_KEY", ""))
    
    if not supabase_url or not supabase_key:
        return False, "Chưa cấu hình Supabase URL hoặc Key."
    
    headers = {
        "apikey": supabase_key,
        "Authorization": f"Bearer {supabase_key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
    }
    url = f"{supabase_url.rstrip('/')}/rest/v1/ontin8_results"
    
    try:
        resp = requests.post(url, json=result_data, headers=headers, timeout=10)
        if resp.status_code in [200, 201, 204]:
            return True, "Lưu Supabase thành công."
        else:
            return False, f"Lỗi Supabase ({resp.status_code}): {resp.text}"
    except Exception as e:
        return False, f"Lỗi kết nối Supabase: {e}"

# =========================================================================
# GIAO DIỆN HỌC SINH
# =========================================================================
def student_ui():
    st.header("📚 Khu vực Kiểm tra Trắc nghiệm & Tự luận")
    
    if not os.path.exists(QUIZ_FILE):
        github_file = fetch_docx_from_github()
        if github_file:
            q, e = parse_quiz_from_word(github_file)
            save_quiz_data(q, e)
            
    quiz_raw, essay_raw = load_quiz_data()
    if not quiz_raw:
        st.warning("⚠️ Chưa có dữ liệu đề thi. Vui lòng thông báo giáo viên cập nhật hệ thống.")
        return

    is_submitted = st.session_state.get("quiz_submitted", False)
    doing_quiz = st.session_state.get("doing_quiz", False)

    if not is_submitted and not doing_quiz:
        st.info("Vui lòng điền đầy đủ thông tin để bắt đầu làm bài.")
        with st.form("student_login_form"):
            name = st.text_input("✍️ Họ và Tên học sinh:", key="stu_name_form")
            clas = st.text_input("🏫 Lớp (VD: 8A1):", key="stu_class_form")
            submitted_login = st.form_submit_button("🚀 Bắt đầu làm bài", type="primary")

        if submitted_login:
            st.session_state["stu_name"] = name.strip()
            st.session_state["stu_class"] = clas.strip()
            
            if not st.session_state["stu_name"] or not st.session_state["stu_class"]:
                st.error("⚠️ Vui lòng nhập đầy đủ Họ và Tên cùng Lớp.")
                return

            quiz = [dict(q) for q in quiz_raw]
            random.shuffle(quiz)
            for q in quiz:
                norm_opts = []
                for opt in q["options"]:
                    m = re.match(r"^[A-D][\.\)]\s*(.*)", opt)
                    norm_opts.append(m.group(1).strip() if m else opt.strip())
                opts_with_flag = [(text, text == q["answer"]) for text in norm_opts]
                random.shuffle(opts_with_flag)
                q["options"] = [t for t, _ in opts_with_flag]
                q["answer"] = next((t for t, flag in opts_with_flag if flag), "")
                
            st.session_state["quiz_data"] = quiz
            st.session_state["essay_data"] = essay_raw
            st.session_state["start_time"] = datetime.now()
            st.session_state["doing_quiz"] = True
            st.session_state["responses"] = {q["question"]: None for q in quiz}
            st.session_state["essay_responses"] = {e["question"]: "" for e in essay_raw}
            st.session_state["quiz_submitted"] = False
            st.rerun()
        return

    name = st.session_state.get("stu_name", "")
    clas = st.session_state.get("stu_class", "")

    if doing_quiz:
        st.markdown("---")
        st.subheader(f"👋 Thí sinh: {name} - Lớp {clas}")
        st.info(f"Đề gồm **{len(quiz_raw)} câu trắc nghiệm** (8 điểm) và **{len(essay_raw)} câu tự luận** (2 điểm). Thời gian: {DEFAULT_TIME_LIMIT} phút.")
        
        quiz = st.session_state.get("quiz_data", [])
        essays = st.session_state.get("essay_data", [])

        start_time = st.session_state.get("start_time", datetime.now())
        elapsed = (datetime.now() - start_time).total_seconds()
        remaining = max(DEFAULT_TIME_LIMIT * 60 - int(elapsed), 0)
        mins, secs = divmod(remaining, 60)
        st.progress(min(1.0, elapsed / (DEFAULT_TIME_LIMIT * 60)))
        st.markdown(f"⏳ **Thời gian còn lại: {int(mins):02d}:{int(secs):02d}**")

        auto_submit = False
        if remaining == 0:
            st.warning("⏰ Đã hết thời gian làm bài! Hệ thống tự động nộp.")
            auto_submit = True
            
        with st.form("quiz_form"):
            st.markdown("### I. PHẦN TRẮC NGHIỆM (8.0 điểm)")
            for idx, q in enumerate(quiz, start=1):
                prev_choice = st.session_state["responses"].get(q["question"], None)
                try:
                    default_index = q["options"].index(prev_choice)
                except (ValueError, AttributeError):
                    default_index = None

                choice = st.radio(
                    f"**Câu {idx}:** {q['question']}",
                    q["options"],
                    index=default_index,
                    key=f"q_{idx}_radio"
                )
                st.session_state["responses"][q["question"]] = choice
                st.write("---")
            
            st.markdown("### II. PHẦN TỰ LUẬN (2.0 điểm)")
            for idx, e in enumerate(essays, start=1):
                prev_ans = st.session_state["essay_responses"].get(e["question"], "")
                ans = st.text_area(
                    f"**Câu tự luận {idx}:** {e['question']}",
                    value=prev_ans,
                    key=f"essay_{idx}_text",
                    height=120
                )
                st.session_state["essay_responses"][e["question"]] = ans
                st.markdown("---")
            
            submitted = st.form_submit_button("✅ Nộp bài thi", type="primary")

        if auto_submit or submitted:
            mcq_score = 0
            total_mcq = len(quiz)
            for q in quiz:
                chosen = st.session_state["responses"].get(q["question"], None)
                if chosen and chosen == q["answer"]: 
                    mcq_score += 1
            
            final_mcq_score_10 = round(mcq_score * (8.0 / total_mcq), 2) if total_mcq else 0
            
            st.session_state["mcq_score"] = mcq_score
            st.session_state["final_mcq_score_10"] = final_mcq_score_10
            
            st.balloons() 
            st.toast("🎉 Nộp bài thành công!")
            
            submission_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            essay_summary = json.dumps(st.session_state["essay_responses"], ensure_ascii=False)
            
            result_record = {
                "Tên Học Sinh": name,
                "Lớp": clas,
                "Điểm Trắc Nghiệm": final_mcq_score_10,
                "Điểm Tự Luận": 0.0,
                "Tổng Điểm": final_mcq_score_10,
                "Chi Tiết Tự Luận": essay_summary,
                "Thời Gian Nộp Bài": submission_time
            }
            
            try:
                df = pd.read_excel(SCORES_FILE) if os.path.exists(SCORES_FILE) else pd.DataFrame(columns=EXPECTED_COLUMNS)
                df = pd.concat([df, pd.DataFrame([result_record])], ignore_index=True)
                df.to_excel(SCORES_FILE, index=False)
            except Exception as ex:
                st.error(f"Lưu file cục bộ thất bại: {ex}")
                
            ok, msg = save_to_supabase(result_record)
            if not ok:
                st.warning(f"Lưu Supabase: {msg}")

            st.session_state["quiz_submitted"] = True 
            st.session_state["doing_quiz"] = False
            st.success(f"Điểm trắc nghiệm tạm tính: {final_mcq_score_10}/8.0 điểm. Phần tự luận sẽ được giáo viên chấm sau.")
            time.sleep(2)
            st.rerun()
        
        if remaining > 0 and not submitted:
            time.sleep(1)
            st.rerun()
        return

    if is_submitted and not doing_quiz:
        st.markdown("---")
