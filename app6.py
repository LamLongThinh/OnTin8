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

# ====== HÀM XỬ LÝ FILE WORD THẾ HỆ MỚI (CHÍNH XÁC 100%) ======
def parse_quiz_from_word(file):
    doc = Document(file)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    quiz = []
    essays = []
    
    current_mode = "mcq" # mcq hoặc essay
    current_q = None
    
    for text in paragraphs:
        # Chuyển đổi qua phần tự luận nếu gặp tiêu đề tự luận
        if re.search(r"(II\.\s*PHẦN\s*TỰ\s*LUẬN|Tự\s*luận|Câu\s*hỏi\s*tự\s*luận)", text, flags=re.IGNORECASE):
            if current_q and current_mode == "mcq":
                quiz.append(current_q)
                current_q = None
            current_mode = "essay"
            continue
            
        if current_mode == "mcq":
            # Nhận diện câu hỏi trắc nghiệm mới bắt đầu bằng "Câu X:"
            if re.match(r"^Câu\s*\d+[:.]", text, flags=re.IGNORECASE):
                if current_q:
                    quiz.append(current_q)
                q_text = re.sub(r"^Câu\s*\d+[:.]\s*", "", text).strip()
                current_q = {"question": q_text, "options": [], "answer": "", "image_base64": None}
            elif current_q and re.match(r"^[A-D][\.\)]", text):
                # Nhận diện các lựa chọn A, B, C, D
                parts = re.split(r"^[A-D][\.\)]\s*", text)
                opt_content = parts[1] if len(parts) > 1 else text
                current_q["options"].append(opt_content.strip())
            elif current_q and re.search(r"đáp\s*án", text, flags=re.IGNORECASE):
                # Nhận diện đáp án đúng
                ans_letter = text.split(":")[-1].strip().upper()
                # Gán nội dung đáp án tương ứng với chữ cái A, B, C, D
                letter_map = {"A": 0, "B": 1, "C": 2, "D": 3}
                if ans_letter in letter_map and len(current_q["options"]) > letter_map[ans_letter]:
                    current_q["answer"] = current_q["options"][letter_map[ans_letter]]
                    
        elif current_mode == "essay":
            # Nhận diện câu tự luận
            if re.match(r"^(Câu\s*\d+|Bài\s*\d+)", text, flags=re.IGNORECASE):
                e_text = re.sub(r"^(Câu\s*\d+|Bài\s*\d+)[:.]?\s*", "", text).strip()
                essays.append({"question": e_text if e_text else text})
            elif essays:
                # Nối tiếp dòng nội dung dài của câu tự luận nếu xuống dòng
                essays[-1]["question"] += " " + text
            else:
                essays.append({"question": text})
                
    # Thêm câu hỏi cuối cùng của trắc nghiệm nếu còn sót
    if current_q and current_mode == "mcq":
        quiz.append(current_q)
        
    # Dự phòng nếu file không có câu tự luận nào
    if not essays:
        essays = [
            {"question": "Câu hỏi tự luận 1 (Giáo viên cập nhật nội dung)."},
            {"question": "Câu hỏi tự luận 2 (Giáo viên cập nhật nội dung)."}
        ]

    return quiz[:40], essays[:2]

def save_quiz_data(quiz, essays):
    with open(QUIZ_FILE, "w", encoding="utf-8") as f:
        json.dump(quiz, f, ensure_ascii=False, indent=4)
    with open(ESSAY_FILE, "w", encoding="utf-8") as f:
        json.dump(essays, f, ensure_ascii=False, indent=4)
    st.success(f"✅ Đã lưu thành công {len(quiz)} câu trắc nghiệm và {len(essays)} câu tự luận!")

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
        st.subheader("🔍 Kết quả & Đáp án trắc nghiệm")
        mcq_score = st.session_state.get('mcq_score', 0)
        total = len(st.session_state.get('quiz_data', []))
        score_10 = st.session_state.get('final_mcq_score_10', 0)
        
        st.success(f"Điểm trắc nghiệm: **{mcq_score}/{total} câu đúng** ({score_10}/8.0 điểm)")
        st.info("Phần tự luận đã được ghi nhận và gửi giáo viên chấm điểm.")

        quiz = st.session_state.get("quiz_data", [])
        for idx, q in enumerate(quiz, start=1):
            correct_answer = q['answer']
            student_choice = st.session_state["responses"].get(q["question"])
            is_correct = student_choice == correct_answer
            feedback_icon = "✅" if is_correct else "❌"
            
            st.markdown(f"**{feedback_icon} Câu {idx}:** {q['question']}")
            for option in q['options']:
                color = "#e6ffe6" if option == correct_answer else ("#ffe6e6" if option == student_choice else "#f9f9f9")
                st.markdown(f"<div style='background-color: {color}; padding: 5px; margin: 2px 0; border-radius: 4px;'>{option}</div>", unsafe_allow_html=True)
            st.markdown("---")
                
        if st.button("🚀 Làm bài mới / Đăng xuất", type="primary"):
            for key in ["quiz_data", "essay_data", "responses", "essay_responses", "start_time", "doing_quiz", "quiz_submitted", "mcq_score", "stu_name", "stu_class"]:
                if key in st.session_state: del st.session_state[key]
            st.rerun()
        return

# =========================================================================
# GIAO DIỆN GIÁO VIÊN
# =========================================================================
def admin_ui():
    if not st.session_state.get("admin_logged_in", False):
        st.info("🔐 Đăng nhập khu vực Giáo viên")
        pwd = st.text_input("Mật khẩu quản trị:", type="password")
        if st.button("Đăng nhập", type="primary"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.success("Đăng nhập thành công!")
                st.rerun()
            else:
                st.error("Sai mật khẩu!")
        return

    st.success("✅ Đã đăng nhập với quyền Giáo viên.")
    if st.button("🚪 Đăng xuất"):
        st.session_state.admin_logged_in = False
        st.rerun()

    st.header("👨‍🏫 Bảng Điều Khiển Giáo Viên")
    
    st.subheader("1️⃣ Cấu hình hệ thống & GitHub Sync")
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        sup_url = st.text_input("Supabase URL:", value=st.session_state.get("supabase_url", ""))
        st.session_state["supabase_url"] = sup_url
    with col_g2:
        sup_key = st.text_input("Supabase API Key:", type="password", value=st.session_state.get("supabase_key", ""))
        st.session_state["supabase_key"] = sup_key

    if st.button("🔄 Đồng bộ đề thi tự động từ GitHub (OnTin8)"):
        file_obj = fetch_docx_from_github()
        if file_obj:
            q, e = parse_quiz_from_word(file_obj)
            save_quiz_data(q, e)
            st.success(f"Đã cập nhật từ GitHub thành công: {len(q)} trắc nghiệm, {len(e)} tự luận.")
            st.rerun()
        else:
            st.error("Không tìm thấy file word hoặc lỗi kết nối tới repository GitHub.")

    st.markdown("Hoặc tải lên file đề Word (.docx) thủ công:")
    up = st.file_uploader("📄 Chọn file .docx ôn tập", type=["docx"])
    if up:
        q, e = parse_quiz_from_word(up)
        save_quiz_data(q, e)
        st.success(f"Đã tải lên và lưu {len(q)} trắc nghiệm, {len(e)} tự luận.")

    st.subheader("2️⃣ Chấm điểm Tự luận học sinh")
    if os.path.exists(SCORES_FILE) and os.path.getsize(SCORES_FILE) > 0:
        try:
            df_scores = pd.read_excel(SCORES_FILE)
            if not df_scores.empty:
                st.dataframe(df_scores, use_container_width=True)
                
                student_names = df_scores["Tên Học Sinh"].tolist()
                selected_student = st.selectbox("Chọn học sinh để chấm điểm tự luận:", student_names)
                
                if selected_student:
                    stu_row = df_scores[df_scores["Tên Học Sinh"] == selected_student].iloc[0]
                    st.write(f"**Học sinh:** {stu_row['Tên Học Sinh']} - **Lớp:** {stu_row['Lớp']}")
                    
                    try:
                        essay_details = json.loads(stu_row["Chi Tiết Tự Luận"])
                    except:
                        essay_details = {}
                        
                    with st.form("grading_form"):
                        st.markdown("##### Bài làm tự luận của học sinh:")
                        essay_scores = []
                        for idx, (q_text, ans_text) in enumerate(essay_details.items(), 1):
                            st.markdown(f"**Q{idx}:** {q_text}")
                            st.info(f"Trả lời: {ans_text if ans_text else '(Chưa trả lời)'}")
                            escore = st.number_input(f"Điểm câu tự luận {idx} (0 - 1.0 điểm):", 0.0, 1.0, 0.0, step=0.1, key=f"grade_es_{idx}")
                            essay_scores.append(escore)
                            st.markdown("---")
                            
                        update_grade_btn = st.form_submit_button("💾 Lưu điểm tự luận", type="primary")
                        
                    if update_grade_btn:
                        total_essay_score = round(sum(essay_scores), 2)
                        mcq_val = float(stu_row["Điểm Trắc Nghiệm"])
                        total_final = round(mcq_val + total_essay_score, 2)
                        
                        df_scores.loc[df_scores["Tên Học Sinh"] == selected_student, "Điểm Tự Luận"] = total_essay_score
                        df_scores.loc[df_scores["Tên Học Sinh"] == selected_student, "Tổng Điểm"] = total_final
                        df_scores.to_excel(SCORES_FILE, index=False)
                        st.success(f"Đã cập nhật điểm tự luận: {total_essay_score}đ. Tổng điểm mới: {total_final}đ")
                        time.sleep(1)
                        st.rerun()
        except Exception as e:
            st.error(f"Lỗi đọc bảng điểm: {e}")
    else:
        st.info("Chưa có dữ liệu bảng điểm học sinh.")

    st.subheader("3️⃣ Xuất Bảng Điểm")
    if os.path.exists(SCORES_FILE):
        with open(SCORES_FILE, "rb") as f:
            st.download_button("📥 Tải File Excel Bảng Điểm", f, file_name=f"BangDiem_OnTin8_{dt.date.today()}.xlsx")

# ====== Điều hướng chính ======
def main():
    if "mode" not in st.session_state:
        st.session_state.mode = "student"
        
    with st.sidebar:
        st.sidebar.markdown(
            """
            <h3 style='text-align: center; color: #444; font-weight: 800;'>
                Trường THCS Mỹ Hòa
            </h3>
            <hr style='margin-top: -10px; margin-bottom: 10px;'>
            """,
            unsafe_allow_html=True
        )
        st.header("⚙️ Chế độ Ứng dụng")
        mode = st.radio("Chọn chế độ:", ["Học sinh", "Giáo viên"], index=0 if st.session_state.mode == "student" else 1)
        st.session_state.mode = "student" if mode == "Học sinh" else "admin"

    if st.session_state.mode == "student":
        student_ui()
    else:
        admin_ui()

if __name__ == "__main__":
    main()