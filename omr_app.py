import streamlit as st
import pandas as pd
import cv2
import numpy as np
import tempfile

def load_answer_key(file):
    df = pd.read_excel(file)
    return {int(row['Qno']): str(row['Answer']).strip().upper() for _, row in df.iterrows()}

def process_omr(file, answer_key):
    bytes_data = file.read()
    np_arr = np.frombuffer(bytes_data, np.uint8)
    image = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    bubble_contours = [c for c in contours if 300 < cv2.contourArea(c) < 800]
    bubble_contours = sorted(bubble_contours, key=lambda c: (cv2.boundingRect(c)[1], cv2.boundingRect(c)[0]))

    debug_img = image.copy()
    cv2.drawContours(debug_img, bubble_contours, -1, (0, 255, 0), 2)

    questions = []
    for i in range(0, len(bubble_contours), 20):
        row = bubble_contours[i:i+20]
        row = sorted(row, key=lambda c: cv2.boundingRect(c)[0])
        questions.append(row)

    detected_answers = {}
    for row_idx, row in enumerate(questions):
        for subj_idx in range(5):
            q_index = row_idx + subj_idx * len(questions) + 1
            col_bubbles = row[subj_idx*4:(subj_idx+1)*4]
            filled = []
            for bubble in col_bubbles:
                mask = np.zeros(thresh.shape, dtype="uint8")
                cv2.drawContours(mask, [bubble], -1, 255, -1)
                total = cv2.countNonZero(cv2.bitwise_and(thresh, thresh, mask=mask))
                filled.append(total)
            if filled and max(filled) > 0.5 * np.mean(filled):
                selected = "ABCD"[np.argmax(filled)]
                detected_answers[q_index] = selected
            else:
                detected_answers[q_index] = None

    total_questions = len(answer_key)
    total_score = 0
    subject_scores = {"Python":0,"Data Analysis":0,"MySQL":0,"Power BI":0,"Adv Stats":0}
    for q, ans in answer_key.items():
        if q in detected_answers and detected_answers[q] == ans:
            total_score += 1
            if 1 <= q <= 20: subject_scores["Python"] += 1
            elif 21 <= q <= 40: subject_scores["Data Analysis"] += 1
            elif 41 <= q <= 60: subject_scores["MySQL"] += 1
            elif 61 <= q <= 80: subject_scores["Power BI"] += 1
            elif 81 <= q <= 100: subject_scores["Adv Stats"] += 1

    percentage = (total_score / total_questions) * 100
    return {
        "StudentFile": file.name,
        "TotalScore": total_score,
        "TotalQuestions": total_questions,
        "Percentage": percentage,
        "Python": subject_scores["Python"],
        "Data Analysis": subject_scores["Data Analysis"],
        "MySQL": subject_scores["MySQL"],
        "Power BI": subject_scores["Power BI"],
        "Adv Stats": subject_scores["Adv Stats"]
    }, debug_img

st.title("Automated OMR Evaluation System")
uploaded_key = st.file_uploader("Upload Answer Key Excel", type=["xls", "xlsx"])
uploaded_files = st.file_uploader("Upload OMR Sheets (JPEG)", type=["jpeg"], accept_multiple_files=True)

if uploaded_key:
    answer_key = load_answer_key(uploaded_key)
    if uploaded_files:
        results = []
        for file in uploaded_files:
            result, debug_img = process_omr(file, answer_key)
            results.append(result)
            st.image(cv2.cvtColor(debug_img, cv2.COLOR_BGR2RGB), caption=f"Detected Bubbles - {file.name}", use_container_width=True)
        df_results = pd.DataFrame(results)
        st.success("Evaluation complete")
        st.dataframe(df_results)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            df_results.to_excel(tmp.name, index=False)
            st.download_button(
                "Download Results Excel",
                data=open(tmp.name, "rb").read(),
                file_name="evaluation_results.xlsx"
            )
