# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import random
import json
import uuid
from datetime import datetime, timedelta

# =========================
# CONFIG PAGE
# =========================
st.set_page_config(
    page_title="QCM AMF",
    page_icon="📘",
    layout="wide"
)

# =========================
# STYLE
# =========================
st.markdown("""
<style>
.main-title {
    font-size: 2.2rem;
    font-weight: 800;
    margin-bottom: 0.2rem;
}
.sub-title {
    color: #666;
    margin-bottom: 1.2rem;
}
.question-box {
    background-color: #f8f9fa;
    padding: 20px;
    border-radius: 16px;
    border: 1px solid #e6e6e6;
    margin-bottom: 18px;
}
.correct {
    color: #15803d;
    font-weight: 700;
}
.wrong {
    color: #b91c1c;
    font-weight: 700;
}
.info-box {
    background-color: #eef6ff;
    padding: 12px;
    border-radius: 10px;
    border-left: 5px solid #3b82f6;
    margin-bottom: 12px;
}
.success-box {
    background-color: #ecfdf5;
    padding: 14px;
    border-radius: 12px;
    border-left: 5px solid #10b981;
    margin-bottom: 14px;
}
.danger-box {
    background-color: #fef2f2;
    padding: 14px;
    border-radius: 12px;
    border-left: 5px solid #ef4444;
    margin-bottom: 14px;
}
.exam-box {
    background-color: #fff7ed;
    padding: 14px;
    border-radius: 12px;
    border-left: 5px solid #f97316;
    margin-bottom: 14px;
}
.review-box {
    background-color: #fffbea;
    padding: 10px;
    border-radius: 10px;
    border-left: 5px solid #eab308;
    margin-bottom: 12px;
}
.nav-box {
    background-color: #f9fafb;
    padding: 12px;
    border-radius: 10px;
    border: 1px solid #e5e7eb;
    margin-bottom: 12px;
}
.small-muted {
    color: #6b7280;
    font-size: 0.95rem;
}
</style>
""", unsafe_allow_html=True)

# =========================
# PARAMÈTRES
# =========================
EXCEL_FILE = "database_amf.xlsx"

PROFILES = {
    "Profil 1": {
        "wrong_questions": "wrong_questions.json",
        "seen_questions": "seen_questions.json",
        "saved_sessions": "saved_sessions.json",
    },
    "Profil 2": {
        "wrong_questions": "wrong_questions_andrew.json",
        "seen_questions": "seen_questions_andrew.json",
        "saved_sessions": "saved_sessions_andrew.json",
    }
}

profile = st.sidebar.selectbox("👤 Profil", list(PROFILES.keys()))

WRONG_QUESTIONS_FILE = PROFILES[profile]["wrong_questions"]
SEEN_QUESTIONS_FILE = PROFILES[profile]["seen_questions"]
SAVED_SESSIONS_FILE = PROFILES[profile]["saved_sessions"]

EXAM_A_COUNT = 33
EXAM_C_COUNT = 87
EXAM_A_PASS = 27
EXAM_C_PASS = 70
EXAM_TOTAL = 120
EXAM_DURATION_MINUTES = 120

# =========================
# OUTILS JSON
# =========================
def load_json_list(filepath: str) -> list:
    path = Path(filepath)
    if not path.exists():
        return []

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return [str(x) for x in data]
            return []
    except Exception:
        return []


def save_json_list(filepath: str, values: list):
    values = sorted(list(set(str(x) for x in values)))
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(values, f, ensure_ascii=False, indent=2)


def add_json_values(filepath: str, new_values: list):
    existing = load_json_list(filepath)
    merged = list(set(existing + [str(x) for x in new_values]))
    save_json_list(filepath, merged)


def remove_json_values(filepath: str, values_to_remove: list):
    existing = set(load_json_list(filepath))
    values_to_remove = set(str(x) for x in values_to_remove)
    updated = list(existing - values_to_remove)
    save_json_list(filepath, updated)


# =========================
# SESSIONS SAUVEGARDÉES
# =========================
def load_saved_sessions() -> list:
    path = Path(SAVED_SESSIONS_FILE)
    if not path.exists():
        return []

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
            return []
    except Exception:
        return []


def save_saved_sessions(sessions: list):
    with open(SAVED_SESSIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(sessions, f, ensure_ascii=False, indent=2)


def save_session(name: str, qcm_df: pd.DataFrame, mode: str):
    sessions = load_saved_sessions()

    session_data = {
        "id": str(uuid.uuid4()),
        "name": name.strip() if name.strip() else f"Session {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "mode": mode,
        "question_ids": qcm_df["n°identifiant"].astype(str).tolist()
    }

    sessions.append(session_data)
    save_saved_sessions(sessions)


def rename_session(session_id: str, new_name: str):
    sessions = load_saved_sessions()

    for session in sessions:
        if session["id"] == session_id:
            session["name"] = new_name.strip()
            break

    save_saved_sessions(sessions)


def delete_session(session_id: str):
    sessions = load_saved_sessions()
    sessions = [s for s in sessions if s["id"] != session_id]
    save_saved_sessions(sessions)


def load_session_qcm(df: pd.DataFrame, session_id: str) -> pd.DataFrame:
    sessions = load_saved_sessions()
    session = next((s for s in sessions if s["id"] == session_id), None)

    if session is None:
        raise ValueError("Session introuvable.")

    question_ids = [str(x) for x in session["question_ids"]]
    session_df = df[df["n°identifiant"].astype(str).isin(question_ids)].copy()

    if session_df.empty:
        raise ValueError("Impossible de recharger cette session.")

    order_map = {qid: i for i, qid in enumerate(question_ids)}
    session_df["__order"] = session_df["n°identifiant"].astype(str).map(order_map)
    session_df = session_df.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    return session_df


# =========================
# FONCTIONS CHARGEMENT EXCEL
# =========================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    expected_first_col = "n°identifiant"
    header_idx = None

    for i in range(len(df)):
        row_values = [str(x).strip().lower() for x in df.iloc[i].tolist()]
        if expected_first_col.lower() in row_values:
            header_idx = i
            break

    if header_idx is None:
        return pd.DataFrame()

    new_header = df.iloc[header_idx].tolist()
    clean_df = df.iloc[header_idx + 1:].copy()
    clean_df.columns = new_header
    return clean_df


def clean_question_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    required_cols = [
        "n°identifiant",
        "Theme",
        "Sous_theme",
        "Question_Categorie",
        "question contenant le numéro unique",
        "Choix_A",
        "Choix_B",
        "Choix_C",
        "Reponse",
    ]

    if not all(col in df.columns for col in required_cols):
        return pd.DataFrame()

    df = df[required_cols].copy()
    df = df.dropna(subset=["question contenant le numéro unique", "Reponse"])

    for col in required_cols:
        df[col] = df[col].astype(str).str.strip()

    df = df[
        df["Reponse"].isin(["A", "B", "C"])
        & df["Choix_A"].ne("")
        & df["Choix_B"].ne("")
        & df["Choix_C"].ne("")
        & df["question contenant le numéro unique"].ne("")
    ]

    df = df[df["n°identifiant"].str.contains(r"\d", na=False)]

    return df.reset_index(drop=True)


@st.cache_data
def load_amf_database(excel_path: str) -> pd.DataFrame:
    path = Path(excel_path)

    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {excel_path}")

    xls = pd.ExcelFile(path)
    all_questions = []

    for sheet_name in xls.sheet_names:
        if "supprim" in sheet_name.lower():
            continue

        raw_df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        normalized_df = normalize_columns(raw_df)

        if normalized_df.empty:
            continue

        clean_df = clean_question_dataframe(normalized_df)

        if not clean_df.empty:
            clean_df["SourceSheet"] = sheet_name
            all_questions.append(clean_df)

    if not all_questions:
        raise ValueError("Aucune question valide trouvée dans le fichier Excel.")

    final_df = pd.concat(all_questions, ignore_index=True)
    final_df = final_df.drop_duplicates(subset=["n°identifiant"]).reset_index(drop=True)

    final_df["Theme"] = final_df["Theme"].astype(str).str.strip()
    final_df["Sous_theme"] = final_df["Sous_theme"].astype(str).str.strip()
    final_df["n°identifiant"] = final_df["n°identifiant"].astype(str).str.strip()
    final_df["Question_Categorie"] = final_df["Question_Categorie"].astype(str).str.strip().str.upper()

    return final_df


# =========================
# FONCTIONS MÉTIER
# =========================
def get_correct_answer_text(row: pd.Series) -> str:
    if row["Reponse"] == "A":
        return row["Choix_A"]
    if row["Reponse"] == "B":
        return row["Choix_B"]
    if row["Reponse"] == "C":
        return row["Choix_C"]
    return "Réponse inconnue"


def apply_filters(df: pd.DataFrame, theme=None, sous_theme=None, category=None) -> pd.DataFrame:
    filtered_df = df.copy()

    if category and category != "Toutes":
        filtered_df = filtered_df[
            filtered_df["Question_Categorie"].astype(str).str.upper() == category
        ]

    if theme and theme != "Tous":
        filtered_df = filtered_df[
            filtered_df["Theme"].astype(str) == str(theme)
        ]

    if sous_theme and sous_theme != "Tous":
        filtered_df = filtered_df[
            filtered_df["Sous_theme"].astype(str) == str(sous_theme)
        ]

    return filtered_df


def generate_qcm(df: pd.DataFrame, n_questions: int, theme=None, sous_theme=None, category=None) -> pd.DataFrame:
    filtered_df = apply_filters(df, theme, sous_theme, category)

    if filtered_df.empty:
        raise ValueError("Aucune question trouvée avec ces filtres.")

    n_questions = min(n_questions, len(filtered_df))
    return filtered_df.sample(
        n=n_questions,
        random_state=random.randint(1, 999999)
    ).reset_index(drop=True)


def generate_full_random_qcm(df: pd.DataFrame, n_questions: int, category=None) -> pd.DataFrame:
    filtered_df = df.copy()

    if category and category != "Toutes":
        filtered_df = filtered_df[
            filtered_df["Question_Categorie"].astype(str).str.upper() == category
        ]

    if filtered_df.empty:
        raise ValueError("Aucune question trouvée.")

    n_questions = min(n_questions, len(filtered_df))
    return filtered_df.sample(
        n=n_questions,
        random_state=random.randint(1, 999999)
    ).reset_index(drop=True)


def generate_never_seen_qcm(df: pd.DataFrame, n_questions: int, theme=None, sous_theme=None, category=None) -> pd.DataFrame:
    seen_ids = set(load_json_list(SEEN_QUESTIONS_FILE))
    filtered_df = apply_filters(df, theme, sous_theme, category)
    filtered_df = filtered_df[~filtered_df["n°identifiant"].isin(seen_ids)]

    if filtered_df.empty:
        raise ValueError("Il n'y a plus de questions jamais vues avec ces filtres.")

    n_questions = min(n_questions, len(filtered_df))
    return filtered_df.sample(
        n=n_questions,
        random_state=random.randint(1, 999999)
    ).reset_index(drop=True)


def generate_exam_qcm(df: pd.DataFrame) -> pd.DataFrame:
    df_a = df[df["Question_Categorie"] == "A"].copy()
    df_c = df[df["Question_Categorie"] == "C"].copy()

    if len(df_a) < EXAM_A_COUNT:
        raise ValueError(f"Pas assez de questions A dans la base ({len(df_a)} trouvées).")
    if len(df_c) < EXAM_C_COUNT:
        raise ValueError(f"Pas assez de questions C dans la base ({len(df_c)} trouvées).")

    sample_a = df_a.sample(n=EXAM_A_COUNT, random_state=random.randint(1, 999999)).copy()
    sample_c = df_c.sample(n=EXAM_C_COUNT, random_state=random.randint(1, 999999)).copy()

    sample_a["Exam_Block"] = "A"
    sample_c["Exam_Block"] = "C"

    exam_df = pd.concat([sample_a, sample_c], ignore_index=True)
    exam_df = exam_df.sample(frac=1, random_state=random.randint(1, 999999)).reset_index(drop=True)
    return exam_df


def get_wrong_questions_df(df: pd.DataFrame) -> pd.DataFrame:
    wrong_ids = load_json_list(WRONG_QUESTIONS_FILE)
    if not wrong_ids:
        return pd.DataFrame()
    return df[df["n°identifiant"].isin(wrong_ids)].copy().reset_index(drop=True)


def get_seen_questions_df(df: pd.DataFrame) -> pd.DataFrame:
    seen_ids = load_json_list(SEEN_QUESTIONS_FILE)
    if not seen_ids:
        return pd.DataFrame()
    return df[df["n°identifiant"].isin(seen_ids)].copy().reset_index(drop=True)


def start_exam_timer():
    now = datetime.now()
    st.session_state.exam_start_time = now.isoformat()
    st.session_state.exam_end_time = (now + timedelta(minutes=EXAM_DURATION_MINUTES)).isoformat()


def format_seconds(seconds: int) -> str:
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def get_exam_time_left():
    if not st.session_state.exam_end_time:
        return None
    end_time = datetime.fromisoformat(st.session_state.exam_end_time)
    seconds_left = int((end_time - datetime.now()).total_seconds())
    return max(0, seconds_left)


def compute_results(qcm_df: pd.DataFrame, user_answers: dict):
    score_total = 0
    score_a = 0
    score_c = 0
    total_a = 0
    total_c = 0
    wrong_ids_to_add = []

    for i, row in qcm_df.iterrows():
        cat = str(row["Question_Categorie"]).upper()
        is_correct = user_answers.get(i) == row["Reponse"]

        if cat == "A":
            total_a += 1
            if is_correct:
                score_a += 1
        elif cat == "C":
            total_c += 1
            if is_correct:
                score_c += 1

        if is_correct:
            score_total += 1
        else:
            wrong_ids_to_add.append(str(row["n°identifiant"]))

    admitted = (score_a >= EXAM_A_PASS) and (score_c >= EXAM_C_PASS)
    percentage = round((score_total / len(qcm_df)) * 100, 2) if len(qcm_df) > 0 else 0.0

    return {
        "score_total": score_total,
        "score_a": score_a,
        "score_c": score_c,
        "total_a": total_a,
        "total_c": total_c,
        "percentage": percentage,
        "admitted": admitted,
        "wrong_ids_to_add": wrong_ids_to_add
    }


def get_correct_ids(qcm_df: pd.DataFrame, user_answers: dict) -> list:
    correct_ids = []

    for i, row in qcm_df.iterrows():
        if user_answers.get(i) == row["Reponse"]:
            correct_ids.append(str(row["n°identifiant"]))

    return correct_ids


def get_answered_count(qcm_df: pd.DataFrame, answers: dict) -> int:
    return sum(1 for i in range(len(qcm_df)) if answers.get(i) in ["A", "B", "C"])


def get_review_count(marked_for_review: set) -> int:
    return len(marked_for_review)


def question_status(i: int, answers: dict, marked_for_review: set) -> str:
    if i in marked_for_review:
        return "À revoir"
    if answers.get(i) in ["A", "B", "C"]:
        return "Répondue"
    return "Non répondue"


# =========================
# SESSION STATE
# =========================
defaults = {
    "qcm_df": None,
    "submitted": False,
    "user_answers": {},
    "last_mode": None,
    "show_correction": True,
    "exam_start_time": None,
    "exam_end_time": None,
    "current_question_index": 0,
    "marked_for_review": set(),
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# =========================
# CHARGEMENT BASE
# =========================
st.markdown('<div class="main-title">📘 QCM AMF</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Modes aléatoire, jamais vues, erreurs, examen, sessions et statistiques</div>', unsafe_allow_html=True)

try:
    df = load_amf_database(EXCEL_FILE)
except Exception as e:
    st.error(f"Erreur lors du chargement du fichier Excel : {e}")
    st.stop()

wrong_df_global = get_wrong_questions_df(df)
seen_df_global = get_seen_questions_df(df)

# =========================
# SIDEBAR
# =========================
st.sidebar.header("⚙️ Paramètres")

category_filter = st.sidebar.selectbox(
    "Catégorie de question",
    ["Toutes", "A", "C"]
)

quiz_mode = st.sidebar.radio(
    "Mode",
    [
        "QCM aléatoire filtré",
        "QCM full aléatoire",
        "QCM des erreurs",
        "QCM des questions jamais vues",
        "Mode examen AMF"
    ],
    index=0
)

themes = ["Tous"] + sorted(df["Theme"].dropna().astype(str).unique().tolist())
selected_theme = st.sidebar.selectbox("Thème", themes)

if selected_theme != "Tous":
    sous_themes_list = (
        ["Tous"]
        + sorted(
            df[df["Theme"] == selected_theme]["Sous_theme"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
    )
else:
    sous_themes_list = ["Tous"] + sorted(df["Sous_theme"].dropna().astype(str).unique().tolist())

selected_sous_theme = st.sidebar.selectbox("Sous-thème", sous_themes_list)

if quiz_mode != "Mode examen AMF":
    n_questions = st.sidebar.slider("Nombre de questions", min_value=5, max_value=100, value=20, step=5)
else:
    n_questions = EXAM_TOTAL
    st.sidebar.success("Mode examen : 33 A + 87 C")
    st.sidebar.info(f"Partie A : {EXAM_A_COUNT} questions | seuil {EXAM_A_PASS}")
    st.sidebar.info(f"Partie C : {EXAM_C_COUNT} questions | seuil {EXAM_C_PASS}")
    st.sidebar.info(f"Durée : {EXAM_DURATION_MINUTES} minutes")
    st.sidebar.warning("Le filtre catégorie/thème/sous-thème ne s'applique pas au mode examen.")

st.sidebar.markdown("---")
st.sidebar.info(f"Questions totales : {len(df)}")
st.sidebar.info(f"Questions A : {len(df[df['Question_Categorie'] == 'A'])}")
st.sidebar.info(f"Questions C : {len(df[df['Question_Categorie'] == 'C'])}")
st.sidebar.info(f"Questions déjà vues : {len(seen_df_global)}")
st.sidebar.info(f"Questions dans la liste d'erreurs : {len(wrong_df_global)}")

if st.sidebar.button("🗑️ Vider la liste des erreurs", use_container_width=True):
    save_json_list(WRONG_QUESTIONS_FILE, [])
    st.success("Liste des erreurs vidée.")
    st.rerun()

if st.sidebar.button("🧹 Réinitialiser les questions vues", use_container_width=True):
    save_json_list(SEEN_QUESTIONS_FILE, [])
    st.success("Historique des questions vues réinitialisé.")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("💾 Sessions sauvegardées")

saved_sessions = load_saved_sessions()

if saved_sessions:
    session_labels = [
        f"{s['name']} | {s['mode']} | {s['created_at']}"
        for s in saved_sessions
    ]

    selected_session_label = st.sidebar.selectbox(
        "Choisir une session",
        options=session_labels
    )

    selected_session = saved_sessions[session_labels.index(selected_session_label)]

    new_session_name = st.sidebar.text_input(
        "Renommer la session",
        value=selected_session["name"]
    )

    col_s1, col_s2 = st.sidebar.columns(2)

    with col_s1:
        if st.button("▶️ Rejouer"):
            try:
                qcm_df = load_session_qcm(df, selected_session["id"])
                st.session_state.qcm_df = qcm_df
                st.session_state.submitted = False
                st.session_state.user_answers = {}
                st.session_state.last_mode = f"Session rejouée - {selected_session['name']}"
                st.session_state.current_question_index = 0
                st.session_state.marked_for_review = set()
                st.session_state.show_correction = True
                st.session_state.exam_start_time = None
                st.session_state.exam_end_time = None
                st.rerun()
            except Exception as e:
                st.sidebar.error(str(e))

    with col_s2:
        if st.button("🗑️ Supprimer"):
            delete_session(selected_session["id"])
            st.sidebar.success("Session supprimée.")
            st.rerun()

    if st.sidebar.button("✏️ Renommer la session"):
        if new_session_name.strip():
            rename_session(selected_session["id"], new_session_name)
            st.sidebar.success("Session renommée.")
            st.rerun()
        else:
            st.sidebar.warning("Le nom ne peut pas être vide.")
else:
    st.sidebar.info("Aucune session sauvegardée.")

# =========================
# BOUTONS LANCEMENT
# =========================
col1, col2 = st.columns([1, 1])

with col1:
    if st.button("🎲 Lancer le QCM", use_container_width=True):
        try:
            if quiz_mode == "QCM aléatoire filtré":
                qcm_df = generate_qcm(
                    df=df,
                    n_questions=n_questions,
                    theme=selected_theme,
                    sous_theme=selected_sous_theme,
                    category=category_filter
                )
                st.session_state.show_correction = True
                st.session_state.exam_start_time = None
                st.session_state.exam_end_time = None

            elif quiz_mode == "QCM full aléatoire":
                qcm_df = generate_full_random_qcm(
                    df=df,
                    n_questions=n_questions,
                    category=category_filter
                )
                st.session_state.show_correction = True
                st.session_state.exam_start_time = None
                st.session_state.exam_end_time = None

            elif quiz_mode == "QCM des erreurs":
                source_df = get_wrong_questions_df(df)
                source_df = apply_filters(source_df, selected_theme, selected_sous_theme, category_filter)

                if source_df.empty:
                    raise ValueError("Aucune question dans la liste d'erreurs avec ce filtre.")

                n_real = min(n_questions, len(source_df))
                qcm_df = source_df.sample(
                    n=n_real,
                    random_state=random.randint(1, 999999)
                ).reset_index(drop=True)

                st.session_state.show_correction = True
                st.session_state.exam_start_time = None
                st.session_state.exam_end_time = None

            elif quiz_mode == "QCM des questions jamais vues":
                qcm_df = generate_never_seen_qcm(
                    df=df,
                    n_questions=n_questions,
                    theme=selected_theme,
                    sous_theme=selected_sous_theme,
                    category=category_filter
                )
                st.session_state.show_correction = True
                st.session_state.exam_start_time = None
                st.session_state.exam_end_time = None

            else:
                qcm_df = generate_exam_qcm(df)
                st.session_state.show_correction = False
                start_exam_timer()

            add_json_values(SEEN_QUESTIONS_FILE, qcm_df["n°identifiant"].astype(str).tolist())

            st.session_state.qcm_df = qcm_df
            st.session_state.submitted = False
            st.session_state.user_answers = {}
            st.session_state.last_mode = quiz_mode
            st.session_state.current_question_index = 0
            st.session_state.marked_for_review = set()
            st.rerun()

        except Exception as e:
            st.error(str(e))

with col2:
    if st.button("🔄 Réinitialiser l'écran", use_container_width=True):
        st.session_state.qcm_df = None
        st.session_state.submitted = False
        st.session_state.user_answers = {}
        st.session_state.last_mode = None
        st.session_state.show_correction = True
        st.session_state.exam_start_time = None
        st.session_state.exam_end_time = None
        st.session_state.current_question_index = 0
        st.session_state.marked_for_review = set()
        st.rerun()

# =========================
# INFOS
# =========================
st.markdown(
    f"""
    <div class="info-box">
        <b>Mode :</b> {quiz_mode}<br>
        <b>Catégorie :</b> {category_filter}<br>
        <b>Base chargée :</b> {len(df)} questions<br>
        <b>Questions vues :</b> {len(seen_df_global)}<br>
        <b>Liste d'erreurs :</b> {len(wrong_df_global)}<br>
        <b>Filtre :</b> Thème = {selected_theme} | Sous-thème = {selected_sous_theme}
    </div>
    """,
    unsafe_allow_html=True
)

# =========================
# DASHBOARD STATS
# =========================
st.subheader("📊 Tableau de bord")

seen_count = len(seen_df_global)
never_seen_count = len(df) - seen_count
wrong_count = len(wrong_df_global)

d1, d2, d3, d4 = st.columns(4)
d1.metric("Total questions", len(df))
d2.metric("Déjà vues", seen_count)
d3.metric("Jamais vues", never_seen_count)
d4.metric("Questions erreur", wrong_count)

with st.expander("Voir mes stats par catégorie"):
    seen_a = len(seen_df_global[seen_df_global["Question_Categorie"] == "A"]) if not seen_df_global.empty else 0
    seen_c = len(seen_df_global[seen_df_global["Question_Categorie"] == "C"]) if not seen_df_global.empty else 0
    wrong_a = len(wrong_df_global[wrong_df_global["Question_Categorie"] == "A"]) if not wrong_df_global.empty else 0
    wrong_c = len(wrong_df_global[wrong_df_global["Question_Categorie"] == "C"]) if not wrong_df_global.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Vues A", seen_a)
    c2.metric("Vues C", seen_c)
    c3.metric("Erreurs A", wrong_a)
    c4.metric("Erreurs C", wrong_c)

# =========================
# GESTION LISTE ERREURS
# =========================
with st.expander("Gérer la liste des questions ratées"):
    wrong_df_preview = get_wrong_questions_df(df)

    if wrong_df_preview.empty:
        st.info("Aucune question ratée enregistrée.")
    else:
        st.write(f"Nombre de questions enregistrées : {len(wrong_df_preview)}")

        display_df = wrong_df_preview[[
            "n°identifiant",
            "Question_Categorie",
            "Theme",
            "Sous_theme",
            "question contenant le numéro unique"
        ]].copy()

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        removable_ids = wrong_df_preview["n°identifiant"].astype(str).tolist()

        ids_to_remove = st.multiselect(
            "Sélectionne les questions à retirer manuellement de la liste d'erreurs",
            options=removable_ids
        )

        if st.button("➖ Supprimer de la liste d'erreurs", use_container_width=True):
            if ids_to_remove:
                remove_json_values(WRONG_QUESTIONS_FILE, ids_to_remove)
                st.success(f"{len(ids_to_remove)} question(s) retirée(s) de la liste d'erreurs.")
                st.rerun()
            else:
                st.warning("Sélectionne au moins une question.")

# =========================
# TIMER EXAMEN
# =========================
if st.session_state.last_mode == "Mode examen AMF" and st.session_state.qcm_df is not None and not st.session_state.submitted:
    seconds_left = get_exam_time_left()
    if seconds_left is not None:
        st.markdown(
            f"""
            <div class="exam-box">
                <b>Mode examen AMF</b><br>
                Partie A : {EXAM_A_COUNT} questions, minimum {EXAM_A_PASS}<br>
                Partie C : {EXAM_C_COUNT} questions, minimum {EXAM_C_PASS}<br>
                Temps restant : <b>{format_seconds(seconds_left)}</b>
            </div>
            """,
            unsafe_allow_html=True
        )

# =========================
# SAUVEGARDE SESSION
# =========================
if st.session_state.qcm_df is not None and not st.session_state.submitted:
    st.markdown("---")
    st.subheader("💾 Sauvegarder cette session")

    default_session_name = f"{st.session_state.last_mode or 'QCM'} - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    session_name_input = st.text_input(
        "Nom de la session",
        value=default_session_name
    )

    if st.button("💾 Sauvegarder la session", use_container_width=True):
        save_session(
            name=session_name_input,
            qcm_df=st.session_state.qcm_df,
            mode=st.session_state.last_mode or "QCM"
        )
        st.success("Session sauvegardée.")

# =========================
# QUESTIONNAIRE
# =========================
if st.session_state.qcm_df is not None and not st.session_state.submitted:
    qcm_df = st.session_state.qcm_df
    total_questions = len(qcm_df)
    idx = st.session_state.current_question_index
    row = qcm_df.iloc[idx]

    answered_count = get_answered_count(qcm_df, st.session_state.user_answers)
    review_count = get_review_count(st.session_state.marked_for_review)
    non_answered_count = total_questions - answered_count
    progress_ratio = answered_count / total_questions if total_questions > 0 else 0

    st.subheader("📝 Questionnaire")

    if st.session_state.last_mode == "Mode examen AMF":
        st.warning("Mode examen : le corrigé est masqué jusqu’à la fin.")

    c1, c2, c3 = st.columns(3)
    c1.metric("Répondues", answered_count)
    c2.metric("Non répondues", non_answered_count)
    c3.metric("À revoir", review_count)

    st.progress(progress_ratio)

    st.markdown(
        f"""
        <div class="nav-box">
            <b>Question {idx + 1} / {total_questions}</b><br>
            <span class="small-muted">Statut : {question_status(idx, st.session_state.user_answers, st.session_state.marked_for_review)}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    if idx in st.session_state.marked_for_review:
        st.markdown(
            """
            <div class="review-box">
                Cette question est marquée <b>à revoir</b>.
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown('<div class="question-box">', unsafe_allow_html=True)
    st.write(f"**Catégorie : {row['Question_Categorie']}**")
    st.write(row["question contenant le numéro unique"])

    options = {
        "A": row["Choix_A"],
        "B": row["Choix_B"],
        "C": row["Choix_C"],
    }

    previous_answer = st.session_state.user_answers.get(idx, "A")

    selected = st.radio(
        "Choisis une réponse :",
        options=["A", "B", "C"],
        index=["A", "B", "C"].index(previous_answer),
        format_func=lambda x: f"{x}. {options[x]}",
        key=f"question_current_{idx}"
    )

    st.session_state.user_answers[idx] = selected
    st.markdown("</div>", unsafe_allow_html=True)

    nav1, nav2, nav3, nav4 = st.columns([1, 1, 1, 1])

    with nav1:
        if st.button("⬅️ Précédent", use_container_width=True, disabled=(idx == 0)):
            st.session_state.current_question_index -= 1
            st.rerun()

    with nav2:
        if idx in st.session_state.marked_for_review:
            if st.button("✅ Enlever 'à revoir'", use_container_width=True):
                st.session_state.marked_for_review.discard(idx)
                st.rerun()
        else:
            if st.button("🔖 Marquer à revoir", use_container_width=True):
                st.session_state.marked_for_review.add(idx)
                st.rerun()

    with nav3:
        if st.button("➡️ Suivant", use_container_width=True, disabled=(idx == total_questions - 1)):
            st.session_state.current_question_index += 1
            st.rerun()

    with nav4:
        if st.button("✅ Terminer le QCM", use_container_width=True):
            st.session_state.submitted = True
            st.rerun()

    st.markdown("---")

    jump_to = st.selectbox(
        "Aller directement à la question",
        options=list(range(total_questions)),
        index=idx,
        format_func=lambda x: f"Question {x + 1} — {question_status(x, st.session_state.user_answers, st.session_state.marked_for_review)}"
    )

    if jump_to != idx:
        st.session_state.current_question_index = jump_to
        st.rerun()

# =========================
# RÉSULTATS
# =========================
if st.session_state.qcm_df is not None and st.session_state.submitted:
    qcm_df = st.session_state.qcm_df
    user_answers = st.session_state.user_answers

    results = compute_results(qcm_df, user_answers)

    current_wrong_before = set(load_json_list(WRONG_QUESTIONS_FILE))
    correct_ids = get_correct_ids(qcm_df, user_answers)
    removed_from_wrong = [qid for qid in correct_ids if qid in current_wrong_before]

    if results["wrong_ids_to_add"]:
        add_json_values(WRONG_QUESTIONS_FILE, results["wrong_ids_to_add"])

    if correct_ids:
        remove_json_values(WRONG_QUESTIONS_FILE, correct_ids)

    st.subheader("📊 Résultat")

    c1, c2, c3 = st.columns(3)
    c1.metric("Partie A", f"{results['score_a']}/{results['total_a']}")
    c2.metric("Partie C", f"{results['score_c']}/{results['total_c']}")
    c3.metric("Total", f"{results['score_total']}/{len(qcm_df)}")

    extra1, extra2 = st.columns(2)
    extra1.metric("Ajoutées à la liste d'erreurs", len(results["wrong_ids_to_add"]))
    extra2.metric("Retirées de la liste d'erreurs", len(removed_from_wrong))

    if st.session_state.last_mode == "Mode examen AMF":
        if results["admitted"]:
            st.markdown(
                f"""
                <div class="success-box">
                    <b>ADMIS ✅</b><br>
                    Partie A : {results['score_a']}/{results['total_a']} (minimum {EXAM_A_PASS})<br>
                    Partie C : {results['score_c']}/{results['total_c']} (minimum {EXAM_C_PASS})<br>
                    Total : {results['score_total']}/{len(qcm_df)}
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f"""
                <div class="danger-box">
                    <b>NON ADMIS ❌</b><br>
                    Partie A : {results['score_a']}/{results['total_a']} (minimum {EXAM_A_PASS})<br>
                    Partie C : {results['score_c']}/{results['total_c']} (minimum {EXAM_C_PASS})<br>
                    Total : {results['score_total']}/{len(qcm_df)}
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.metric("Pourcentage", f"{results['percentage']}%")

    if st.session_state.last_mode == "Mode examen AMF" and not st.session_state.show_correction:
        if st.button("👁️ Afficher le corrigé", use_container_width=True):
            st.session_state.show_correction = True
            st.rerun()

    if st.session_state.show_correction:
        st.markdown("---")
        st.subheader("📖 Corrigé détaillé")

        for i, row in qcm_df.iterrows():
            user_choice = user_answers.get(i)
            correct_choice = row["Reponse"]
            correct_text = get_correct_answer_text(row)

            st.markdown('<div class="question-box">', unsafe_allow_html=True)
            st.markdown(f"**Question {i+1}**")
            st.write(f"**Catégorie : {row['Question_Categorie']}**")
            st.write(row["question contenant le numéro unique"])

            if user_choice == correct_choice:
                st.markdown(
                    f"<div class='correct'>✅ Ta réponse : {user_choice} — Bonne réponse</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"<div class='wrong'>❌ Ta réponse : {user_choice}</div>",
                    unsafe_allow_html=True
                )
                st.markdown(
                    f"<div class='correct'>✅ Bonne réponse : {correct_choice}</div>",
                    unsafe_allow_html=True
                )

            if i in st.session_state.marked_for_review:
                st.write("**Marquée à revoir :** Oui")

            st.write(f"**Réponse correcte :** {correct_text}")
            st.write(f"**ID question :** {row['n°identifiant']}")
            st.write(f"**Thème :** {row['Theme']} | **Sous-thème :** {row['Sous_theme']}")
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Le corrigé est masqué pour simuler l’examen.")

else:
    st.info("Choisis un mode puis clique sur **Lancer le QCM**.")              