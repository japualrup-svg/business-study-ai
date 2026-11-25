import os
import csv
import json
import time
import base64
import random
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pypdf import PdfReader
from docx import Document
from pptx import Presentation

# ==========================================================
# ⚙️ إعدادات المبرمج (تُضبط هنا مرة واحدة ولا تظهر للمستخدمين)
# ==========================================================
LLM_PROVIDER = st.secrets.get("LLM_PROVIDER", "Groq")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

ENABLE_EMAIL = True
SENDER_EMAIL = st.secrets.get("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = st.secrets.get("SENDER_APP_PASSWORD", "")

# مدة كل جولة بالثواني (لكل 5 أسئلة)
STAGE_1_SECONDS = 90
STAGE_2_SECONDS = 100

# ==========================================================
# 1. إعدادات الصفحة والهوية البصرية (Agentic Mind × إرث الأعمال)
# ==========================================================
st.set_page_config(
    page_title="Agentic Mind | مقياس الجاهزية الذكي",
    page_icon="🧠",
    layout="centered"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "courses_data")
LOG_FILE = os.path.join(BASE_DIR, "analytics_log.csv")

def get_logo_base64():
    for ext in ["logo.png", "logo.jpg", "logo.jpeg", "logo.webp"]:
        logo_path = os.path.join(BASE_DIR, ext)
        if os.path.isfile(logo_path):
            with open(logo_path, "rb") as f:
                data = base64.b64encode(f.read()).decode("utf-8")
            mime = "image/png" if ext.endswith("png") else "image/jpeg"
            return f"data:{mime};base64,{data}"
    return None

LOGO_URI = get_logo_base64()

def get_course_theme(course_name):
    c = str(course_name)
    if any(k in c for k in ["مالية", "تمويل", "بنوك", "استثمار"]):
        return "💰", "#F59E0B"
    elif any(k in c for k in ["محاسبة", "تكاليف", "تدقيق"]):
        return "📒", "#10B981"
    elif any(k in c for k in ["إدارة", "ادارة", "موارد", "تنظيم"]):
        return "🧭", "#38BDF8"
    elif any(k in c for k in ["اقتصاد", "جزئي", "كلي"]):
        return "📈", "#A855F7"
    elif any(k in c for k in ["تسويق", "مبيعات"]):
        return "🎯", "#EC4899"
    return "📚", "#38BDF8"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap');

/* إخفاء شريط المطورين */
#MainMenu, footer, header, [data-testid="stSidebar"], [data-testid="collapsedControl"] {
    display: none !important;
    visibility: hidden !important;
}

html, body, .stApp, .main, .block-container, p, div, span, label, input, button {
    font-family: 'Cairo', sans-serif !important;
    direction: rtl;
    text-align: right;
}

/* خلفية شبكية هادئة منخفضة التوهج ومريحة للعين */
.stApp {
    background-color: #0B1120;
    background-image: 
        radial-gradient(circle at 15% 15%, rgba(56, 189, 248, 0.08) 0%, transparent 45%),
        radial-gradient(circle at 85% 80%, rgba(129, 140, 248, 0.08) 0%, transparent 45%),
        linear-gradient(rgba(148, 163, 184, 0.04) 1px, transparent 1px),
        linear-gradient(90deg, rgba(148, 163, 184, 0.04) 1px, transparent 1px);
    background-size: 100% 100%, 100% 100%, 36px 36px, 36px 36px;
}

.block-container {
    max-width: 900px;
    padding-top: 1.2rem;
    padding-bottom: 2.5rem;
}

@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(14px); }
    to   { opacity: 1; transform: translateY(0); }
}

@keyframes shimmerSweep {
    0%   { background-position: -200% center; }
    100% { background-position: 200% center; }
}

@keyframes pulseDot {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%      { opacity: 0.4; transform: scale(0.8); }
}

/* شريط القياس التقني لهوية الفريق */
.telemetry-bar {
    background: rgba(15, 23, 42, 0.85);
    border: 1px solid rgba(56, 189, 248, 0.25);
    border-radius: 12px;
    padding: 8px 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 12.5px;
    color: #94A3B8;
    margin-bottom: 12px;
}
.status-live {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    color: #38BDF8;
    font-weight: 800;
}
.live-dot {
    width: 8px;
    height: 8px;
    background: #10B981;
    border-radius: 50%;
    box-shadow: 0 0 8px #10B981;
    animation: pulseDot 1.8s infinite;
}

/* بطاقة الترويسة مع العلامة المائية */
.hero-card {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, rgba(23, 33, 54, 0.92) 0%, rgba(13, 20, 36, 0.96) 100%);
    backdrop-filter: blur(14px);
    border: 1px solid rgba(56, 189, 248, 0.28);
    border-radius: 22px;
    padding: 22px 28px;
    margin-bottom: 18px;
    box-shadow: 0 14px 32px rgba(0, 0, 0, 0.35), inset 0 1px 0 rgba(255, 255, 255, 0.06);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
}
.hero-watermark {
    position: absolute;
    left: -20px;
    bottom: -25px;
    width: 190px;
    opacity: 0.04;
    pointer-events: none;
    transform: rotate(-10deg);
}
.hero-content { flex: 1; z-index: 1; }
.hero-logo-box {
    width: 92px;
    height: 92px;
    border-radius: 20px;
    background: rgba(11, 17, 32, 0.9);
    border: 1.5px solid rgba(56, 189, 248, 0.45);
    padding: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 8px 22px rgba(56, 189, 248, 0.18);
    flex-shrink: 0;
    z-index: 1;
}
.hero-logo-box img {
    width: 100%;
    height: 100%;
    object-fit: contain;
    border-radius: 12px;
}
.hero-badge {
    display: inline-block;
    background: rgba(56, 189, 248, 0.14);
    border: 1px solid rgba(56, 189, 248, 0.35);
    color: #38BDF8;
    padding: 3px 14px;
    border-radius: 30px;
    font-size: 12.5px;
    font-weight: 800;
    margin-bottom: 8px;
}
.hero-title {
    color: #F8FAFC;
    font-size: 26px;
    font-weight: 900;
    margin-bottom: 5px;
}
.hero-desc {
    color: #CBD5E1;
    font-size: 14.5px;
    line-height: 1.65;
}

/* شريط الخطوات */
.stepper {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 18px;
}
.step-item {
    flex: 1;
    background: rgba(23, 33, 54, 0.6);
    border: 1px solid rgba(148, 163, 184, 0.16);
    border-radius: 12px;
    padding: 10px 6px;
    text-align: center;
    color: #64748B;
    font-size: 13.5px;
    font-weight: 700;
    transition: all 0.3s ease;
}
.step-active {
    background: linear-gradient(135deg, rgba(2, 132, 199, 0.25), rgba(56, 189, 248, 0.12));
    border-color: rgba(56, 189, 248, 0.6);
    color: #F8FAFC;
    box-shadow: 0 4px 14px rgba(56, 189, 248, 0.15);
}

/* بطاقة السؤال الواحد المركزية (Focus Card) */
.focus-q-card {
    background: linear-gradient(180deg, rgba(26, 37, 60, 0.95) 0%, rgba(17, 25, 42, 0.98) 100%);
    border: 1px solid rgba(56, 189, 248, 0.3);
    border-top: 4px solid #38BDF8;
    border-radius: 20px;
    padding: 24px 28px;
    margin-top: 10px;
    margin-bottom: 14px;
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.35);
}
.q-header-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 14px;
    flex-wrap: wrap;
    gap: 10px;
}
.q-pill-badge {
    background: rgba(56, 189, 248, 0.16);
    color: #38BDF8;
    border: 1px solid rgba(56, 189, 248, 0.35);
    padding: 4px 14px;
    border-radius: 20px;
    font-size: 13.5px;
    font-weight: 800;
}
.dots-tracker {
    display: flex;
    gap: 7px;
    align-items: center;
}
.dot-step {
    width: 28px;
    height: 8px;
    border-radius: 6px;
    background: #1E293B;
    border: 1px solid #334155;
    transition: all 0.3s ease;
}
.dot-done { background: #10B981; border-color: #10B981; }
.dot-curr { background: #38BDF8; border-color: #38BDF8; width: 38px; box-shadow: 0 0 10px rgba(56, 189, 248, 0.5); }

.q-main-text {
    color: #F1F5F9;
    font-size: 21px;
    font-weight: 800;
    line-height: 1.65;
}

/* بطاقات الإجابات ثابتة بدون رمشة عند النقر */
div[role="radiogroup"] {
    gap: 12px !important;
    width: 100%;
}
div[role="radiogroup"] > label {
    background: rgba(23, 33, 54, 0.78) !important;
    border: 1.5px solid rgba(148, 163, 184, 0.2) !important;
    border-radius: 15px !important;
    padding: 16px 22px !important;
    margin-bottom: 2px !important;
    cursor: pointer !important;
    transition: background-color 0.15s ease, border-color 0.15s ease, transform 0.15s ease !important;
    width: 100% !important;
    display: flex !important;
    align-items: center !important;
}
div[role="radiogroup"] > label:hover {
    background: rgba(56, 189, 248, 0.12) !important;
    border-color: #38BDF8 !important;
    transform: translateY(-2px);
    box-shadow: 0 6px 18px rgba(56, 189, 248, 0.15) !important;
}
div[role="radiogroup"] > label p {
    font-size: 19px !important;
    font-weight: 700 !important;
    color: #F1F5F9 !important;
    line-height: 1.5 !important;
    margin: 0 !important;
}
div[role="radiogroup"] > label:has(input:checked) {
    background: linear-gradient(90deg, rgba(2, 132, 199, 0.34), rgba(56, 189, 248, 0.18)) !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 6px 20px rgba(56, 189, 248, 0.28) !important;
}

/* بلوكات Bento Grid للتقرير النهائي */
.bento-card {
    background: linear-gradient(145deg, rgba(23, 33, 54, 0.9), rgba(13, 20, 36, 0.96));
    border: 1px solid rgba(148, 163, 184, 0.16);
    border-radius: 20px;
    padding: 22px 24px;
    margin-bottom: 16px;
    box-shadow: 0 10px 26px rgba(0, 0, 0, 0.28);
    animation: fadeInUp 0.5s ease-out;
}
.bento-green { border-top: 4px solid #10B981; }
.bento-amber { border-top: 4px solid #F59E0B; }
.bento-blue  { border-top: 4px solid #38BDF8; }

.bento-title {
    font-size: 18.5px;
    font-weight: 800;
    margin-bottom: 14px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.clean-row {
    background: rgba(11, 17, 32, 0.65);
    border: 1px solid rgba(148, 163, 184, 0.12);
    border-radius: 12px;
    padding: 12px 15px;
    margin-bottom: 9px;
    color: #E2E8F0;
    font-size: 15.5px;
    font-weight: 600;
    line-height: 1.6;
    display: flex;
    align-items: flex-start;
    gap: 10px;
}
.clean-row:last-child { margin-bottom: 0; }
.step-badge {
    background: rgba(56, 189, 248, 0.16);
    color: #38BDF8;
    border: 1px solid rgba(56, 189, 248, 0.35);
    border-radius: 8px;
    min-width: 26px;
    height: 26px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 13.5px;
    font-weight: 800;
    flex-shrink: 0;
}

/* شاشة الانتظار المخصصة بهوية الفريق */
.loader-box {
    background: linear-gradient(135deg, rgba(23, 33, 54, 0.95), rgba(11, 17, 32, 0.98));
    border: 1.5px solid rgba(56, 189, 248, 0.45);
    border-radius: 22px;
    padding: 36px 24px;
    text-align: center;
    margin: 18px 0;
    box-shadow: 0 14px 34px rgba(0, 0, 0, 0.4);
}
.pulse-ring {
    width: 82px;
    height: 82px;
    margin: 0 auto 16px auto;
    border-radius: 50%;
    border: 3px solid #38BDF8;
    display: flex;
    align-items: center;
    justify-content: center;
    animation: pulseGlow 1.6s infinite ease-in-out;
    background: rgba(11, 17, 32, 0.95);
}
.pulse-ring img { width: 56px; height: 56px; object-fit: contain; border-radius: 50%; }
@keyframes pulseGlow {
    0%   { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(56, 189, 248, 0.6); }
    70%  { transform: scale(1.04); box-shadow: 0 0 0 18px rgba(56, 189, 248, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(56, 189, 248, 0); }
}

/* حقول الإدخال والأزرار */
.stTextInput > div > div > input {
    font-size: 17px !important;
    padding: 14px 16px !important;
    border-radius: 13px !important;
    background-color: #131C31 !important;
    border: 1px solid rgba(148, 163, 184, 0.25) !important;
    color: #F8FAFC !important;
}
.stButton > button[kind="primary"], .stFormSubmitButton > button {
    height: 56px;
    font-size: 18px !important;
    font-weight: 800 !important;
    border-radius: 14px !important;
    background: linear-gradient(90deg, #0284C7, #38BDF8, #0284C7) !important;
    background-size: 200% auto !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 8px 22px rgba(2, 132, 199, 0.32) !important;
    animation: shimmerSweep 4s linear infinite;
    transition: transform 0.2s ease !important;
}
.stButton > button[kind="primary"]:hover, .stFormSubmitButton > button:hover {
    transform: translateY(-2px);
}
.stButton > button[kind="secondary"] {
    height: 56px;
    font-size: 16px !important;
    font-weight: 700 !important;
    border-radius: 14px !important;
    background: rgba(30, 41, 59, 0.75) !important;
    border: 1px solid rgba(148, 163, 184, 0.25) !important;
    color: #CBD5E1 !important;
}

/* شريط التوقيع السفلي */
.brand-footer {
    margin-top: 32px;
    padding: 14px 20px;
    border-top: 1px solid rgba(148, 163, 184, 0.14);
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #64748B;
    font-size: 13px;
    font-weight: 700;
}
/* إخفاء زر الإغلاق التلقائي للوقت بالكامل لمنع تداخل أيقونات Streamlit */
[data-testid="stExpander"] {
    position: fixed !important;
    bottom: -100px !important;
    opacity: 0 !important;
    pointer-events: none !important;
    height: 0 !important;
    overflow: hidden !important;
}

/* ضبط عرض خيارات الإجابة ومنع انكماشها */
div[role="radiogroup"] > label > div:last-child {
    width: 100% !important;
    flex: 1 !important;
}
</style>
""", unsafe_allow_html=True)

# ==========================================================
# 2. المكونات التفاعلية (العداد، الحلقة، البطاقات 3D)
# ==========================================================
def render_custom_loader(placeholder, title_text, sub_text):
    img_tag = f'<img src="{LOGO_URI}">' if LOGO_URI else '<span style="font-size:34px;">🧠</span>'
    placeholder.markdown(f"""
    <div class="loader-box">
        <div class="pulse-ring">{img_tag}</div>
        <div style="color:#F8FAFC; font-size:20px; font-weight:800; margin-bottom:6px;">{title_text}</div>
        <div style="color:#38BDF8; font-size:15px; font-weight:600;">{sub_text}</div>
    </div>
    """, unsafe_allow_html=True)

def render_gamified_timer(duration_seconds, round_title, timeout_button_match):
    components.html(f"""
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap" rel="stylesheet">
    <style>
        body {{ margin: 0; padding: 0; background: transparent; font-family: 'Cairo', sans-serif; overflow: hidden; }}
        .timer-box {{
            background: linear-gradient(135deg, rgba(15, 23, 42, 0.94), rgba(23, 33, 54, 0.94));
            border: 1.5px solid rgba(56, 189, 248, 0.4);
            border-radius: 16px;
            padding: 12px 20px;
            box-shadow: 0 8px 22px rgba(0, 0, 0, 0.28);
            transition: border-color 0.4s ease;
        }}
        .top-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }}
        .round-label {{ color: #F8FAFC; font-size: 15.5px; font-weight: 800; }}
        .vibe-msg {{ color: #38BDF8; font-size: 13.5px; font-weight: 700; }}
        .time-pill {{
            background: rgba(56, 189, 248, 0.16); color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.4);
            padding: 3px 15px; border-radius: 25px; font-size: 16.5px; font-weight: 900; min-width: 92px; text-align: center;
        }}
        .bar-bg {{ width: 100%; height: 9px; background: #0B1120; border-radius: 10px; overflow: hidden; border: 1px solid #334155; }}
        .bar-fill {{ height: 100%; width: 100%; background: linear-gradient(90deg, #38BDF8, #818CF8); border-radius: 10px; transition: width 0.25s linear, background 0.5s ease; }}
    </style>
    </head>
    <body>
    <div class="timer-box" id="tBox">
        <div class="top-row">
            <div>
                <div class="round-label">⚡ {round_title}</div>
                <div class="vibe-msg" id="vibeText">ركّز في السؤال الحالي.. الانتقال تلقائي وسريع!</div>
            </div>
            <div class="time-pill" id="timeDisplay">{duration_seconds} ثانية</div>
        </div>
        <div class="bar-bg"><div class="bar-fill" id="barFill"></div></div>
    </div>
    <script>
        const total = {duration_seconds};
        const startTime = Date.now();
        let triggered = false;

        const interval = setInterval(() => {{
            const elapsed = (Date.now() - startTime) / 1000;
            const remaining = Math.max(0, total - elapsed);
            const pct = (remaining / total) * 100;

            const disp = document.getElementById("timeDisplay");
            const bar = document.getElementById("barFill");
            const vibe = document.getElementById("vibeText");
            const box = document.getElementById("tBox");

            const mins = Math.floor(remaining / 60);
            const secs = Math.ceil(remaining % 60);
            disp.innerText = mins > 0 ? `${{mins}}:${{secs < 10 ? '0' : ''}}${{secs}}` : `${{Math.ceil(remaining)}} ثانية`;
            bar.style.width = pct + "%";

            if (remaining <= 15 && remaining > 0) {{
                bar.style.background = "linear-gradient(90deg, #F59E0B, #FBBF24)";
                disp.style.color = "#FBBF24";
                disp.style.borderColor = "rgba(251, 191, 36, 0.5)";
                disp.style.background = "rgba(251, 191, 36, 0.15)";
                box.style.borderColor = "rgba(251, 191, 36, 0.45)";
                vibe.style.color = "#FBBF24";
                vibe.innerText = "✨ اللمسات الأخيرة! حسم سريع قبل إغلاق الجولة";
            }} else if (remaining <= total * 0.5 && remaining > 15) {{
                bar.style.background = "linear-gradient(90deg, #818CF8, #C084FC)";
                vibe.innerText = "🚀 إيقاع ممتاز! استمر بنفس التركيز";
            }}

            if (remaining <= 0 && !triggered) {{
                triggered = true;
                clearInterval(interval);
                disp.innerText = "انتهى الوقت!";
                vibe.innerText = "⏱️ تم إغلاق الجولة تلقائياً.. جاري الانتقال!";
                const buttons = window.parent.document.querySelectorAll("button");
                for (let btn of buttons) {{
                    if (btn.innerText && btn.innerText.includes("{timeout_button_match}")) {{
                        btn.click();
                        break;
                    }}
                }}
            }}
        }}, 200);
    </script>
    </body>
    </html>
    """, height=88)

def render_animated_score_ring(score_pct, total_seconds):
    ring_color = "#10B981" if score_pct >= 80 else ("#38BDF8" if score_pct >= 50 else "#F59E0B")
    components.html(f"""
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap" rel="stylesheet">
    <style>
        body {{ margin: 0; background: transparent; font-family: 'Cairo', sans-serif; display: flex; justify-content: center; align-items: center; }}
        .gauge-wrap {{
            background: linear-gradient(145deg, rgba(23, 33, 54, 0.95), rgba(11, 17, 32, 0.98));
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 20px;
            padding: 16px 20px;
            text-align: center;
            width: 100%;
        }}
        svg {{ width: 122px; height: 122px; transform: rotate(-90deg); }}
        circle {{ fill: none; stroke-width: 10; stroke-linecap: round; }}
        .bg {{ stroke: #1E293B; }}
        .fg {{
            stroke: {ring_color};
            stroke-dasharray: 314;
            stroke-dashoffset: 314;
            transition: stroke-dashoffset 1.3s cubic-bezier(0.22, 1, 0.36, 1);
        }}
        .center-text {{
            position: relative;
            margin-top: -83px;
            margin-bottom: 35px;
            color: #F8FAFC;
            font-size: 27px;
            font-weight: 900;
        }}
        .time-badge {{
            display: inline-block;
            background: rgba(56, 189, 248, 0.14);
            color: #38BDF8;
            border: 1px solid rgba(56, 189, 248, 0.35);
            padding: 3px 14px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 800;
        }}
    </style>
    </head>
    <body>
    <div class="gauge-wrap">
        <svg viewBox="0 0 120 120">
            <circle class="bg" cx="60" cy="60" r="50"></circle>
            <circle class="fg" id="progCircle" cx="60" cy="60" r="50"></circle>
        </svg>
        <div class="center-text" id="pctText">0%</div>
        <div class="time-badge">⚡ الزمن: {total_seconds} ثانية</div>
    </div>
    <script>
        const target = {score_pct};
        const circle = document.getElementById("progCircle");
        const txt = document.getElementById("pctText");
        setTimeout(() => {{
            const offset = 314 - (314 * target) / 100;
            circle.style.strokeDashoffset = offset;
            let curr = 0;
            const step = Math.max(1, Math.ceil(target / 30));
            const counter = setInterval(() => {{
                curr = Math.min(target, curr + step);
                txt.innerText = curr + "%";
                if (curr >= target) clearInterval(counter);
            }}, 30);
        }}, 120);
    </script>
    </body>
    </html>
    """, height=198)

def render_3d_flip_flashcards(flashcards):
    cards_html = ""
    for i, c in enumerate(flashcards):
        cards_html += f"""
        <div class="flip-card" onclick="this.classList.toggle('flipped')">
            <div class="flip-inner">
                <div class="flip-front">
                    <div class="hint">بطاقة {i+1} • اضغط لقلب البطاقة وكشف الإجابة 🔄</div>
                    <div class="q-body">❓ {c['front']}</div>
                </div>
                <div class="flip-back">
                    <div class="hint-back">💡 الإجابة والقاعدة العلمية (اضغط للعودة)</div>
                    <div class="a-body">{c['back']}</div>
                </div>
            </div>
        </div>
        """
    height_px = len(flashcards) * 165 + 20
    components.html(f"""
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@600;700;800&display=swap" rel="stylesheet">
    <style>
        body {{ margin: 0; padding: 4px; background: transparent; font-family: 'Cairo', sans-serif; }}
        .flip-card {{ background-color: transparent; width: 100%; height: 145px; perspective: 1000px; margin-bottom: 14px; cursor: pointer; }}
        .flip-inner {{ position: relative; width: 100%; height: 100%; text-align: right; transition: transform 0.65s cubic-bezier(0.4, 0.2, 0.2, 1); transform-style: preserve-3d; }}
        .flip-card.flipped .flip-inner {{ transform: rotateY(180deg); }}
        .flip-front, .flip-back {{
            position: absolute; width: 100%; height: 100%; box-sizing: border-box;
            -webkit-backface-visibility: hidden; backface-visibility: hidden;
            border-radius: 16px; padding: 18px 22px; display: flex; flex-direction: column; justify-content: center;
        }}
        .flip-front {{ background: linear-gradient(145deg, #172136, #0B1120); border: 1.5px solid #334155; border-right: 5px solid #38BDF8; }}
        .flip-back  {{ background: linear-gradient(145deg, #0F292E, #0B1120); border: 1.5px solid #10B981; border-right: 5px solid #10B981; transform: rotateY(180deg); }}
        .hint {{ color: #94A3B8; font-size: 12.5px; font-weight: 700; margin-bottom: 6px; }}
        .hint-back {{ color: #34D399; font-size: 12.5px; font-weight: 700; margin-bottom: 6px; }}
        .q-body {{ color: #F8FAFC; font-size: 17.5px; font-weight: 800; line-height: 1.5; }}
        .a-body {{ color: #E2E8F0; font-size: 16px; font-weight: 600; line-height: 1.6; }}
    </style>
    </head>
    <body>{cards_html}</body>
    </html>
    """, height=height_px)

# ==========================================================
# 3. محرك المعرفة وسجل الإحصائيات الحية
# ==========================================================
def extract_text_from_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    text = ""
    try:
        if ext == ".pdf":
            reader = PdfReader(filepath)
            for page in reader.pages[:25]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        elif ext == ".docx":
            doc = Document(filepath)
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif ext == ".pptx":
            prs = Presentation(filepath)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text + "\n"
        elif ext == ".txt":
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
    except Exception:
        pass
    return text

@st.cache_data(show_spinner=False)
def load_all_courses_knowledge():
    courses_chunks = {}
    total_files = 0
    if not os.path.exists(DATA_DIR):
        return courses_chunks, 0

    for course_name in sorted(os.listdir(DATA_DIR)):
        course_path = os.path.join(DATA_DIR, course_name)
        if not os.path.isdir(course_path):
            continue

        full_text = ""
        for fname in os.listdir(course_path):
            fpath = os.path.join(course_path, fname)
            if os.path.isfile(fpath):
                extracted = extract_text_from_file(fpath)
                if extracted.strip():
                    full_text += "\n" + extracted
                    total_files += 1

        if full_text.strip():
            words = full_text.split()
            chunk_size = 140
            chunks = [" ".join(words[i:i + chunk_size]) for i in range(0, len(words), chunk_size)]
            courses_chunks[course_name] = chunks

    return courses_chunks, total_files

def sample_context(courses_chunks, target_courses=None):
    selected_courses = target_courses if target_courses else list(courses_chunks.keys())
    valid_courses = [c for c in selected_courses if c in courses_chunks]
    if not valid_courses:
        valid_courses = list(courses_chunks.keys())

    sampled_sections = []
    for course in valid_courses:
        chunks = courses_chunks[course]
        chosen = random.sample(chunks, min(2, len(chunks)))
        sampled_sections.append(f"=== مادة: {course} ===\n" + "\n".join(chosen))
    return "\n\n".join(sampled_sections)

def log_student_analytics(session_id, results_list):
    file_exists = os.path.isfile(LOG_FILE)
    with open(LOG_FILE, mode="a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp", "session_id", "course", "topic", "is_correct", "stage"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for item in results_list:
            writer.writerow([
                now,
                session_id,
                item.get("course", "غير محدد"),
                item.get("topic", "عام"),
                1 if item["is_correct"] else 0,
                item.get("stage", 1)
            ])

# ==========================================================
# 4. محرك الذكاء الاصطناعي التكيفي (باللغة العربية بالكامل)
# ==========================================================
def call_llm(system_prompt, user_prompt):
    if LLM_PROVIDER.lower() == "groq":
        from groq import Groq
        client = Groq(api_key=GROQ_API_KEY)
        all_models = [m.id for m in client.models.list().data]
        chat_models = [
            m for m in all_models
            if any(k in m.lower() for k in ["llama", "qwen", "gemma", "gpt-oss", "mixtral"])
            and not any(bad in m.lower() for bad in ["guard", "vision", "whisper", "tts", "orpheus"])
        ]
        priority = ["llama-3.3-70b-versatile", "openai/gpt-oss-120b", "qwen/qwen3-32b", "llama-3.1-8b-instant"]
        ordered_models = [p for p in priority if p in chat_models] + [m for m in chat_models if m not in priority]

        last_err = None
        for model_name in ordered_models:
            try:
                resp = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.35
                )
                return json.loads(resp.choices[0].message.content)
            except Exception as e:
                last_err = e
                continue
        raise RuntimeError(f"تعذر الاتصال بنموذج Groq: {last_err}")
    else:
        from google import genai
        client = genai.Client(api_key=GEMINI_API_KEY)
        resp = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=f"{system_prompt}\n\n{user_prompt}",
            config={"response_mime_type": "application/json"}
        )
        return json.loads(resp.text)

def generate_stage_1_questions(context_text):
    sys_prompt = """
    أنت لجنة أكاديمية في كلية الأعمال. صمم المرحلة الأولى من تحدي الذكاء التكيفي:
    5 أسئلة اختيار من متعدد (MCQ) ممتعة وواضحة تقيس المفاهيم الأساسية والتطبيق المباشر موزعة على المواد المرفقة.
    قاعدة لغوية صارمة: يجب أن تكون جميع النصوص (اسم الموضوع، نص السؤال، الخيارات الأربعة، والتوضيح العلمي) مكتوبة باللغة العربية بالكامل، حتى لو كان المحتوى الدراسي المرفق باللغة الإنجليزية.
    أخرج JSON فقط بهذا الهيكل:
    {
        "questions": [
            {
                "course": "اسم المادة كما وردت في النص",
                "topic": "اسم الموضوع أو الدرس بالعربية",
                "question": "نص السؤال باللغة العربية بوضوح",
                "options": ["الخيار الأول بالعربية", "الخيار الثاني بالعربية", "الخيار الثالث بالعربية", "الخيار الرابع بالعربية"],
                "correct_index": 0,
                "explanation": "توضيح علمي مختصر باللغة العربية"
            }
        ]
    }
    """
    return call_llm(sys_prompt, f"المحتوى الدراسي:\n\n{context_text}")

def generate_stage_2_adaptive(context_text, stage1_results, mode):
    if mode == "advanced":
        instruction = """
        أثبت الطالب تمكناً عالياً في الجولة الأولى!
        صمم له الآن 5 أسئلة تحدي متقدمة (حالات عملية واتخاذ قرارات إدارية أو مالية أو محاسبية) بناءً على المحتوى المرفق.
        """
    else:
        weak_topics = [f"{r['course']} ({r['topic']})" for r in stage1_results if not r["is_correct"]]
        instruction = f"""
        أظهر الطالب حاجة لتعزيز فهمه في الجولة الأولى وتحديداً في: {', '.join(weak_topics)}.
        صمم له الآن 5 أسئلة تشخيصية مركزة ومبسطة على نفس هذه المواد والمفاهيم لتحديد أين يكمن الخلل في فهمه بدقة.
        """

    sys_prompt = f"""
    أنت محرك تحدي تكيفي ذكي في كلية الأعمال.
    {instruction}
    قاعدة لغوية صارمة: يجب أن تكون جميع الأسئلة، الخيارات الأربعة، والتوضيحات مكتوبة باللغة العربية بالكامل.
    أخرج JSON فقط بهذا الهيكل:
    {{
        "questions": [
            {{
                "course": "اسم المادة",
                "topic": "اسم الموضوع بالعربية",
                "question": "نص السؤال التكيفي باللغة العربية",
                "options": ["الخيار الأول بالعربية", "الخيار الثاني بالعربية", "الخيار الثالث بالعربية", "الخيار الرابع بالعربية"],
                "correct_index": 0,
                "explanation": "توضيح علمي مختصر باللغة العربية"
            }}
        ]
    }}
    """
    return call_llm(sys_prompt, f"المحتوى الدراسي المرجعي:\n\n{context_text}")

def analyze_full_adaptive_exam(student_name, all_results, adaptive_mode, total_seconds):
    mode_desc = "مسار التحدي المتقدم (حالات عملية)" if adaptive_mode == "advanced" else "مسار التشخيص العلاجي المركز"
    sys_prompt = f"""
    أنت مستشار أكاديمي ذكي في كلية الأعمال. حلل أداء الطالب في التحدي التكيفي (10 أسئلة على جولتين - المسار الذي وصل إليه: {mode_desc} - الزمن المستغرق: {total_seconds} ثانية).
    قواعد صارمة:
    1. اكتب جميع العناصر باللغة العربية بالكامل بأسلوب مشجع وواضح.
    2. لا تسرد كل سؤال كسطر منفصل؛ اجمع نقاط القوة في (2 إلى 3 نقاط مركزة كحد أقصى)، واجمع المواضيع التي تحتاج لتقوية في (3 نقاط رئيسية مركزة كحد أقصى) مع ذكر اسم المادة باختصار.
    3. استنتج من إجاباته التخصص الأنسب له في كلية الأعمال ومسمى وظيفي مقترح.
    أخرج JSON فقط بهذا الهيكل:
    {{
        "score_percentage": 80,
        "level_title": "التصنيف الأكاديمي للطالب بالعربية",
        "overall_summary": "ملخص دقيق ومشجع بالعربية لمستوى الطالب وسرعته وكيف تطور أداؤه بين الجولتين",
        "career_compass": "توصية ذكية بالتخصص الدقيق الأنسب لطالب الأعمال بناءً على إجاباته (مثل المالية، المحاسبة، الإدارة، أو التسويق) مع مسمى وظيفي مقترح",
        "strengths": ["محور قوة 1 مختصر مع اسم المادة", "محور قوة 2 مختصر"],
        "weaknesses": ["محور تطوير 1 مختصر مع اسم المادة", "محور تطوير 2 مختصر", "محور تطوير 3 مختصر"],
        "advice_and_plan": ["خطوة عملية 1 بالعربية", "خطوة عملية 2 بالعربية", "خطوة عملية 3 بالعربية"],
        "flashcards": [
            {{"front": "مفهوم أو سؤال مراجعة 1 بالعربية", "back": "الشرح المبسط والقاعدة 1 بالعربية"}},
            {{"front": "مفهوم أو سؤال مراجعة 2 بالعربية", "back": "الشرح المبسط والقاعدة 2 بالعربية"}},
            {{"front": "مفهوم أو سؤال مراجعة 3 بالعربية", "back": "الشرح المبسط والقاعدة 3 بالعربية"}}
        ]
    }}
    """
    user_prompt = f"اسم الطالب: {student_name}\nسجل الإجابات العشرة:\n{json.dumps(all_results, ensure_ascii=False)}"
    return call_llm(sys_prompt, user_prompt)

# ==========================================================
# 5. محرك إرسال التقرير على الإيميل
# ==========================================================
def send_email_report(to_email, student_name, report, mode_label, total_seconds):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🎓 بطاقة إنجازك وتقييمك الذكي في كلية الأعمال — {student_name}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email

    strengths_html = "".join([f"<li style='margin-bottom:6px;'>{s}</li>" for s in report["strengths"]])
    weaknesses_html = "".join([f"<li style='margin-bottom:6px;'>{w}</li>" for w in report["weaknesses"]])
    plan_html = "".join([f"<li style='margin-bottom:6px;'>{p}</li>" for p in report["advice_and_plan"]])
    cards_html = "".join([
        f"<div style='background:#f8fafc;border-right:4px solid #0284c7;padding:12px;margin-bottom:10px;border-radius:6px;'>"
        f"<b style='color:#0369a1;'>❓ {c['front']}</b><br><span style='color:#334155;'>💡 {c['back']}</span></div>"
        for c in report["flashcards"]
    ])

    html_body = f"""
    <html dir="rtl">
    <body style="font-family: Tahoma, Arial, sans-serif; line-height: 1.8; color: #1e293b; background-color: #f1f5f9; padding: 20px;">
        <div style="max-width: 650px; margin: auto; background: #ffffff; padding: 30px; border-radius: 14px; border: 1px solid #cbd5e1;">
            <h2 style="color: #0284c7; margin-top: 0;">أهلاً بك {student_name} 👋</h2>
            <p>إليك نتيجة مشاركتك في تحدي الجاهزية الذكي لكلية الأعمال (<b>{mode_label}</b> — الزمن: <b>{total_seconds} ثانية</b>):</p>
            <div style="background: #f0f9ff; border: 1px solid #bae6fd; padding: 18px; border-radius: 10px; text-align: center; margin: 20px 0;">
                <div style="font-size: 26px; font-weight: bold; color: #0369a1;">النتيجة الكلية: {report['score_percentage']}%</div>
                <div style="font-size: 16px; font-weight: bold; color: #0f172a; margin-top: 4px;">{report['level_title']}</div>
                <p style="font-size: 14px; color: #475569; margin-bottom: 0;">{report['overall_summary']}</p>
            </div>
            <div style="background: #faf5ff; border-right: 4px solid #9333ea; padding: 12px 16px; border-radius: 8px; margin-bottom: 18px;">
                <b style="color: #7e22ce;">🧭 بوصلة التخصص والمسار المهني:</b><br>
                <span>{report.get('career_compass', '')}</span>
            </div>
            <h3 style="color: #15803d;">✅ أبرز نقاط القوة لديك:</h3>
            <ul>{strengths_html}</ul>
            <h3 style="color: #b91c1c;">🎯 محاور يُنصح بتقويتها:</h3>
            <ul>{weaknesses_html}</ul>
            <h3 style="color: #0369a1;">🚀 خطتك المقترحة للتطوير:</h3>
            <ol>{plan_html}</ol>
            <h3 style="color: #0f172a;">🗂️ بطاقات المراجعة السريعة (Flashcards):</h3>
            {cards_html}
            <hr style="border: none; border-top: 1px solid #e2e8f0; margin-top: 25px;">
            <p style="font-size: 12px; color: #64748b; text-align: center;">تطوير فريق Agentic Mind × إرث الأعمال</p>
        </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())

# ==========================================================
# 6. إدارة الجلسة والتمرير التلقائي لأعلى الشاشة
# ==========================================================
defaults = {
    "step": 1,
    "show_dashboard": False,
    "quiz_s1": None,
    "q_idx_s1": 0,
    "ans_s1": [None] * 5,
    "results_s1": None,
    "quiz_s2": None,
    "q_idx_s2": 0,
    "ans_s2": [None] * 5,
    "adaptive_mode": None,
    "report": None,
    "all_review": None,
    "email_sent": False,
    "stage_start_time": 0.0,
    "time_s1": 0,
    "time_s2": 0,
    "celebrated": False
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_all():
    for k, v in defaults.items():
        st.session_state[k] = v
    st.session_state.ans_s1 = [None] * 5
    st.session_state.ans_s2 = [None] * 5

# علامة مرجعية أعلى الصفحة لضمان بداية كل شاشة وسؤال من الأعلى
st.markdown('<div id="top-anchor"></div>', unsafe_allow_html=True)

current_view_state = f"{st.session_state.step}_{st.session_state.show_dashboard}_{st.session_state.q_idx_s1}_{st.session_state.q_idx_s2}"
if st.session_state.get("last_view_state") != current_view_state:
    st.session_state.last_view_state = current_view_state
    components.html(
        f"""
        <script>
            function forceScrollTop() {{
                const anchor = window.parent.document.getElementById('top-anchor');
                if (anchor) anchor.scrollIntoView({{ behavior: 'instant', block: 'start' }});
                window.parent.scrollTo(0, 0);
                const containers = window.parent.document.querySelectorAll(
                    '[data-testid="stAppViewContainer"], [data-testid="stMain"], section.main, .main'
                );
                containers.forEach(c => {{ c.scrollTop = 0; }});
            }}
            forceScrollTop();
            setTimeout(forceScrollTop, 120);
        </script>
        <!-- {time.time()} -->
        """,
        height=0,
    )

# شريط القياس التقني لهوية الفريق + زر لوحة الدكاترة
t_col1, t_col2 = st.columns([3, 1.35])
with t_col1:
    st.markdown("""
    <div class="telemetry-bar">
        <span class="status-live"><span class="live-dot"></span> محرك Agentic Mind التكيفي نشط</span>
        <span>الوضع: تشخيص ذكي • RTL</span>
    </div>
    """, unsafe_allow_html=True)
with t_col2:
    dash_label = "🔙 العودة للتحدي" if st.session_state.show_dashboard else "📊 مؤشرات الكلية"
    if st.button(dash_label, use_container_width=True, key="btn_toggle_dash"):
        st.session_state.show_dashboard = not st.session_state.show_dashboard
        st.rerun()

logo_html = f'<div class="hero-logo-box"><img src="{LOGO_URI}" alt="Agentic Mind Logo"></div>' if LOGO_URI else ""
watermark_html = f'<img src="{LOGO_URI}" class="hero-watermark">' if LOGO_URI else ""

# ----------------------------------------------------------
# لوحة المؤشرات الحية للدكاترة والعمادة
# ----------------------------------------------------------
if st.session_state.show_dashboard:
    st.markdown(f"""
    <div class="hero-card">
        {watermark_html}
        <div class="hero-content">
            <div class="hero-badge">Agentic Mind Analytics • لوحة البيانات الحية</div>
            <div class="hero-title">📊 مؤشرات أداء طلبة كلية الأعمال اليوم</div>
            <div class="hero-desc">تحليل لحظي ومجهول الهوية لنتائج الطلبة المشاركين في الفعالية لتحديد المواد والمفاهيم الأكثر تحدياً.</div>
        </div>
        {logo_html}
    </div>
    """, unsafe_allow_html=True)

    if not os.path.isfile(LOG_FILE):
        st.info("لا توجد بيانات مسجلة بعد. ستظهر الإحصائيات والرسوم البيانية تلقائياً بمجرد إتمام أول طالب للتحدي.")
    else:
        df = pd.read_csv(LOG_FILE)
        total_students = df["session_id"].nunique()
        total_questions = len(df)
        avg_accuracy = round(df["is_correct"].mean() * 100, 1)

        m1, m2, m3 = st.columns(3)
        m1.metric("👥 عدد الطلبة المشاركين", f"{total_students} طالب")
        m2.metric("📝 إجمالي الأسئلة المجابة", f"{total_questions} سؤال")
        m3.metric("🎯 متوسط الاستيعاب العام", f"{avg_accuracy}%")

        st.markdown('<div class="bento-card bento-blue"><div class="bento-title" style="color:#38BDF8;">📈 نسبة الإتقان حسب كل مادة دراسية (%)</div>', unsafe_allow_html=True)
        course_stats = df.groupby("course")["is_correct"].mean().mul(100).round(1).sort_values(ascending=False)
        for c_name, pct_val in course_stats.items():
            icon, hex_col = get_course_theme(c_name)
            st.markdown(f"""
            <div style="margin-bottom:12px;">
                <div style="display:flex; justify-content:space-between; color:#F8FAFC; font-weight:700; margin-bottom:4px;">
                    <span>{icon} {c_name}</span>
                    <span style="color:{hex_col};">{pct_val}%</span>
                </div>
                <div style="width:100%; height:11px; background:#0B1120; border-radius:8px; overflow:hidden; border:1px solid #334155;">
                    <div style="width:{pct_val}%; height:100%; background:{hex_col}; border-radius:8px;"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="bento-card bento-amber"><div class="bento-title" style="color:#FBBF24;">⚠️ أكثر 5 مواضيع/دروس واجه فيها الطلبة صعوبة</div></div>', unsafe_allow_html=True)
        topic_stats = (
            df.groupby(["course", "topic"])
            .agg(عدد_المحاولات=("is_correct", "count"), نسبة_الإتقان=("is_correct", lambda x: round(x.mean() * 100, 1)))
            .reset_index()
            .sort_values(by="نسبة_الإتقان", ascending=True)
            .head(5)
        )
        topic_stats.columns = ["المادة", "الموضوع / الدرس", "عدد الأسئلة", "نسبة الإتقان (%)"]
        st.dataframe(topic_stats, use_container_width=True, hide_index=True)

    st.stop()

# ----------------------------------------------------------
# الترويسة وشريط المراحل
# ----------------------------------------------------------
s1 = "step-item step-active" if st.session_state.step == 1 else "step-item"
s2 = "step-item step-active" if st.session_state.step == 2 else "step-item"
s3 = "step-item step-active" if st.session_state.step == 3 else "step-item"
s4 = "step-item step-active" if st.session_state.step == 4 else "step-item"

st.markdown(f"""
<div class="hero-card">
    {watermark_html}
    <div class="hero-content">
        <div class="hero-badge">Agentic Mind × إرث الأعمال</div>
        <div class="hero-title">🎓 تحدي الجاهزية الذكي لكلية الأعمال</div>
        <div class="hero-desc">
            تحدٍّ تفاعلي سريع من جولتين (10 أسئلة) يتكيف مع طريقة تفكيرك لحظياً، ويكشف نقاط قوتك وتخصصك الأنسب مع خطة تطوير تصلك على إيميلك!
        </div>
    </div>
    {logo_html}
</div>
<div class="stepper">
    <div class="{s1}">1️⃣ اللاعب</div>
    <div class="{s2}">2️⃣ الجولة الأولى</div>
    <div class="{s3}">3️⃣ الجولة التكيفية</div>
    <div class="{s4}">4️⃣ النتيجة والخطة</div>
</div>
""", unsafe_allow_html=True)

courses_chunks, total_files = load_all_courses_knowledge()

# ----------------------------------------------------------
# دوال إنهاء الجولات
# ----------------------------------------------------------
def complete_stage_1(total_elapsed):
    res_s1 = []
    correct_count = 0
    for idx, q in enumerate(st.session_state.quiz_s1):
        c_text = q["options"][q["correct_index"]]
        user_ans = st.session_state.ans_s1[idx]
        chosen = user_ans if user_ans is not None else "تم تخطي السؤال (انتهى وقت الجولة)"
        is_ok = (user_ans == c_text)
        if is_ok:
            correct_count += 1
        res_s1.append({
            "stage": 1,
            "course": q.get("course", ""),
            "topic": q.get("topic", ""),
            "question": q["question"],
            "student_answer": chosen,
            "correct_answer": c_text,
            "is_correct": is_ok,
            "explanation": q["explanation"]
        })

    mode = "advanced" if correct_count >= 4 else "diagnostic"
    weak_courses = list({r["course"] for r in res_s1 if not r["is_correct"]})

    loader_slot = st.empty()
    render_custom_loader(
        loader_slot,
        "🤖 محرك Agentic Mind يحلل إجاباتك ويضبط الجولة الثانية...",
        "يتم الآن بناء 5 أسئلة تكيفية مخصصة لمستواك"
    )
    try:
        ctx_s2 = sample_context(courses_chunks, target_courses=None if mode == "advanced" else weak_courses)
        s2_data = generate_stage_2_adaptive(ctx_s2, res_s1, mode)
        st.session_state.time_s1 = min(total_elapsed, STAGE_1_SECONDS)
        st.session_state.results_s1 = res_s1
        st.session_state.adaptive_mode = mode
        st.session_state.quiz_s2 = s2_data["questions"][:5]
        st.session_state.q_idx_s2 = 0
        st.session_state.ans_s2 = [None] * 5
        st.session_state.stage_start_time = time.time()
        st.session_state.step = 3
        st.rerun()
    except Exception as e:
        loader_slot.empty()
        st.error(f"حدث خطأ أثناء توليد الجولة التكيفية: {e}")

def complete_stage_2(total_elapsed_2):
    res_s2 = []
    for idx, q in enumerate(st.session_state.quiz_s2):
        c_text = q["options"][q["correct_index"]]
        user_ans = st.session_state.ans_s2[idx]
        chosen = user_ans if user_ans is not None else "تم تخطي السؤال (انتهى وقت الجولة)"
        res_s2.append({
            "stage": 2,
            "course": q.get("course", ""),
            "topic": q.get("topic", ""),
            "question": q["question"],
            "student_answer": chosen,
            "correct_answer": c_text,
            "is_correct": (user_ans == c_text),
            "explanation": q["explanation"]
        })

    st.session_state.time_s2 = min(total_elapsed_2, STAGE_2_SECONDS)
    total_time_taken = st.session_state.time_s1 + st.session_state.time_s2

    all_results = st.session_state.results_s1 + res_s2
    session_id = f"{datetime.now().strftime('%H%M%S')}_{random.randint(100,999)}"
    log_student_analytics(session_id, all_results)

    loader_slot = st.empty()
    render_custom_loader(
        loader_slot,
        "🔍 جاري بناء تقريرك الشخصي وبوصلة تخصصك...",
        "نرسل نسخة من بطاقة إنجازك إلى بريدك الإلكتروني الآن"
    )
    try:
        report = analyze_full_adaptive_exam(
            st.session_state.student_name,
            all_results,
            st.session_state.adaptive_mode,
            total_time_taken
        )
        st.session_state.all_review = all_results
        st.session_state.report = report

        mode_label = "جولة المحترفين" if st.session_state.adaptive_mode == "advanced" else "التحدي الذكي المخصص"
        if ENABLE_EMAIL and "your_email" not in SENDER_EMAIL:
            try:
                send_email_report(
                    st.session_state.student_email,
                    st.session_state.student_name,
                    report,
                    mode_label,
                    total_time_taken
                )
                st.session_state.email_sent = True
            except Exception:
                st.session_state.email_sent = False

        st.session_state.step = 4
        st.rerun()
    except Exception as e:
        loader_slot.empty()
        st.error(f"حدث خطأ أثناء التحليل: {e}")

# ----------------------------------------------------------
# الخطوة 1: بيانات الطالب
# ----------------------------------------------------------
if st.session_state.step == 1:
    with st.form("login_form", clear_on_submit=False):
        st.markdown("### 👋 جاهز للتحدي؟ أدخل بياناتك وابدأ الجولة!")
        student_name = st.text_input("الاسم الكامل:", placeholder="اكتب اسمك هنا...")
        student_email = st.text_input("البريد الإلكتروني (لاستلام بطاقة الإنجاز وملخص المراجعة):", placeholder="name@example.com")
        st.write("")
        submitted = st.form_submit_button("🚀 انطلق في الجولة الأولى (90 ثانية)", type="primary", use_container_width=True)

    if submitted:
        if not student_name.strip():
            st.warning("يرجى إدخال اسمك للبدء.")
        elif not student_email.strip() or "@" not in student_email:
            st.warning("يرجى إدخال بريد إلكتروني صحيح.")
        elif total_files == 0:
            st.error("⚠️ لم يتم العثور على ملفات مواد داخل مجلد courses_data.")
        else:
            loader_slot = st.empty()
            render_custom_loader(
                loader_slot,
                "🧠 الذكاء الاصطناعي يقرأ مناهج كلية الأعمال...",
                "جاري صياغة 5 تحديات ذكية مخصصة لك الآن"
            )
            try:
                ctx = sample_context(courses_chunks)
                s1_data = generate_stage_1_questions(ctx)
                st.session_state.student_name = student_name.strip()
                st.session_state.student_email = student_email.strip()
                st.session_state.quiz_s1 = s1_data["questions"][:5]
                st.session_state.q_idx_s1 = 0
                st.session_state.ans_s1 = [None] * 5
                st.session_state.stage_start_time = time.time()
                st.session_state.step = 2
                st.rerun()
            except Exception as e:
                loader_slot.empty()
                st.error(f"حدث خطأ تقني: {e}")

# ----------------------------------------------------------
# الخطوة 2: الجولة الأولى (سؤال واحد بدون رمشة عند الاختيار)
# ----------------------------------------------------------
@st.fragment
def render_stage_1_question_box():
    idx = st.session_state.q_idx_s1
    q = st.session_state.quiz_s1[idx]
    saved_val = st.session_state.ans_s1[idx]
    saved_index = q["options"].index(saved_val) if saved_val in q["options"] else None

    dots_html = "".join([
        f'<span class="dot-step {"dot-curr" if i == idx else ("dot-done" if st.session_state.ans_s1[i] is not None else "")}"></span>'
        for i in range(5)
    ])

    st.markdown(f"""
    <div class="focus-q-card">
        <div class="q-header-row">
            <span class="q-pill-badge">⚡ تحدي {idx + 1} من 5</span>
            <div class="dots-tracker">{dots_html}</div>
        </div>
        <div class="q-main-text">{q['question']}</div>
    </div>
    """, unsafe_allow_html=True)

    chosen_opt = st.radio(
        f"s1_single_{idx}",
        q["options"],
        index=saved_index,
        key=f"radio_s1_{idx}",
        label_visibility="collapsed"
    )
    st.session_state.ans_s1[idx] = chosen_opt

    st.write("")
    nav1, nav2 = st.columns([1, 2.2])
    with nav1:
        if idx > 0:
            if st.button("➡️ السابق", type="secondary", use_container_width=True, key=f"prev_s1_{idx}"):
                st.session_state.q_idx_s1 -= 1
                st.rerun(scope="fragment")
    with nav2:
        if idx < 4:
            if st.button("السؤال التالي ⬅️", type="primary", use_container_width=True, key=f"next_s1_{idx}"):
                if chosen_opt is None:
                    st.warning("✨ اختر إجابة أولاً للانتقال للسؤال التالي.")
                else:
                    st.session_state.q_idx_s1 += 1
                    st.rerun(scope="fragment")
        else:
            if st.button("⚡ تحليل أدائي والانتقال للجولة التكيفية", type="primary", use_container_width=True, key="finish_s1_btn"):
                total_elapsed = int(time.time() - st.session_state.stage_start_time)
                if chosen_opt is None and total_elapsed < (STAGE_1_SECONDS - 1):
                    st.warning("✨ اختر إجابة هذا السؤال قبل الانتقال للجولة الثانية.")
                else:
                    complete_stage_1(total_elapsed)

if st.session_state.step == 2 and st.session_state.quiz_s1:
    elapsed_now = int(time.time() - st.session_state.stage_start_time)
    remaining_s1 = max(1, STAGE_1_SECONDS - elapsed_now)

    render_gamified_timer(
        duration_seconds=remaining_s1,
        round_title=f"الجولة الأولى: تحدي الأساسيات • {st.session_state.student_name}",
        timeout_button_match="إغلاق الجولة الأولى"
    )

    render_stage_1_question_box()

    with st.expander("خيارات الجولة", expanded=False):
        if st.button("⏱️ إغلاق الجولة الأولى والانتقال", key="timeout_s1_trigger"):
            total_elapsed = int(time.time() - st.session_state.stage_start_time)
            complete_stage_1(total_elapsed)

# ----------------------------------------------------------
# الخطوة 3: الجولة الثانية التكيفية (سؤال واحد بدون رمشة عند الاختيار)
# ----------------------------------------------------------
@st.fragment
def render_stage_2_question_box():
    idx2 = st.session_state.q_idx_s2
    q2 = st.session_state.quiz_s2[idx2]
    saved_val2 = st.session_state.ans_s2[idx2]
    saved_index2 = q2["options"].index(saved_val2) if saved_val2 in q2["options"] else None

    dots_html2 = "".join([
        f'<span class="dot-step {"dot-curr" if i == idx2 else ("dot-done" if st.session_state.ans_s2[i] is not None else "")}"></span>'
        for i in range(5)
    ])

    st.markdown(f"""
    <div class="focus-q-card" style="border-top-color: #818CF8;">
        <div class="q-header-row">
            <span class="q-pill-badge" style="background:rgba(129, 140, 248, 0.16); color:#A5B4FC; border-color:rgba(129, 140, 248, 0.4);">
                🎯 تحدي تكيفي {idx2 + 1} من 5
            </span>
            <div class="dots-tracker">{dots_html2}</div>
        </div>
        <div class="q-main-text">{q2['question']}</div>
    </div>
    """, unsafe_allow_html=True)

    chosen_opt2 = st.radio(
        f"s2_single_{idx2}",
        q2["options"],
        index=saved_index2,
        key=f"radio_s2_{idx2}",
        label_visibility="collapsed"
    )
    st.session_state.ans_s2[idx2] = chosen_opt2

    st.write("")
    nav2_1, nav2_2 = st.columns([1, 2.2])
    with nav2_1:
        if idx2 > 0:
            if st.button("➡️ السابق", type="secondary", use_container_width=True, key=f"prev_s2_{idx2}"):
                st.session_state.q_idx_s2 -= 1
                st.rerun(scope="fragment")
    with nav2_2:
        if idx2 < 4:
            if st.button("السؤال التالي ⬅️", type="primary", use_container_width=True, key=f"next_s2_{idx2}"):
                if chosen_opt2 is None:
                    st.warning("✨ اختر إجابة أولاً للانتقال للسؤال التالي.")
                else:
                    st.session_state.q_idx_s2 += 1
                    st.rerun(scope="fragment")
        else:
            if st.button("📊 إصدار بطاقة الإنجاز والتقرير لإيميلي", type="primary", use_container_width=True, key="finish_s2_btn"):
                total_elapsed_2 = int(time.time() - st.session_state.stage_start_time)
                if chosen_opt2 is None and total_elapsed_2 < (STAGE_2_SECONDS - 1):
                    st.warning("✨ اختر إجابة هذا السؤال قبل إصدار التقرير.")
                else:
                    complete_stage_2(total_elapsed_2)

if st.session_state.step == 3 and st.session_state.quiz_s2:
    elapsed_now_2 = int(time.time() - st.session_state.stage_start_time)

    remaining_s2 = max(1, STAGE_2_SECONDS - elapsed_now_2)

    if st.session_state.adaptive_mode == "advanced":
        round_lbl = "الجولة الثانية: تحدي المحترفين (حالات عملية)"
    else:
        round_lbl = "الجولة الثانية: التحدي الذكي المخصص"

    render_gamified_timer(
        duration_seconds=remaining_s2,
        round_title=round_lbl,
        timeout_button_match="إغلاق الجولة الثانية"
    )

    render_stage_2_question_box()

    with st.expander("خيارات الجولة", expanded=False):
        if st.button("⏱️ إغلاق الجولة الثانية وإصدار التقرير", key="timeout_s2_trigger"):
            total_elapsed_2 = int(time.time() - st.session_state.stage_start_time)
            complete_stage_2(total_elapsed_2)

# ----------------------------------------------------------
# الخطوة 4: التقرير النهائي بنظام البلوكات (Bento Grid)
# ----------------------------------------------------------
elif st.session_state.step == 4 and st.session_state.report:
    rep = st.session_state.report
    total_time_taken = st.session_state.time_s1 + st.session_state.time_s2
    score_val = int(rep.get("score_percentage", 0))

    if score_val >= 80 and not st.session_state.celebrated:
        st.balloons()
        st.session_state.celebrated = True

    if st.session_state.email_sent:
        st.success(f"📩 تم إرسال بطاقة إنجازك وبوصلة تخصصك إلى: **{st.session_state.student_email}**")

    # صف Bento العلوي: الملخص + المؤشر الدائري
    b1, b2 = st.columns([2.2, 1])
    with b1:
        st.markdown(f"""
        <div class="bento-card bento-blue" style="height:100%;">
            <div class="bento-title" style="color:#38BDF8; font-size:22px;">🎯 {rep['level_title']}</div>
            <p style="color:#E2E8F0; font-size:15.5px; line-height:1.75; margin-bottom:12px;">{rep['overall_summary']}</p>
            <div style="background:rgba(168, 85, 247, 0.12); border:1px solid rgba(168, 85, 247, 0.35); border-radius:12px; padding:11px 15px;">
                <span style="color:#C084FC; font-weight:800;">🧭 بوصلة التخصص والمسار: </span>
                <span style="color:#F1F5F9; font-size:14.5px; font-weight:600;">{rep.get('career_compass', '')}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with b2:
        render_animated_score_ring(score_val, total_time_taken)

    # صف Bento الأوسط: نقاط القوة ومحاور التطوير
    col_s, col_w = st.columns(2)
    with col_s:
        s_items = "".join([
            f'<div class="clean-row"><span style="color:#10B981; font-weight:900;">✔</span><span>{x}</span></div>'
            for x in rep["strengths"][:4]
        ])
        st.markdown(f"""
        <div class="bento-card bento-green">
            <div class="bento-title" style="color:#34D399;">🌟 أبرز نقاط القوة لديك</div>
            {s_items}
        </div>
        """, unsafe_allow_html=True)

    with col_w:
        w_items = "".join([
            f'<div class="clean-row"><span style="color:#F59E0B; font-weight:900;">◈</span><span>{x}</span></div>'
            for x in rep["weaknesses"][:4]
        ])
        st.markdown(f"""
        <div class="bento-card bento-amber">
            <div class="bento-title" style="color:#FBBF24;">🎯 محاور يُنصح بتقويتها</div>
            {w_items}
        </div>
        """, unsafe_allow_html=True)

    # صف Bento السفلي: خطة التطوير
    plan_items = "".join([
        f'<div class="clean-row"><span class="step-badge">{i}</span><span>{x}</span></div>'
        for i, x in enumerate(rep["advice_and_plan"][:4], 1)
    ])
    st.markdown(f"""
    <div class="bento-card bento-blue">
        <div class="bento-title" style="color:#38BDF8;">🚀 خطتك المقترحة للتطوير</div>
        {plan_items}
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🗂️ بطاقات المراجعة التفاعلية (3D)", "📋 سجل الأسئلة العشرة"])
    with tab1:
        render_3d_flip_flashcards(rep["flashcards"])

    with tab2:
        for idx, item in enumerate(st.session_state.all_review, 1):
            badge = "جولة 1" if item["stage"] == 1 else "جولة 2"
            if item["is_correct"]:
                st.success(f"**س{idx} [{badge} - {item['course']}]:** إجابتك صحيحة ({item['student_answer']}) ✅\n\n*{item['explanation']}*")
            else:
                st.warning(f"**س{idx} [{badge} - {item['course']}]:** إجابتك ({item['student_answer']}) ❌ | الصحيح: **{item['correct_answer']}**\n\n*{item['explanation']}*")

    st.write("")
    if st.button("🔄 لاعب جديد", type="primary", use_container_width=True, key="btn_new_player"):
        reset_all()
        st.rerun()

# شريط توقيع الفريق الثابت أسفل الصفحة
st.markdown("""
<div class="brand-footer">
    <span>🧠 تطوير وتنفيذ فريق Agentic Mind</span>
    <span>بالتعاون مع إرث الأعمال • 2026</span>
</div>
""", unsafe_allow_html=True)