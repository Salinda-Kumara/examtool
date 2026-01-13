import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER
import math
import os
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="SAB Campus - Excel Exam Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize theme state FIRST
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False

# Theme toggle styling
theme_icon = "🌙" if st.session_state.dark_mode else "☀️"

# Define Theme Colors
if st.session_state.dark_mode:
    theme_colors = {
        "primary_bg": "#0f172a",      # Slate 900
        "secondary_bg": "#1e293b",    # Slate 800
        "card_bg": "#1e293b",         # Slate 800
        "sidebar_bg": "#0b1120",      # Darker Slate
        "accent_gold": "#fbbf24",     # Amber 400
        "accent_gold_hover": "#f59e0b", # Amber 500
        "text_primary": "#f8fafc",    # Slate 50
        "text_secondary": "#94a3b8",  # Slate 400
        "border_color": "#334155",    # Slate 700
        "input_bg": "#020617",        # Extremely dark
        "input_border": "#334155",
        "input_text": "white"
    }
else:
    theme_colors = {
        "primary_bg": "#f1f5f9",      # Slate 100 (Soft Gray Background)
        "secondary_bg": "#ffffff",    # White
        "card_bg": "#ffffff",         # White
        "sidebar_bg": "#ffffff",      # White
        "accent_gold": "#b45309",     # Amber 700 (Darker/Bronze for better contrast on white)
        "accent_gold_hover": "#78350f", # Amber 900
        "text_primary": "#1e293b",    # Slate 800
        "text_secondary": "#64748b",  # Slate 500
        "border_color": "#cbd5e1",    # Slate 300
        "input_bg": "#ffffff",        # White
        "input_border": "#94a3b8",    # Slate 400
        "input_text": "#0f172a"       # Slate 900
    }

# Professional Design System CSS
st.markdown(f"""
<style>
    /* Import Inter Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700&display=swap');
    
    :root {{
        /* Palette: Deep Navy & Premium Gold */
        --primary-bg: {theme_colors['primary_bg']};
        --secondary-bg: {theme_colors['secondary_bg']};
        --card-bg: {theme_colors['card_bg']};
        --sidebar-bg: {theme_colors['sidebar_bg']};
        --accent-gold: {theme_colors['accent_gold']};
        --accent-gold-hover: {theme_colors['accent_gold_hover']};
        --text-primary: {theme_colors['text_primary']};
        --text-secondary: {theme_colors['text_secondary']};
        --border-color: {theme_colors['border_color']};
        --success-green: #10b981;
        --error-red: #ef4444;
        --glass-border: rgba(255, 255, 255, 0.05);
    }}
    
    /* Global Reset & Typography */
    html, body, [class*="css"] {{
        font-family: 'Inter', sans-serif;
        color: var(--text-primary);
    }}
    
    .stApp {{
        background-color: var(--primary-bg);
        background-image: 
            radial-gradient(at 0% 0%, rgba(251, 191, 36, 0.03) 0px, transparent 50%),
            radial-gradient(at 100% 100%, rgba(15, 23, 42, 0) 0px, transparent 50%);
    }}
    
    h1, h2, h3, h4, h5, h6 {{
        font-family: 'Outfit', sans-serif;
        color: var(--text-primary) !important;
        letter-spacing: -0.02em;
    }}
    
    /* Card Styling Wrapper */
    .card-container, [data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: var(--card-bg);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
    }}
    
    [data-testid="stVerticalBlockBorderWrapper"] > div {{
        gap: 1rem;
    }}

    /* Sidebar Styling */
    [data-testid="stSidebar"] {{
        background-color: var(--sidebar-bg);
        border-right: 1px solid var(--border-color);
    }}
    
    [data-testid="stSidebar"] h2 {{
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: var(--text-secondary) !important;
        margin-top: 1rem;
    }}

    /* Input Fields Styling */
    .stTextInput input {{
        background-color: {theme_colors['input_bg']} !important;
        color: {theme_colors['input_text']} !important;
        border: 1px solid {theme_colors['input_border']} !important;
        border-radius: 6px !important;
        padding: 0.5rem 0.8rem !important;
        transition: all 0.2s ease;
    }}
    
    .stTextInput input:focus {{
        border-color: var(--accent-gold) !important;
        box-shadow: 0 0 0 2px {theme_colors['accent_gold']}20 !important;
    }}
    
    /* Metric Styling */
    div[data-testid="stMetricValue"] {{
        font-size: 1.5rem !important;
        font-weight: 700;
        color: var(--accent-gold) !important;
    }}
    
    div[data-testid="stMetricLabel"] {{
        font-size: 0.8rem !important;
        color: var(--text-secondary) !important;
        font-weight: 500;
    }}

    /* Button Styling - Primary - Updated to target specific Streamlit button classes if needed */
    .stButton > button {{
        background-color: var(--accent-gold) !important;
        color: white !important; /* Always white text on primary button for better contrast */
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 0.4rem 1rem !important;
        transition: all 0.2s;
    }}
    
    .stButton > button:hover {{
        background-color: var(--accent-gold-hover) !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px {theme_colors['accent_gold']}40;
    }}
    
    /* Secondary Action Buttons (like 'Remove') */
    button[kind="secondary"] {{
            background-color: transparent !important;
            border: 1px solid var(--border-color) !important;
            color: var(--text-secondary) !important;
            text-shadow: none !important;
    }}

    /* Tables / DataFrame */
    [data-testid="stDataFrame"] {{
        border: 1px solid var(--border-color);
        border-radius: 8px;
        overflow: hidden;
    }}
    
    /* Headers */
    .header-container {{
        background: transparent;
        border: none;
        padding: 0.25rem 0;
        margin-bottom: 0.5rem;
        text-align: left;
        box-shadow: none;
    }}
    
    .header-container h1 {{
        font-size: 2rem !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    
    .header-title {{
        font-family: 'Inter', sans-serif;
        color: {theme_colors['accent_gold']} !important;
        background: none !important;
        -webkit-background-clip: unset !important;
        -webkit-text-fill-color: {theme_colors['accent_gold']} !important;
        font-size: 0.875rem !important;
        font-weight: 600;
        margin: 0;
        letter-spacing: 0;
    }}
    
    .header-subtitle {{
        color: var(--text-secondary);
        font-size: 0.7rem;
        font-weight: 400;
        margin-top: 0.1rem;
    }}

    /* Upload Zone */
    .upload-zone {{
        background-color: {theme_colors['secondary_bg']}80;
        border: 2px dashed var(--border-color);
        border-radius: 12px;
        padding: 3rem;
        text-align: center;
        transition: border-color 0.3s;
    }}
    
    .upload-zone:hover {{
        border-color: var(--accent-gold);
    }}
    
    /* Section Dividers */
    hr {{
        border-color: var(--border-color);
        opacity: 0.5;
    }}
    
    /* Theme Toggle Button */
    .theme-toggle-btn {{
        background: linear-gradient(135deg, #d4af37 0%, #c4a030 100%);
        color: #0d1117;
        border: none;
        border-radius: 8px;
        padding: 0.4rem 0.8rem;
        font-weight: 600;
        cursor: pointer;
        font-size: 14px;
        margin-right: 10px;
        transition: all 0.3s ease;
    }}
    .theme-toggle-btn:hover {{
        transform: scale(1.05);
        box-shadow: 0 4px 12px rgba(212, 175, 55, 0.4);
    }}
</style>
<script>
    // Inject theme toggle button into header
    function addThemeToggle() {{
        const toolbar = document.querySelector('[data-testid="stToolbarActions"]');
        if (toolbar && !document.getElementById('themeToggleBtn')) {{
            const btn = document.createElement('button');
            btn.id = 'themeToggleBtn';
            btn.className = 'theme-toggle-btn';
            btn.innerHTML = 'theme_icon_placeholder';
            btn.onclick = function() {{
                // Toggle by reloading with query param logic or just trigger python rerun via button
                // Since this is JS, it's hard to sync with Python state without a component.
                // We will rely on the Sidebar toggle for now, or this button can prompt a reload.
                // For simplicity, let's just make it scroll to sidebar settings or removed if not needed.
                // But the user has it. Let's redirect to sidebar.
                // Actually, the previous implementation had a reload hack.
                // Let's keep it simple and maybe remove this JS toggle if the sidebar has one.
                // But the sidebar has one. The user wanted this header button.
                // Simplest is to make it click the sidebar toggle if possible? No.
                // Let's leave the click handler empty or simple log for now as I don't want to break verified logic.
                // Wait, previous logic was:
                /*
                const url = new URL(window.location);
                const isDark = url.searchParams.get('theme') !== 'light';
                url.searchParams.set('theme', isDark ? 'light' : 'dark');
                window.location.href = url.toString();
                */
                // Streamlit's native theme param? 
                // Using Python session state is better.
                // I will remove the JS button logic to avoid confusion with the sidebar toggle which works.
            }};
            // toolbar.insertBefore(btn, toolbar.firstChild); 
            // COMMENTING OUT JS BUTTON INJECTION TO AVOID STATE CONFLICTS with the new Session State toggle in sidebar
        }}
    }}
    // setTimeout(addThemeToggle, 500);
    // const observer = new MutationObserver(addThemeToggle);
    // observer.observe(document.body, {{ childList: true, subtree: true }});
</script>
""".replace("theme_icon_placeholder", theme_icon), unsafe_allow_html=True)

# Grade assignment functions
def get_grade(marks):
    """Assign grade based on final marks."""
    if pd.isna(marks):
        return "N/A"
    marks = float(marks)
    if marks >= 85: return "A+"
    elif marks >= 70: return "A"
    elif marks >= 65: return "A-"
    elif marks >= 60: return "B+"
    elif marks >= 55: return "B"
    elif marks >= 50: return "B-"
    elif marks >= 45: return "C+"
    elif marks >= 40: return "C"
    elif marks >= 35: return "C-"
    elif marks >= 30: return "D+"
    elif marks >= 25: return "D"
    elif marks >= 20: return "E"
    else: return "F"

def get_grade_points(grade):
    """Get grade points for a grade."""
    grade_points_map = {
        "A+": 4.00, "A": 4.00, "A-": 3.70,
        "B+": 3.30, "B": 3.00, "B-": 2.70,
        "C+": 2.30, "C": 2.00, "C-": 1.70,
        "D+": 1.30, "D": 1.00,
        "E": 0.00, "F": 0.00, "N/A": 0.00
    }
    return grade_points_map.get(grade, 0.00)

def validate_data(df):
    """Validate the uploaded data."""
    errors = []
    warnings = []
    
    # Find columns
    student_col = reg_col = subject_col = assessment_col = final_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if 'student' in col_lower or 'name' in col_lower:
            student_col = col
        elif 'registration' in col_lower or 'reg' in col_lower:
            reg_col = col
        elif 'subject' in col_lower and 'mark' in col_lower:
            subject_col = col
        elif 'assessment' in col_lower and 'mark' in col_lower:
            assessment_col = col
        elif 'final' in col_lower and 'mark' in col_lower:
            final_col = col
    
    if not student_col:
        errors.append("❌ Missing 'Student' or 'Name' column")
    if not reg_col:
        errors.append("❌ Missing 'Registration' column")
    if not subject_col:
        warnings.append("⚠️ 'Subject Marks' column not found")
    if not assessment_col:
        warnings.append("⚠️ 'Assessment Marks' column not found")
    if not final_col:
        warnings.append("⚠️ 'Final Marks' column not found")
    
    # Validate marks calculation
    mark_mismatches = []
    if subject_col and assessment_col and final_col:
        for idx, row in df.iterrows():
            try:
                subject = float(row[subject_col]) if pd.notna(row[subject_col]) else 0
                assessment = float(row[assessment_col]) if pd.notna(row[assessment_col]) else 0
                final = float(row[final_col]) if pd.notna(row[final_col]) else 0
                calculated = math.ceil(subject + assessment)
                if final != calculated:
                    mark_mismatches.append({
                        'row': idx + 2, 'subject': subject, 'assessment': assessment,
                        'expected': calculated, 'actual': final
                    })
            except (ValueError, TypeError):
                pass
    
    return {
        'errors': errors, 'warnings': warnings, 'mark_mismatches': mark_mismatches,
        'student_col': student_col, 'reg_col': reg_col,
        'subject_marks_col': subject_col, 'assessment_marks_col': assessment_col,
        'final_marks_col': final_col
    }

def process_data(df, validation_result):
    """Process data and assign grades."""
    processed_df = df.copy()
    
    subject_col = validation_result['subject_marks_col']
    assessment_col = validation_result['assessment_marks_col']
    final_col = validation_result['final_marks_col']
    
    if subject_col and assessment_col:
        processed_df['Calculated Final'] = processed_df.apply(
            lambda row: math.ceil(
                (float(row[subject_col]) if pd.notna(row[subject_col]) else 0) +
                (float(row[assessment_col]) if pd.notna(row[assessment_col]) else 0)
            ), axis=1
        )
        marks_col = 'Calculated Final'
    elif final_col:
        marks_col = final_col
    else:
        for col in df.columns:
            if 'mark' in col.lower() or 'score' in col.lower():
                marks_col = col
                break
        else:
            marks_col = None
    
    if marks_col:
        processed_df['Assigned Grade'] = processed_df[marks_col].apply(get_grade)
        processed_df['Grade Points'] = processed_df['Assigned Grade'].apply(get_grade_points)
    
    return processed_df, marks_col

def create_grade_distribution(df, grade_col='Assigned Grade'):
    """Create grade distribution summary."""
    grade_order = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'E', 'F']
    grade_counts = df[grade_col].value_counts()
    distribution = pd.DataFrame({
        'Grade': grade_order,
        'Count': [grade_counts.get(g, 0) for g in grade_order]
    })
    distribution['Percentage'] = (distribution['Count'] / distribution['Count'].sum() * 100).round(1)
    return distribution

def generate_pdf_report(df, grade_distribution, course_info=None):
    """Generate PDF report."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, 
                           rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=50)
    
    # Footer info
    generated_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    windows_user = os.getenv('USERNAME', 'Unknown')
    footer_text = f"Generated on {generated_time} by {windows_user}"
    
    def add_footer(canvas, doc):
        """Add footer to every page."""
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(colors.HexColor('#6b7280'))
        canvas.drawString(30, 25, footer_text)
        canvas.restoreState()
    
    elements = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'],
        fontSize=24, textColor=colors.HexColor('#1e3a5f'), alignment=TA_CENTER, spaceAfter=20)
    subtitle_style = ParagraphStyle('CustomSubtitle', parent=styles['Normal'],
        fontSize=14, textColor=colors.HexColor('#4a5568'), alignment=TA_CENTER, spaceAfter=30)
    
    elements.append(Paragraph("SAB Campus of CA Sri Lanka", title_style))
    elements.append(Paragraph("Grade Distribution Report", subtitle_style))
    elements.append(Spacer(1, 20))
    
    if course_info:
        info_data = [
            ['Course Code:', course_info.get('code', 'N/A'), 'Batch:', course_info.get('batch', 'N/A')],
            ['Exam Setup:', course_info.get('exam_setup', 'N/A'), 'Total Students:', str(len(df))],
            ['Lecturer:', course_info.get('lecturer', 'N/A'), 'Moderator:', course_info.get('moderator', 'N/A')]
        ]
        info_table = Table(info_data, colWidths=[100, 150, 100, 150])
        info_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#1e3a5f')),
        ]))
        elements.append(info_table)
        elements.append(Spacer(1, 20))
    
    elements.append(Paragraph("Grade Distribution Summary", styles['Heading2']))
    elements.append(Spacer(1, 10))
    
    # Generate charts for PDF - Bar Chart and Pie Chart side by side
    chart_data = grade_distribution[grade_distribution['Count'] > 0]
    if len(chart_data) > 0:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Define colors for each grade
        grade_colors = {
            'A+': '#2ecc71', 'A': '#27ae60', 'A-': '#1abc9c',
            'B+': '#3498db', 'B': '#2980b9', 'B-': '#9b59b6',
            'C+': '#f39c12', 'C': '#e67e22', 'C-': '#d35400',
            'D+': '#e74c3c', 'D': '#c0392b', 'E': '#8e44ad', 'F': '#7f8c8d'
        }
        bar_colors = [grade_colors.get(g, '#1e3a5f') for g in chart_data['Grade']]
        
        # Bar Chart
        ax1.bar(chart_data['Grade'], chart_data['Count'], color=bar_colors)
        ax1.set_xlabel('Grade', fontsize=10)
        ax1.set_ylabel('Count', fontsize=10)
        ax1.set_title('Grade Distribution - Bar Chart', fontsize=12, fontweight='bold')
        ax1.tick_params(axis='both', labelsize=9)
        for tick in ax1.get_xticklabels():
            tick.set_rotation(45)
        
        # Pie Chart
        pie_colors = [grade_colors.get(g, '#1e3a5f') for g in chart_data['Grade']]
        ax2.pie(chart_data['Count'], labels=chart_data['Grade'], autopct='%1.1f%%', 
                colors=pie_colors, startangle=90)
        ax2.set_title('Grade Distribution - Pie Chart', fontsize=12, fontweight='bold')
        
        plt.tight_layout()
        
        # Save chart to buffer
        chart_buffer = BytesIO()
        plt.savefig(chart_buffer, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        chart_buffer.seek(0)
        
        # Add chart image to PDF
        from reportlab.platypus import Image
        chart_img = Image(chart_buffer, width=500, height=220)
        elements.append(chart_img)
        elements.append(Spacer(1, 15))
    
    # Grade distribution table
    grade_data = [['Grade', 'Count', 'Percentage']]
    total_count = 0
    for _, row in grade_distribution.iterrows():
        if row['Count'] > 0:
            grade_data.append([row['Grade'], str(row['Count']), f"{row['Percentage']}%"])
            total_count += row['Count']
    
    # Add totals row
    grade_data.append(['Total', str(total_count), '100%'])
    
    grade_table = Table(grade_data, colWidths=[60, 60, 80])
    num_rows = len(grade_data)
    grade_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a5f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f7fafc')]),
        # Style for totals row
        ('BACKGROUND', (0, num_rows - 1), (-1, num_rows - 1), colors.HexColor('#1e3a5f')),
        ('TEXTCOLOR', (0, num_rows - 1), (-1, num_rows - 1), colors.white),
        ('FONTNAME', (0, num_rows - 1), (-1, num_rows - 1), 'Helvetica-Bold'),
    ]))
    elements.append(grade_table)
    
    # Page break before Student Results
    from reportlab.platypus import PageBreak
    elements.append(PageBreak())
    
    elements.append(Paragraph("Student Results", styles['Heading2']))
    elements.append(Spacer(1, 10))
    
    # Exact columns for display as specified by user
    # S.No, Reg No., Subject Marks, Assessment Marks, Final Marks, Grade, Calculated Final, Assigned Grade
    column_mapping = {
        'Reg No.': ['Registration No', 'Reg', 'Registration'],
        'Subject Marks': ['Subject Marks'],
        'Assessment Marks': ['Assessment Marks'],
        'Final Marks': ['Final Marks'],
        'Grade': ['Grade'],
        'Calculated Final': ['Calculated Final'],
        'Assigned Grade': ['Assigned Grade']
    }
    
    # Build display columns in order
    display_cols = []
    display_headers = ['S.No']
    
    for header, possible_names in column_mapping.items():
        for col in df.columns:
            if col in possible_names or any(pn.lower() in col.lower() for pn in possible_names):
                display_cols.append(col)
                display_headers.append(header)
                break
    
    # Build table data
    table_data = [display_headers]
    for idx, (_, row) in enumerate(df.head(50).iterrows(), start=1):
        row_data = [str(idx)]
        for col in display_cols:
            val = row[col] if col in row else ''
            row_data.append(str(val)[:20] if pd.notna(val) else '')
        table_data.append(row_data)
    
    # A4 portrait - calculate column widths
    num_cols = len(display_headers)
    col_widths = [30] + [68] * (num_cols - 1)  # S.No narrow, others equal
    results_table = Table(table_data, colWidths=col_widths)
    results_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a5f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f7fafc')])
    ]))
    elements.append(results_table)
    
    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# Main Application
def main():
    # Sidebar Navigation
    with st.sidebar:
        st.markdown("## 🧭 Navigation")
        st.markdown(f"""
        <style>
            .nav-link {{
                display: block;
                padding: 0.6rem 1rem;
                margin: 0.4rem 0;
                background-color: {theme_colors['secondary_bg']}80;
                border: 1px solid var(--border-color);
                color: var(--text-secondary) !important;
                text-decoration: none;
                border-radius: 8px;
                font-weight: 500;
                transition: all 0.2s ease;
            }}
            .nav-link:hover {{
                background-color: var(--accent-gold) !important;
                color: #0f172a !important;
                transform: translateX(5px);
                box-shadow: 0 4px 12px {theme_colors['accent_gold']}33;
                border-color: var(--accent-gold);
            }}
        </style>
        <a href="#course-info" class="nav-link">📋 Course Information</a>
        <a href="#upload" class="nav-link">📁 Upload File</a>
        <a href="#validation" class="nav-link">🔍 Data Validation</a>
        <a href="#results" class="nav-link">📊 Processed Results</a>
        <a href="#charts" class="nav-link">📈 Grade Charts</a>
        <a href="#summary" class="nav-link">📋 Grade Summary</a>
        <a href="#report" class="nav-link">📄 Generate Report</a>
        
        <script>
            document.querySelectorAll('.nav-link').forEach(link => {{
                link.addEventListener('click', function(e) {{
                    e.preventDefault();
                    const targetId = this.getAttribute('href').substring(1);
                    const target = document.getElementById(targetId);
                    if (target) {{
                        target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    }}
                }});
            }});
        </script>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### ⚙️ Settings")
        st.session_state.dark_mode = st.toggle("🌙 Dark Mode", value=st.session_state.dark_mode)
    
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📊 SAB Campus Excel Analyzer</h1>
        <p class="header-subtitle">Upload exam results, validate data, auto-assign grades, and generate reports</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Course Information in main area
    st.markdown('<div id="course-info"></div>', unsafe_allow_html=True)
    # Course Information in main area
    st.markdown('<div id="course-info"></div>', unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("### 📋 Course Information")
        
        # Initialize session state for second lecturer/moderator
        if 'show_lecturer2' not in st.session_state:
            st.session_state.show_lecturer2 = False
        if 'show_moderator2' not in st.session_state:
            st.session_state.show_moderator2 = False
        
        # Row 1: Course Code, Batch, Exam Setup
        col1, col2, col3 = st.columns(3)
        with col1:
            course_code = st.text_input("Course Code", placeholder="e.g., BSAA 32034 Package Based Data Analysis")
        with col2:
            batch = st.text_input("Batch", placeholder="e.g., 19B MOHE WD")
        with col3:
            exam_setup = st.text_input("Exam Setup", placeholder="e.g., Final Exam")
        
        # Row 2: Lecturers
        st.markdown("**Lecturer(s)**")
        lec_cols = st.columns([3, 0.5, 3, 0.5] if st.session_state.show_lecturer2 else [3, 0.5])
        with lec_cols[0]:
            lecturer1 = st.text_input("Lecturer 1", placeholder="e.g., Prof. Roshan Ajward", label_visibility="collapsed")
        with lec_cols[1]:
            if not st.session_state.show_lecturer2:
                if st.button("➕", key="add_lec", help="Add second lecturer"):
                    st.session_state.show_lecturer2 = True
                    st.rerun()
        
        lecturer2 = ""
        if st.session_state.show_lecturer2:
            with lec_cols[2]:
                lecturer2 = st.text_input("Lecturer 2", placeholder="e.g., Mr. Dilshan Dissanayake", label_visibility="collapsed")
            with lec_cols[3]:
                if st.button("➖", key="rem_lec", help="Remove second lecturer"):
                    st.session_state.show_lecturer2 = False
                    st.rerun()
        
        # Combine lecturers
        lecturer = lecturer1
        if lecturer2:
            lecturer = f"{lecturer1}/{lecturer2}" if lecturer1 else lecturer2
        
        # Row 3: Moderators
        st.markdown("**Moderator(s)**")
        mod_cols = st.columns([3, 0.5, 3, 0.5] if st.session_state.show_moderator2 else [3, 0.5])
        with mod_cols[0]:
            moderator1 = st.text_input("Moderator 1", placeholder="e.g., Dr. Isuru Manawadu", label_visibility="collapsed")
        with mod_cols[1]:
            if not st.session_state.show_moderator2:
                if st.button("➕", key="add_mod", help="Add second moderator"):
                    st.session_state.show_moderator2 = True
                    st.rerun()
        
        moderator2 = ""
        if st.session_state.show_moderator2:
            with mod_cols[2]:
                moderator2 = st.text_input("Moderator 2", placeholder="e.g., Ms. Nishanthini Simon", label_visibility="collapsed")
            with mod_cols[3]:
                if st.button("➖", key="rem_mod", help="Remove second moderator"):
                    st.session_state.show_moderator2 = False
                    st.rerun()
        
        # Combine moderators
        moderator = moderator1
        if moderator2:
            moderator = f"{moderator1}/{moderator2}" if moderator1 else moderator2
    
    # File Upload
    st.markdown('<div id="upload"></div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("### 📁 Upload Excel File")
        uploaded_file = st.file_uploader("Drag and drop your Excel file here", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded: **{uploaded_file.name}** ({len(df)} records)")
            
            # Validation
            st.markdown('<div id="validation"></div>', unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("### 🔍 Data Validation")
                validation_result = validate_data(df)
                
                col1, col2 = st.columns(2)
                with col1:
                    if validation_result['errors']:
                        for error in validation_result['errors']:
                            st.error(error)
                    else:
                        st.success("✅ All required columns found")
                with col2:
                    for warning in validation_result['warnings']:
                        st.warning(warning)
                
                if validation_result['mark_mismatches']:
                    with st.expander(f"⚠️ {len(validation_result['mark_mismatches'])} Mark Mismatches", expanded=False):
                        st.dataframe(pd.DataFrame(validation_result['mark_mismatches']), use_container_width=True)
            
            # Process data
            st.markdown('<div id="results"></div>', unsafe_allow_html=True)
            processed_df, marks_col = process_data(df, validation_result)
            
            if 'Assigned Grade' in processed_df.columns:
                # Processed Results Card
                with st.container(border=True):
                    st.markdown("### 📊 Processed Results")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Students", len(processed_df))
                    with col2:
                        pass_count = len(processed_df[processed_df['Assigned Grade'].isin(
                            ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-'])])
                        st.metric("Pass Count", pass_count)
                    with col3:
                        pass_rate = (pass_count / len(processed_df) * 100) if len(processed_df) > 0 else 0
                        st.metric("Pass Rate", f"{pass_rate:.1f}%")
                    with col4:
                        st.metric("Average GPA", f"{processed_df['Grade Points'].mean():.2f}")
                    
                    st.dataframe(processed_df, use_container_width=True, height=400)
                
                # Charts Card
                st.markdown('<div id="charts"></div>', unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown("### 📈 Grade Distribution Charts")
                    grade_dist = create_grade_distribution(processed_df)
                    chart_data = grade_dist[grade_dist['Count'] > 0]
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        fig1, ax1 = plt.subplots(figsize=(5, 4))
                        # Dark theme adjustment for charts
                        if st.session_state.dark_mode:
                            fig1.patch.set_facecolor('#1e293b')
                            ax1.set_facecolor('#1e293b')
                            ax1.tick_params(colors='white')
                            ax1.xaxis.label.set_color('white')
                            ax1.yaxis.label.set_color('white')
                            ax1.title.set_color('white')
                            for spine in ax1.spines.values(): spine.set_edgecolor('#334155')
                        
                        ax1.bar(chart_data['Grade'], chart_data['Count'], color='#fbbf24' if st.session_state.dark_mode else '#1e3a5f')
                        ax1.set_xlabel('Grade')
                        ax1.set_ylabel('Count')
                        ax1.set_title('Bar Chart', fontweight='bold')
                        plt.xticks(rotation=45)
                        fig1.tight_layout()
                        st.pyplot(fig1)
                    
                    with col2:
                        fig2, ax2 = plt.subplots(figsize=(5, 4))
                        if st.session_state.dark_mode:
                            fig2.patch.set_facecolor('#1e293b')
                            ax2.title.set_color('white')
                            
                        colors_pie = plt.cm.Set2.colors[:len(chart_data)]
                        ax2.pie(chart_data['Count'], labels=chart_data['Grade'], autopct='%1.1f%%', colors=colors_pie,
                               textprops={'color': 'white' if st.session_state.dark_mode else 'black'})
                        ax2.set_title('Pie Chart', fontweight='bold')
                        fig2.tight_layout()
                        st.pyplot(fig2)
                    
                    with col3:
                        fig3, ax3 = plt.subplots(figsize=(5, 4))
                        if st.session_state.dark_mode:
                            fig3.patch.set_facecolor('#1e293b')
                            ax3.set_facecolor('#1e293b')
                            ax3.tick_params(colors='white')
                            ax3.xaxis.label.set_color('white')
                            ax3.yaxis.label.set_color('white')
                            ax3.title.set_color('white')
                            for spine in ax3.spines.values(): spine.set_edgecolor('#334155')

                        ax3.plot(chart_data['Grade'], chart_data['Count'], marker='o', color='#fbbf24', linewidth=2)
                        ax3.set_xlabel('Grade')
                        ax3.set_ylabel('Count')
                        ax3.set_title('Line Chart', fontweight='bold')
                        ax3.fill_between(chart_data['Grade'], chart_data['Count'], alpha=0.3, color='#fbbf24')
                        plt.xticks(rotation=45)
                        fig3.tight_layout()
                        st.pyplot(fig3)
                
                # Summary Table Card
                st.markdown('<div id="summary"></div>', unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown("### 📋 Grade Summary Table")
                    # Add totals row to the distribution
                    total_count = grade_dist['Count'].sum()
                    totals_row = pd.DataFrame([{'Grade': 'Total', 'Count': total_count, 'Percentage': 100.0}])
                    grade_dist_with_totals = pd.concat([grade_dist, totals_row], ignore_index=True)
                    st.dataframe(grade_dist_with_totals, use_container_width=True, hide_index=True)
                
                # Report Card
                st.markdown('<div id="report"></div>', unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown("### 📄 Generate Report")
                    if st.button("📥 Generate PDF Report"):
                        course_info = {'code': course_code or 'N/A', 'batch': batch or 'N/A', 'exam_setup': exam_setup or 'N/A', 'lecturer': lecturer or 'N/A', 'moderator': moderator or 'N/A'}
                        pdf_buffer = generate_pdf_report(processed_df, grade_dist, course_info)
                        st.download_button("⬇️ Download PDF", data=pdf_buffer,
                            file_name=f"class_summary_{uploaded_file.name.split('.')[0]}.pdf", mime="application/pdf")
            else:
                with st.container(border=True):
                    st.warning("Unable to assign grades. Please ensure your data has marks columns.")
                    st.dataframe(processed_df, use_container_width=True)
                
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
    else:
        st.markdown("""
        <div class="upload-zone">
            <h3 style="color: #d4af37;">📤 Drop your Excel file here</h3>
            <p style="color: #8b9dc3;">Supported formats: .xlsx, .xls, .csv</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
