import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime
import os
try:
    from scipy.interpolate import make_interp_spline
except ImportError:
    make_interp_spline = None

# Page configuration
st.set_page_config(
    page_title="Semester Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize theme
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False

# Theme colors
if st.session_state.dark_mode:
    theme_colors = {
        "primary_bg": "#0f172a", "secondary_bg": "#1e293b", "card_bg": "#1e293b",
        "accent": "#fbbf24", "accent_hover": "#f59e0b", "text_primary": "#f8fafc",
        "text_secondary": "#94a3b8", "border": "#334155"
    }
else:
    theme_colors = {
        "primary_bg": "#f1f5f9", "secondary_bg": "#ffffff", "card_bg": "#ffffff",
        "accent": "#b45309", "accent_hover": "#78350f", "text_primary": "#1e293b",
        "text_secondary": "#64748b", "border": "#cbd5e1"
    }

# CSS Styling
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700&display=swap');
    
    * {{ font-family: 'Inter', sans-serif; }}
    
    .stApp {{
        background-color: {theme_colors['primary_bg']};
    }}
    
    h1, h2, h3, h4 {{
        font-family: 'Outfit', sans-serif;
        color: {theme_colors['text_primary']} !important;
    }}
    
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3,
    [data-testid="stMarkdownContainer"] p,
    label, span {{
        color: {theme_colors['text_primary']} !important;
    }}
    
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {theme_colors['card_bg']};
        border: 1px solid {theme_colors['border']};
        border-radius: 12px;
        padding: 1.5rem;
    }}
    
    [data-testid="stSidebar"] {{
        background-color: {theme_colors['secondary_bg']};
        border-right: 1px solid {theme_colors['border']};
    }}
    
    .stButton > button {{
        background-color: {theme_colors['accent']} !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
        transition: all 0.2s;
    }}
    
    .stButton > button:hover {{
        background-color: {theme_colors['accent_hover']} !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(180, 83, 9, 0.3);
    }}
    
    div[data-testid="stMetricValue"] {{
        font-size: 1.8rem !important;
        font-weight: 700;
        color: {theme_colors['accent']} !important;
    }}
    
    div[data-testid="stMetricLabel"] {{
        color: {theme_colors['text_secondary']} !important;
        font-weight: 500;
    }}
    
    .header-title {{
        color: {theme_colors['accent']} !important;
        font-size: 2rem !important;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }}
    
    .info-badge {{
        background: linear-gradient(135deg, {theme_colors['accent']} 0%, {theme_colors['accent_hover']} 100%);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-weight: 600;
        display: inline-block;
        margin: 0.25rem;
    }}
</style>
""", unsafe_allow_html=True)

def parse_semester_marksheet(file):
    """Parse the semester mark sheet Excel file - extracts only valid student records."""
    try:
        # Read the entire file without headers
        df_raw = pd.read_excel(file, header=None)
        
        # Extract metadata from specific rows
        course = str(df_raw.iloc[1, 8]) if not pd.isna(df_raw.iloc[1, 8]) else "Unknown Course"
        exam = str(df_raw.iloc[3, 8]) if not pd.isna(df_raw.iloc[3, 8]) else "Unknown Exam"
        subject = str(df_raw.iloc[5, 8]) if not pd.isna(df_raw.iloc[5, 8]) else "Unknown Subject"
        
        # Read student data starting from row 12 (0-indexed)
        # Every student has 2 rows: data row and blank row
        students = []
        for i in range(12, len(df_raw), 2):
            # Validate this is a student record row
            student_num = df_raw.iloc[i, 0]
            reg_num = df_raw.iloc[i, 1]
            grade = df_raw.iloc[i, 13]
            
            # Only include if:
            # 1. Student number exists and is numeric
            # 2. Registration number exists and is not empty
            # 3. Grade exists
            if pd.isna(student_num) or pd.isna(reg_num):
                # Stop when we hit rows without student number or registration number
                break
            
            # Validate student number is numeric (keep as string for display)
            try:
                # Verify it's a valid number but keep as string
                int(float(student_num))
                student_num_str = str(int(float(student_num)))
            except (ValueError, TypeError):
                # If not a valid number, skip this row
                continue
            
            # Validate registration number is a string and not empty
            reg_num_str = str(reg_num).strip()
            if not reg_num_str or reg_num_str == 'nan':
                continue
            
            # Get grade or set to N/A
            grade_str = str(grade).strip() if not pd.isna(grade) else "N/A"
            if grade_str == 'nan':
                grade_str = "N/A"
            
            # Only add valid student records
            students.append({
                '#': student_num_str,
                'Registration Number': reg_num_str,
                'Grade': grade_str
            })
        
        df_students = pd.DataFrame(students)
        
        # Final validation: ensure we have students
        if len(df_students) == 0:
            raise ValueError("No valid student records found in the file")
        
        metadata = {
            'course': course,
            'exam': exam,
            'subject': subject
        }
        
        return df_students, metadata
        
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None, None

def calculate_grade_distribution(df):
    """Calculate grade distribution statistics."""
    if 'Grade' not in df.columns:
        return pd.DataFrame()
    
    # Grade order - AB first, then worst to best
    grade_order = ['AB', 'E', 'D', 'D+', 'C-', 'C', 'C+', 'B-', 'B', 'B+', 'A-', 'A', 'A+']
    grade_counts = df['Grade'].value_counts()
    
    distribution = []
    for grade in grade_order:
        count = grade_counts.get(grade, 0)
        # Include all grades, even those with 0 counts
        distribution.append({
            'Grade': grade,
            'Count': count,
            'Percentage': round((count / len(df)) * 100, 1) if len(df) > 0 else 0
        })
    
    return pd.DataFrame(distribution)

def create_grade_chart(distribution_df):
    """Create grade distribution charts for web display (bar + pie)."""
    if len(distribution_df) == 0:
        return None
    
    # Filter out AB for the chart (but keep it in the table)
    chart_data = distribution_df[~distribution_df['Grade'].isin(['AB'])]
    
    if len(chart_data) == 0:
        return None
    
    grade_colors = {
        'A+': '#10b981', 'A': '#059669', 'A-': '#047857',
        'B+': '#3b82f6', 'B': '#2563eb', 'B-': '#1d4ed8',
        'C+': '#f59e0b', 'C': '#d97706', 'C-': '#b45309',
        'D+': '#ef4444', 'D': '#dc2626', 'E': '#991b1b', 'F': '#7f1d1d'
    }
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    
    # Bar chart with smooth line overlay
    colors_list = [grade_colors.get(g, '#6b7280') for g in chart_data['Grade']]
    ax1.bar(chart_data['Grade'], chart_data['Count'], color=colors_list, alpha=0.7)
    
    # Add smooth line chart overlay using interpolation
    if len(chart_data) > 2:  # Need at least 3 points for smooth curve
        x_pos = np.arange(len(chart_data))
        y_vals = chart_data['Count'].values
        
        # Use simple plot for small number of points or if scipy is missing
        if len(x_pos) > 2 and make_interp_spline is not None:
            try:
                # Create a smooth B-spline
                spline = make_interp_spline(x_pos, y_vals, k=3)  # Cubic spline
                
                x_smooth = np.linspace(x_pos.min(), x_pos.max(), 300)
                y_smooth = spline(x_smooth)
                
                # Ensure line doesn't dip below 0
                y_smooth = np.maximum(y_smooth, 0)
                
                ax1.plot(x_smooth, y_smooth, color='#b45309', linewidth=2, linestyle='-', zorder=3)
            except Exception:
                 # Fallback
             ax1.plot(x_pos, y_vals, color='#b45309', linewidth=2, zorder=3)
        else:
             ax1.plot(x_pos, y_vals, color='#b45309', linewidth=2, zorder=3)
    else:
        # Fallback for single/two points
        ax1.plot(chart_data['Grade'], chart_data['Count'], color='#b45309', marker='o', 
                 linewidth=2, markersize=8, markerfacecolor='white', 
                 markeredgewidth=2, markeredgecolor='#b45309', zorder=3)
    
    ax1.set_xlabel('Grade', fontsize=11, fontweight='bold')
    ax1.set_ylabel('Count', fontsize=11, fontweight='bold')
    ax1.set_title('Grade Distribution', fontsize=13, fontweight='bold')
    ax1.grid(axis='y', alpha=0.3)
    
    # Pie chart
    ax2.pie(chart_data['Count'], labels=chart_data['Grade'], autopct='%1.1f%%',
            colors=colors_list, startangle=90, textprops={'fontsize': 10})
    ax2.set_title('Grade Percentage', fontsize=13, fontweight='bold')
    
    plt.tight_layout()
    return fig

def create_grade_chart_pdf(distribution_df):
    """Create grade distribution bar chart only for PDF report."""
    if len(distribution_df) == 0:
        return None
    
    # Filter out AB for the chart (but keep it in the table)
    chart_data = distribution_df[~distribution_df['Grade'].isin(['AB'])]
    
    if len(chart_data) == 0:
        return None
    
    grade_colors = {
        'A+': '#10b981', 'A': '#059669', 'A-': '#047857',
        'B+': '#3b82f6', 'B': '#2563eb', 'B-': '#1d4ed8',
        'C+': '#f59e0b', 'C': '#d97706', 'C-': '#b45309',
        'D+': '#ef4444', 'D': '#dc2626', 'E': '#991b1b', 'F': '#7f1d1d'
    }
    
    # Single bar chart for PDF
    fig, ax = plt.subplots(1, 1, figsize=(10, 5))
    
    colors_list = [grade_colors.get(g, '#6b7280') for g in chart_data['Grade']]
    ax.bar(chart_data['Grade'], chart_data['Count'], color=colors_list, width=0.6, alpha=0.7)
    
    # Add smooth line chart overlay using interpolation
    if len(chart_data) > 2:  # Need at least 3 points for smooth curve
        x_pos = np.arange(len(chart_data))
        y_vals = chart_data['Count'].values
        
        # Use simple plot for small number of points or if scipy is missing
        if len(x_pos) > 2 and make_interp_spline is not None:
            try:
                # Create a smooth B-spline
                spline = make_interp_spline(x_pos, y_vals, k=3)
                
                x_smooth = np.linspace(x_pos.min(), x_pos.max(), 300)
                y_smooth = spline(x_smooth)
                
                # Ensure line doesn't dip below 0
                y_smooth = np.maximum(y_smooth, 0)
                
                ax.plot(x_smooth, y_smooth, color='#b45309', linewidth=2, linestyle='-', zorder=3)
            except Exception:
                # Fallback
             ax.plot(x_pos, y_vals, color='#b45309', linewidth=2, zorder=3)
        else:
             ax.plot(x_pos, y_vals, color='#b45309', linewidth=2, zorder=3)
    else:
        # Fallback for single point
        ax.plot(chart_data['Grade'], chart_data['Count'], color='#b45309', marker='o', 
                linewidth=2, markersize=8, markerfacecolor='white', 
                markeredgewidth=2, markeredgecolor='#b45309', zorder=3)
    
    ax.set_xlabel('Grade', fontsize=12, fontweight='bold')
    ax.set_ylabel('Number of Students', fontsize=12, fontweight='bold')
    ax.set_title('Grade Distribution', fontsize=14, fontweight='bold', pad=20)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Add value labels on top of bars
    for i, (grade, count) in enumerate(zip(chart_data['Grade'], chart_data['Count'])):
        ax.text(i, count, str(count), ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    plt.tight_layout()
    return fig

def generate_pdf_report(df, metadata, distribution_df):
    """Generate PDF report."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, 
                           rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Title
    title_style = ParagraphStyle('Title', parent=styles['Heading1'],
        fontSize=24, textColor=colors.HexColor('#1e3a8a'), alignment=TA_CENTER, 
        spaceAfter=5, fontName='Helvetica-Bold')
    
    report_type_style = ParagraphStyle('ReportType', parent=styles['Heading2'],
        fontSize=16, textColor=colors.HexColor('#334155'), alignment=TA_CENTER, 
        spaceAfter=10, fontName='Helvetica-Bold')
    
    subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'],
        fontSize=11, textColor=colors.HexColor('#64748b'), alignment=TA_CENTER, 
        spaceAfter=20)
    
    elements.append(Paragraph("SAB Campus of CA Sri Lanka", title_style))
    elements.append(Paragraph("Semester Mark Sheet Report", report_type_style))
    elements.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", subtitle_style))
    elements.append(Spacer(1, 5))
    
    # Course Information - Each field on new line
    info_style = ParagraphStyle('Info', parent=styles['Normal'], fontSize=10, leading=14)
    info_data = [
        ['Course:', metadata['course']],
        ['Exam:', metadata['exam']],
        ['Subject:', metadata['subject']],
        ['Total Students:', str(len(df))]
    ]
    
    info_table = Table(info_data, colWidths=[100, 430])
    info_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#1e293b')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 8))
    
    # Grade Distribution
    elements.append(Paragraph("<b>Grade Distribution Summary</b>", styles['Heading2']))
    elements.append(Spacer(1, 4))
    
    if len(distribution_df) > 0:
        # Create chart (bar chart only for PDF)
        fig = create_grade_chart_pdf(distribution_df)
        if fig:
            chart_buffer = BytesIO()
            plt.savefig(chart_buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close(fig)
            chart_buffer.seek(0)
            
            chart_img = Image(chart_buffer, width=460, height=255)
            elements.append(chart_img)
            elements.append(Spacer(1, 8))
        
        # Distribution table
        dist_data = [['Grade', 'Count', 'Percentage']]
        total_count = 0
        
        for _, row in distribution_df.iterrows():
            dist_data.append([row['Grade'], str(row['Count']), f"{row['Percentage']}%"])
            total_count += row['Count']
        
        # Add totals row (percentage is always 100%)
        dist_data.append(['Total', str(total_count), "100%"])
        
        dist_table = Table(dist_data, colWidths=[70, 70, 80])
        num_rows = len(dist_data)
        dist_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a8a')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cbd5e1')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f1f5f9')]),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            # Style for totals row
            ('BACKGROUND', (0, num_rows - 1), (-1, num_rows - 1), colors.HexColor('#1e3a8a')),
            ('TEXTCOLOR', (0, num_rows - 1), (-1, num_rows - 1), colors.white),
            ('FONTNAME', (0, num_rows - 1), (-1, num_rows - 1), 'Helvetica-Bold'),
        ]))
        elements.append(dist_table)
        elements.append(Spacer(1, 10))
    
    elements.append(PageBreak())
    
    # Student Results
    elements.append(Paragraph("<b>Student Results</b>", styles['Heading2']))
    elements.append(Spacer(1, 10))
    
    student_data = [['#', 'Registration Number', 'Grade']]
    for _, row in df.iterrows():
        student_data.append([
            str(row['#'])[:5],
            str(row['Registration Number'])[:35],
            str(row['Grade'])
        ])
    
    student_table = Table(student_data, colWidths=[40, 350, 60])
    student_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a8a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),  # Align Registration Number column to left
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')])
    ]))
    elements.append(student_table)
    
    # Footer
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#64748b'))
        canvas.drawString(40, 25, f"Generated by {os.getenv('USERNAME', 'User')} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        canvas.setFont('Helvetica', 7)
        canvas.drawString(40, 15, "Dev@Salinda")
        canvas.restoreState()
    
    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# Main App
def main():
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        st.session_state.dark_mode = st.toggle("🌙 Dark Mode", value=st.session_state.dark_mode)
        st.markdown("---")
        st.markdown("### 📖 About")
        st.info("Upload semester mark sheets and generate comprehensive reports with grade distributions and analytics.")
    
    # Header
    st.markdown('<h1 class="header-title">📊 Semester Report Generator</h1>', unsafe_allow_html=True)
    st.markdown("Upload your semester mark sheet Excel file to generate detailed reports")
    st.markdown("---")
    
    # File Upload
    with st.container(border=True):
        st.markdown("### 📁 Upload Mark Sheet")
        uploaded_file = st.file_uploader("Drop your Excel file here", type=['xls', 'xlsx'], 
                                        help="Upload the semester mark sheet Excel file")
    
    if uploaded_file:
        # Parse the file
        df, metadata = parse_semester_marksheet(uploaded_file)
        
        if df is not None and metadata is not None:
            # Display metadata
            with st.container(border=True):
                st.markdown("### 📋 Course Information")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="info-badge">📚 {metadata["course"]}</div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="info-badge">📝 {metadata["exam"]}</div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="info-badge">📖 {metadata["subject"]}</div>', unsafe_allow_html=True)
            
            # Statistics
            distribution_df = calculate_grade_distribution(df)
            
            with st.container(border=True):
                st.markdown("### 📊 Statistics")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Students", len(df))
                with col2:
                    a_grades = len(df[df['Grade'].str.contains('A', na=False)])
                    st.metric("A Grades", a_grades)
                with col3:
                    b_grades = len(df[df['Grade'].str.contains('B', na=False)])
                    st.metric("B Grades", b_grades)
                with col4:
                    if len(distribution_df) > 0:
                        top_grade = distribution_df.iloc[0]['Grade']
                        st.metric("Most Common", top_grade)
            
            # Grade Distribution
            if len(distribution_df) > 0:
                with st.container(border=True):
                    st.markdown("### 📈 Grade Distribution")
                    
                    col1, col2 = st.columns([3, 2])
                    with col1:
                        fig = create_grade_chart(distribution_df)
                        if fig:
                            st.pyplot(fig)
                    
                    with col2:
                        st.dataframe(distribution_df, hide_index=True, width='stretch')
            
            # Student Data
            with st.container(border=True):
                st.markdown("### 👥 Student Results")
                st.dataframe(df, hide_index=True, width='stretch')
            
            # Generate Reports
            with st.container(border=True):
                st.markdown("### 📄 Generate Reports")
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.button("📥 Download PDF Report", width='stretch'):
                        pdf_buffer = generate_pdf_report(df, metadata, distribution_df)
                        st.download_button(
                            label="⬇️ Download PDF",
                            data=pdf_buffer,
                            file_name=f"Semester_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf",
                            width='stretch'
                        )
                
                with col2:
                    if st.button("📥 Download Excel Report", width='stretch'):
                        excel_buffer = BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, sheet_name='Student Results', index=False)
                            distribution_df.to_excel(writer, sheet_name='Grade Distribution', index=False)
                        excel_buffer.seek(0)
                        
                        st.download_button(
                            label="⬇️ Download Excel",
                            data=excel_buffer,
                            file_name=f"Semester_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
    else:
        # Show welcome message
        st.info("👆 Please upload a semester mark sheet Excel file to get started")

if __name__ == "__main__":
    main()
