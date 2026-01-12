"""
Excel Analyzer - SAB Campus Grade Distribution Tool
Professional UI with modern design matching SAB Campus branding
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import os
from datetime import datetime


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SAB Campus - Grade Distribution Analyzer")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 650)
        
        # Data
        self.df = None
        self.file_path = None
        self.analysis_data = {}
        
        # Professional Color Palette
        self.colors = {
            'primary': '#1e3a5f',        # Deep navy blue
            'primary_light': '#2c5282',   # Lighter navy
            'accent': '#c9a227',          # Gold accent
            'bg': '#f5f7fa',              # Light gray background
            'card': '#ffffff',            # White cards
            'text': '#1a202c',            # Dark text
            'text_light': '#718096',      # Light gray text
            'border': '#e2e8f0',          # Card borders
            'success': '#38a169',         # Green
            'table_header': '#2c5282',    # Table header blue
            'table_row1': '#eef2f7',      # Table alternate row
            'table_row2': '#ffffff',      # Table white row
        }
        
        self.setup_styles()
        self.create_ui()
        
    def setup_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("TFrame", background=self.colors['bg'])
        style.configure("Card.TFrame", background=self.colors['card'])
        style.configure("TLabel", background=self.colors['bg'], 
                       foreground=self.colors['text'], font=("Segoe UI", 10))
        
    def create_ui(self):
        """Create professional UI layout"""
        self.root.configure(bg=self.colors['bg'])
        
        # Main container
        main = tk.Frame(self.root, bg=self.colors['bg'])
        main.pack(fill=tk.BOTH, expand=True)
        
        # === HEADER BAR ===
        header = tk.Frame(main, bg=self.colors['primary'], height=70)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        header_inner = tk.Frame(header, bg=self.colors['primary'])
        header_inner.pack(fill=tk.BOTH, expand=True, padx=25, pady=15)
        
        # Logo/Title
        title_frame = tk.Frame(header_inner, bg=self.colors['primary'])
        title_frame.pack(side=tk.LEFT)
        
        tk.Label(title_frame, text="SAB Campus", font=("Segoe UI", 18, "bold"),
                fg="white", bg=self.colors['primary']).pack(side=tk.LEFT)
        tk.Label(title_frame, text=" of CA Sri Lanka", font=("Segoe UI", 18),
                fg=self.colors['accent'], bg=self.colors['primary']).pack(side=tk.LEFT)
        
        # Header buttons
        btn_frame = tk.Frame(header_inner, bg=self.colors['primary'])
        btn_frame.pack(side=tk.RIGHT)
        
        self._create_header_btn(btn_frame, "📄 Export", self.export_report, self.colors['success'])
        self._create_header_btn(btn_frame, "📂 Open File", self.browse_file, self.colors['accent'])
        
        # === CONTENT ===
        content = tk.Frame(main, bg=self.colors['bg'])
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Course Info Bar
        self._create_course_info(content)
        
        # Main content area
        main_content = tk.Frame(content, bg=self.colors['bg'])
        main_content.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        main_content.grid_columnconfigure(0, weight=2)
        main_content.grid_columnconfigure(1, weight=3)
        main_content.grid_rowconfigure(0, weight=1)
        
        # Left: Grade Table
        self._create_grade_panel(main_content)
        
        # Right: Chart
        self._create_chart_panel(main_content)
        
    def _create_header_btn(self, parent, text, command, bg_color):
        """Create styled header button"""
        btn = tk.Frame(parent, bg=bg_color, cursor="hand2")
        btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        lbl = tk.Label(btn, text=text, font=("Segoe UI", 10, "bold"),
                      fg="white" if bg_color != self.colors['accent'] else self.colors['primary'],
                      bg=bg_color, padx=15, pady=8)
        lbl.pack()
        
        for widget in [btn, lbl]:
            widget.bind("<Button-1>", lambda e, c=command: c())
            widget.bind("<Enter>", lambda e, b=btn: b.configure(bg=self._lighten(bg_color)))
            widget.bind("<Leave>", lambda e, b=btn, c=bg_color: b.configure(bg=c))
    
    def _lighten(self, color):
        """Lighten a hex color"""
        r = min(255, int(color[1:3], 16) + 20)
        g = min(255, int(color[3:5], 16) + 20)
        b = min(255, int(color[5:7], 16) + 20)
        return f"#{r:02x}{g:02x}{b:02x}"
        
    def _create_course_info(self, parent):
        """Create course information bar"""
        info_card = tk.Frame(parent, bg=self.colors['card'], 
                            highlightbackground=self.colors['border'], highlightthickness=1)
        info_card.pack(fill=tk.X)
        
        inner = tk.Frame(info_card, bg=self.colors['card'])
        inner.pack(fill=tk.X, padx=20, pady=15)
        
        # Create info items
        left = tk.Frame(inner, bg=self.colors['card'])
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        right = tk.Frame(inner, bg=self.colors['card'])
        right.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Left column
        row1 = tk.Frame(left, bg=self.colors['card'])
        row1.pack(fill=tk.X, pady=2)
        self._info_field(row1, "Course Code:", "course_code")
        
        row2 = tk.Frame(left, bg=self.colors['card'])
        row2.pack(fill=tk.X, pady=2)
        self._info_field(row2, "Total Students:", "total_students", bold_value=True)
        
        # Right column
        row3 = tk.Frame(right, bg=self.colors['card'])
        row3.pack(fill=tk.X, pady=2)
        self._info_field(row3, "Course Lecturer:", "lecturer")
        
        row4 = tk.Frame(right, bg=self.colors['card'])
        row4.pack(fill=tk.X, pady=2)
        self._info_field(row4, "Course Moderator:", "moderator")
        
    def _info_field(self, parent, label, attr_name, bold_value=False):
        """Create info field with label and value"""
        tk.Label(parent, text=label, font=("Segoe UI", 10, "bold"),
                fg=self.colors['text'], bg=self.colors['card']).pack(side=tk.LEFT)
        
        value_label = tk.Label(parent, text="--", 
                              font=("Segoe UI", 10, "bold" if bold_value else "normal"),
                              fg=self.colors['primary'] if bold_value else self.colors['text_light'],
                              bg=self.colors['card'])
        value_label.pack(side=tk.LEFT, padx=(5, 0))
        setattr(self, f"{attr_name}_label", value_label)
        
    def _create_grade_panel(self, parent):
        """Create grade distribution panel"""
        panel = tk.Frame(parent, bg=self.colors['card'],
                        highlightbackground=self.colors['border'], highlightthickness=1)
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Title
        title_bar = tk.Frame(panel, bg=self.colors['primary'])
        title_bar.pack(fill=tk.X)
        
        tk.Label(title_bar, text="📊 Grade Distribution", font=("Segoe UI", 11, "bold"),
                fg="white", bg=self.colors['primary'], padx=15, pady=10).pack(anchor=tk.W)
        
        # Content
        self.grade_container = tk.Frame(panel, bg=self.colors['card'])
        self.grade_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Placeholder
        tk.Label(self.grade_container, text="Load an Excel file to view\ngrade distribution",
                font=("Segoe UI", 10), fg=self.colors['text_light'],
                bg=self.colors['card'], justify=tk.CENTER).pack(expand=True)
        
    def _create_chart_panel(self, parent):
        """Create chart panel"""
        panel = tk.Frame(parent, bg=self.colors['card'],
                        highlightbackground=self.colors['border'], highlightthickness=1)
        panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        # Title
        title_bar = tk.Frame(panel, bg=self.colors['primary'])
        title_bar.pack(fill=tk.X)
        
        tk.Label(title_bar, text="📈 Grade Distribution Chart", font=("Segoe UI", 11, "bold"),
                fg="white", bg=self.colors['primary'], padx=15, pady=10).pack(anchor=tk.W)
        
        # Content
        self.chart_container = tk.Frame(panel, bg=self.colors['card'])
        self.chart_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Placeholder
        tk.Label(self.chart_container, text="Chart will appear here",
                font=("Segoe UI", 10), fg=self.colors['text_light'],
                bg=self.colors['card']).pack(expand=True)
        
    def browse_file(self):
        """Open file dialog"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.load_and_analyze()
            
    def load_and_analyze(self):
        """Load and analyze Excel file"""
        if not self.file_path:
            return
            
        try:
            if self.file_path.endswith('.xls'):
                self.df = pd.read_excel(self.file_path, engine='xlrd')
            else:
                self.df = pd.read_excel(self.file_path, engine='openpyxl')
            
            self.analyze_data()
            self.display_stats()
            self.create_charts()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
            
    def analyze_data(self):
        """Analyze loaded data - calculate final marks and grades"""
        if self.df is None:
            return
        
        import math
        
        # Calculate Final Marks if Subject and Assessment exist
        if 'Subject Marks' in self.df.columns and 'Assessment Marks' in self.df.columns:
            # Final Marks = Subject + Assessment, rounded up to nearest integer
            self.df['Calculated Final'] = self.df.apply(
                lambda row: math.ceil(row['Subject Marks'] + row['Assessment Marks']) 
                if pd.notna(row['Subject Marks']) and pd.notna(row['Assessment Marks']) 
                else None, axis=1
            )
            
            # Generate grades based on calculated final marks
            def get_grade(marks):
                if pd.isna(marks):
                    return 'AB'  # Absent
                marks = int(marks)
                if marks >= 85: return 'A+'
                elif marks >= 80: return 'A'
                elif marks >= 75: return 'A-'
                elif marks >= 70: return 'B+'
                elif marks >= 65: return 'B'
                elif marks >= 60: return 'B-'
                elif marks >= 55: return 'C+'
                elif marks >= 50: return 'C'
                elif marks >= 45: return 'C-'
                elif marks >= 40: return 'D+'
                elif marks >= 35: return 'D'
                else: return 'E'
            
            self.df['Calculated Grade'] = self.df['Calculated Final'].apply(get_grade)
            
            # Use calculated grade for distribution
            grade_column = 'Calculated Grade'
        else:
            grade_column = 'Grade' if 'Grade' in self.df.columns else None
            
        self.analysis_data = {
            'file_name': os.path.basename(self.file_path),
            'total_rows': len(self.df),
            'grade_distribution': {},
        }
        
        if grade_column and grade_column in self.df.columns:
            self.analysis_data['grade_distribution'] = self.df[grade_column].value_counts().to_dict()
                
    def display_stats(self):
        """Display statistics"""
        for widget in self.grade_container.winfo_children():
            widget.destroy()
        
        data = self.analysis_data
        
        # Update info labels
        self.course_code_label.config(text=data['file_name'].replace('.xlsx', '').replace('.xls', ''))
        self.total_students_label.config(text=str(data['total_rows']))
        
        # Create grade table
        if data['grade_distribution']:
            self._create_grade_table(self.grade_container, data['grade_distribution'])
    
    def _create_grade_table(self, parent, grade_dist):
        """Create professional grade table"""
        table = tk.Frame(parent, bg=self.colors['border'])
        table.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = tk.Frame(table, bg=self.colors['table_header'])
        header_frame.pack(fill=tk.X)
        
        headers = [("Final Grade", 100), ("No", 60), ("%", 70)]
        for text, width in headers:
            tk.Label(header_frame, text=text, font=("Segoe UI", 10, "bold"),
                    fg="white", bg=self.colors['table_header'],
                    width=width//8, pady=8).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Scrollable body
        canvas = tk.Canvas(table, bg=self.colors['card'], highlightthickness=0, height=350)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas, bg=self.colors['card'])
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.create_window((0, 0), window=body, anchor="nw")
        
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))
        
        # Grade order and colors (including AB for Absent)
        grade_order = ['E', 'D', 'D+', 'C-', 'C', 'C+', 'B-', 'B', 'B+', 'A-', 'A', 'A+', 'AB']
        total = sum(grade_dist.values())
        
        for i, grade in enumerate(grade_order):
            count = grade_dist.get(grade, 0)
            pct = (count / total * 100) if total > 0 else 0
            bg = self.colors['table_row1'] if i % 2 == 0 else self.colors['table_row2']
            
            row = tk.Frame(body, bg=bg)
            row.pack(fill=tk.X)
            
            tk.Label(row, text=grade, font=("Segoe UI", 10), width=12,
                    bg=bg, fg=self.colors['text'], pady=6).pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(row, text=str(count), font=("Segoe UI", 10), width=8,
                    bg=bg, fg=self.colors['primary'], pady=6).pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(row, text=f"{pct:.0f}%", font=("Segoe UI", 10), width=8,
                    bg=bg, fg=self.colors['text'], pady=6).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Total row
        total_row = tk.Frame(body, bg=self.colors['primary'])
        total_row.pack(fill=tk.X)
        
        tk.Label(total_row, text="Total", font=("Segoe UI", 10, "bold"), width=12,
                bg=self.colors['primary'], fg="white", pady=8).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(total_row, text=str(total), font=("Segoe UI", 10, "bold"), width=8,
                bg=self.colors['primary'], fg=self.colors['accent'], pady=8).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(total_row, text="100%", font=("Segoe UI", 10, "bold"), width=8,
                bg=self.colors['primary'], fg="white", pady=8).pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    def create_charts(self):
        """Create grade distribution chart"""
        for widget in self.chart_container.winfo_children():
            widget.destroy()
            
        if not self.analysis_data.get('grade_distribution'):
            return
        
        fig = Figure(figsize=(6, 5), facecolor='white', dpi=90)
        ax = fig.add_subplot(111)
        ax.set_facecolor('white')
        
        grade_order = ['E', 'D', 'D+', 'C-', 'C', 'C+', 'B-', 'B', 'B+', 'A-', 'A', 'A+', 'AB']
        grade_dist = self.analysis_data['grade_distribution']
        
        grades = [g for g in grade_order if g in grade_dist or True]
        total = sum(grade_dist.values())
        percentages = [(grade_dist.get(g, 0) / total * 100) for g in grades]
        
        # Cumulative
        cumulative = []
        cum = 0
        for p in percentages:
            cum += p
            cumulative.append(cum)
        
        x = range(len(grades))
        
        # Bars
        bars = ax.bar(x, percentages, color=self.colors['primary_light'], 
                     edgecolor=self.colors['primary'], linewidth=1, alpha=0.8, zorder=2)
        
        # Line
        ax.plot(x, cumulative, color=self.colors['accent'], linestyle='--', 
               marker='o', markersize=4, linewidth=2, zorder=3)
        
        ax.set_xticks(x)
        ax.set_xticklabels(grades, fontsize=9)
        ax.set_xlabel('Grades', fontsize=10, color=self.colors['text_light'])
        ax.set_ylabel('Percentage (%)', fontsize=10, color=self.colors['text_light'])
        ax.set_title('Grade Distribution', fontsize=12, fontweight='bold', 
                    color=self.colors['text'], pad=15)
        ax.tick_params(colors=self.colors['text_light'], labelsize=9)
        ax.grid(axis='y', linestyle='-', alpha=0.3, color=self.colors['border'])
        ax.set_ylim(0, max(max(percentages) * 1.2, 105))
        ax.legend(['Cumulative %', 'Grade %'], loc='upper left', fontsize=8)
        
        for spine in ax.spines.values():
            spine.set_color(self.colors['border'])
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figure = fig
        
    def export_report(self):
        """Export HTML report"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="Save Report",
            defaultextension=".html",
            filetypes=[("HTML Files", "*.html")],
            initialfile=f"Grade_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        )
        
        if not file_path:
            return
            
        try:
            self._generate_html(file_path)
            messagebox.showinfo("Success", "Report saved successfully!")
            import webbrowser
            webbrowser.open(f'file:///{file_path}')
        except Exception as e:
            messagebox.showerror("Error", f"Failed: {str(e)}")
            
    def _generate_html(self, path):
        """Generate HTML report matching exam-visualizer style"""
        data = self.analysis_data
        grade_dist = data.get('grade_distribution', {})
        total = sum(grade_dist.values())
        
        grade_order = ['E', 'D', 'D+', 'C-', 'C', 'C+', 'B-', 'B', 'B+', 'A-', 'A', 'A+', 'AB']
        
        # Grade colors
        grade_colors = {
            'A+': '#059669', 'A': '#10b981', 'A-': '#14b8a6',
            'B+': '#3b82f6', 'B': '#6366f1', 'B-': '#8b5cf6',
            'C+': '#f59e0b', 'C': '#f97316', 'C-': '#fb923c',
            'D+': '#ef4444', 'D': '#dc2626', 'E': '#991b1b', 'AB': '#6b7280'
        }
        
        # Build table rows
        rows = ""
        for g in grade_order:
            c = grade_dist.get(g, 0)
            p = (c/total*100) if total else 0
            color = grade_colors.get(g, '#6b7280')
            rows += f'''<tr>
                <td><span class="grade-badge" style="background:{color}">{g}</span></td>
                <td class="count">{c}</td>
                <td>{p:.0f}%</td>
            </tr>'''
        
        # Build grade summary badges
        badges = ""
        for g in grade_order:
            c = grade_dist.get(g, 0)
            if c > 0:
                p = (c/total*100) if total else 0
                color = grade_colors.get(g, '#6b7280')
                badges += f'<div class="grade-card" style="border-left:4px solid {color}"><div class="grade-label">{g}</div><div class="grade-count">{c}</div><div class="grade-pct">{p:.0f}%</div></div>'
        
        html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Grade Distribution Report - {data['file_name']}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{
            --primary: #4361ee;
            --primary-dark: #3730a3;
            --accent: #7c3aed;
            --success: #10b981;
            --bg: #f4f6f9;
            --card: #ffffff;
            --text: #2d3748;
            --text-light: #718096;
            --border: #e2e8f0;
            --shadow: 0 4px 12px rgba(0,0,0,0.1);
            --radius: 12px;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); line-height: 1.6; }}
        .container {{ max-width: 1200px; margin: 0 auto; padding: 20px; }}
        
        /* Header */
        .header {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            color: white;
            text-align: center;
            padding: 50px 30px;
            border-radius: var(--radius);
            margin-bottom: 30px;
            box-shadow: var(--shadow);
        }}
        .header h1 {{ font-size: 2rem; margin-bottom: 8px; }}
        .header p {{ opacity: 0.9; font-size: 1rem; }}
        
        /* Stats Row */
        .stats-row {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: var(--card);
            padding: 25px;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            text-align: center;
        }}
        .stat-card .icon {{ font-size: 2rem; margin-bottom: 10px; }}
        .stat-card .value {{ font-size: 2rem; font-weight: 700; color: var(--text); }}
        .stat-card .label {{ color: var(--text-light); font-size: 0.9rem; }}
        
        /* Cards */
        .card {{
            background: var(--card);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            margin-bottom: 25px;
            overflow: hidden;
        }}
        .card-header {{
            padding: 18px 25px;
            font-size: 1.1rem;
            font-weight: 600;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .card-body {{ padding: 25px; }}
        
        /* Grid */
        .grid {{ display: grid; grid-template-columns: 1fr 2fr; gap: 25px; }}
        
        /* Table */
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ background: var(--primary); color: white; padding: 12px; text-align: left; font-weight: 600; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); }}
        tr:hover {{ background: #f8fafc; }}
        .total-row {{ background: var(--primary-dark) !important; color: white; font-weight: 600; }}
        .total-row td {{ color: white; border: none; }}
        .grade-badge {{ display: inline-block; padding: 4px 12px; border-radius: 20px; color: white; font-weight: 600; font-size: 0.85rem; }}
        .count {{ font-weight: 600; color: var(--primary); }}
        
        /* Grade Summary */
        .grade-summary {{ display: flex; flex-wrap: wrap; gap: 12px; }}
        .grade-card {{
            background: #f8fafc;
            padding: 15px 20px;
            border-radius: 8px;
            text-align: center;
            min-width: 80px;
        }}
        .grade-label {{ font-weight: 700; font-size: 1.1rem; color: var(--text); }}
        .grade-count {{ font-size: 1.5rem; font-weight: 700; color: var(--primary); }}
        .grade-pct {{ font-size: 0.8rem; color: var(--text-light); }}
        
        /* Chart */
        .chart-container {{ height: 350px; }}
        
        /* Actions */
        .actions {{ text-align: center; padding: 30px; }}
        .btn {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 14px 28px;
            background: var(--success);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }}
        .btn:hover {{ background: #059669; transform: translateY(-2px); }}
        
        @media print {{
            .actions {{ display: none; }}
            .header {{ break-after: avoid; }}
            body {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header class="header">
            <h1><i class="fas fa-graduation-cap"></i> SAB Campus of CA Sri Lanka</h1>
            <p>Grade Distribution Report • {data['file_name']} • {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        </header>
        
        <div class="stats-row">
            <div class="stat-card">
                <div class="icon" style="color:#4361ee"><i class="fas fa-users"></i></div>
                <div class="value">{data['total_rows']}</div>
                <div class="label">Total Students</div>
            </div>
            <div class="stat-card">
                <div class="icon" style="color:#10b981"><i class="fas fa-check-circle"></i></div>
                <div class="value">{sum(grade_dist.get(g,0) for g in ['A+','A','A-','B+','B','B-','C+','C','C-'])}</div>
                <div class="label">Passed</div>
            </div>
            <div class="stat-card">
                <div class="icon" style="color:#ef4444"><i class="fas fa-times-circle"></i></div>
                <div class="value">{sum(grade_dist.get(g,0) for g in ['D+','D','E'])}</div>
                <div class="label">Failed</div>
            </div>
            <div class="stat-card">
                <div class="icon" style="color:#6b7280"><i class="fas fa-user-slash"></i></div>
                <div class="value">{grade_dist.get('AB',0)}</div>
                <div class="label">Absent</div>
            </div>
        </div>
        
        <div class="card">
            <div class="card-header"><i class="fas fa-th-large"></i> Grade Summary</div>
            <div class="card-body">
                <div class="grade-summary">{badges}</div>
            </div>
        </div>
        
        <div class="grid">
            <div class="card">
                <div class="card-header"><i class="fas fa-table"></i> Grade Distribution</div>
                <div class="card-body">
                    <table>
                        <thead><tr><th>Grade</th><th>Count</th><th>Percentage</th></tr></thead>
                        <tbody>
                            {rows}
                            <tr class="total-row"><td>Total</td><td>{total}</td><td>100%</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header"><i class="fas fa-chart-bar"></i> Distribution Chart</div>
                <div class="card-body">
                    <div class="chart-container">
                        <canvas id="gradeChart"></canvas>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="actions">
            <button class="btn" onclick="window.print()">
                <i class="fas fa-print"></i> Print Report
            </button>
        </div>
    </div>
    
    <script>
        const grades = {list(grade_order)};
        const counts = {[grade_dist.get(g,0) for g in grade_order]};
        const total = {total if total > 0 else 1};
        const pcts = counts.map(c => (c/total*100).toFixed(1));
        let cum = 0;
        const cumulative = pcts.map(p => {{ cum += parseFloat(p); return cum.toFixed(1); }});
        
        new Chart(document.getElementById('gradeChart'), {{
            type: 'bar',
            data: {{
                labels: grades,
                datasets: [
                    {{
                        type: 'line',
                        label: 'Cumulative %',
                        data: cumulative,
                        borderColor: '#7c3aed',
                        borderDash: [5, 5],
                        pointRadius: 4,
                        pointBackgroundColor: '#7c3aed',
                        fill: false,
                        tension: 0.3
                    }},
                    {{
                        type: 'bar',
                        label: 'Grade %',
                        data: pcts,
                        backgroundColor: '#4361ee',
                        borderRadius: 4
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'top' }}
                }},
                scales: {{
                    y: {{ beginAtZero: true, max: 105 }}
                }}
            }}
        }});
    </script>
</body>
</html>'''
        
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)


def main():
    root = tk.Tk()
    app = ExcelAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
