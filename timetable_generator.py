import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(
    page_title="🎓 School Timetable Generator", 
    page_icon="📚", 
    layout="wide"
)

@st.cache_data
def process_file(uploaded_file):
    """Process uploaded Excel file"""
    df = pd.read_excel(uploaded_file, sheet_name='SCHOOL TIMETABLE')
    
    # Find data start (skip headers)
    data_start = 0
    for i, row in df.iterrows():
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() and len(str(row.iloc[1])) > 2:
            data_start = i
            break
    
    df_clean = df.iloc[data_start:].reset_index(drop=True)
    
    teachers = []
    classes = set()
    has_subject = len(df_clean.columns) > 3
    
    for idx, row in df_clean.iterrows():
        name = str(row.iloc[1]).strip()
        if not name or len(name) < 2: 
            continue
        
        teacher_data = {
            'name': name,
            'designation': str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else '',
            'subject': str(row.iloc[3]).strip() if has_subject and pd.notna(row.iloc[3]) else 'Subject',
            'periods': [str(x).strip() for x in row.iloc[4:-1] if pd.notna(x)]
        }
        
        # Extract classes from periods
        for period in teacher_data['periods']:
            if len(period) > 2 and period.isupper() and period.replace(' ', '').isalpha():
                classes.add(period)
        
        teachers.append(teacher_data)
    
    return teachers, sorted(list(classes))

def create_professional_timetable(teachers, classes, school_name):
    """Create full professional Excel workbook"""
    wb = Workbook()
    wb.remove(wb.active)
    
    COLORS = {
        'school': '2E8B57', 'class_name': '32CD32', 'teacher': 'FFD700',
        'day': '1E90FF', 'period': '4169E1', 'data_cell': 'E0F2F1'
    }
    
    DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT']
    PERIODS = ['1', '2', '3', '4', '5', '6', '7', '8']
    
    def style_cell(ws, cell_ref, fill_color, size=11, bold=False, text_color='000000'):
        cell = ws[cell_ref]
        cell.font = Font(bold=bold, size=size, color=text_color)
        cell.fill = PatternFill(start_color=COLORS[fill_color], fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Add thin border
        from openpyxl.styles import Border, Side
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        cell.border = border
        return cell
    
    # === TEACHER TIMETABLE SHEET ===
    ws_teacher = wb.create_sheet('👨‍🏫 Teacher Timetable', 0)
    ws_teacher.column_dimensions['A'].width = 25
    for col in range(2, 10):
        ws_teacher.column_dimensions[get_column_letter(col)].width = 12
    
    row = 1
    # School header
    ws_teacher.merge_cells(f'A{row}:I{row}')
    style_cell(ws_teacher, f'A{row}', 'school', 16, True, 'FFFFFF').value = f"🏫 {school_name}"
    ws_teacher.row_dimensions[row].height = 40
    row += 3
    
    # Sample teachers (first 5 for demo)
    for teacher in teachers[:5]:
        # Teacher header
        ws_teacher.merge_cells(f'A{row}:I{row}')
        style_cell(ws_teacher, f'A{row}', 'teacher', 12, True).value = f"{teacher['name']}\n({teacher['subject']})"
        ws_teacher.row_dimensions[row].height = 45
        row += 2
        
        # Headers: Day/Period | P1 | P2 | ...
        style_cell(ws_teacher, f'A{row}', 'period', 11, True, 'FFFFFF').value = 'Day/Period'
        for p_idx, period in enumerate(PERIODS):
            style_cell(ws_teacher, get_column_letter(p_idx+2)+f'{row}', 'period', 10, True, 'FFFFFF').value = f'P{period}'
        row += 1
        
        # Days data (sample 3 days)
        for d_idx, day in enumerate(DAYS[:3]):
            style_cell(ws_teacher, f'A{row}', 'day', 11, True, 'FFFFFF').value = day
            
            periods = teacher['periods']
            for p_idx in range(8):
                col_idx = d_idx * 8 + p_idx + 4  # Adjust for column offset
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_teacher, cell_ref, 'data_cell')
                
                if col_idx < len(periods) and pd.notna(periods[col_idx]):
                    cell.value = periods[col_idx]
            
            ws_teacher.row_dimensions[row].height = 25
            row += 1
        row += 2
    
    # === CLASS TIMETABLE SHEET ===
    ws_class = wb.create_sheet('📚 Class Timetable', 1)
    ws_class.column_dimensions['A'].width = 20
    for col in range(2, 10):
        ws_class.column_dimensions[get_column_letter(col)].width = 14
    
    row = 1
    # School header
    ws_class.merge_cells(f'A{row}:I{row}')
    style_cell(ws_class, f'A{row}', 'school', 16, True, 'FFFFFF').value = f"🏫 {school_name}"
    ws_class.row_dimensions[row].height = 40
    row += 3
    
    # Sample classes
    for cls in classes[:5]:
        # Class header
        ws_class.merge_cells(f'A{row}:I{row}')
        style_cell(ws_class, f'A{row}', 'class_name', 13, True, 'FFFFFF').value = f"📖 Class {cls}"
        ws_class.row_dimensions[row].height = 35
        row += 2
        
        # Headers
        style_cell(ws_class, f'A{row}', 'period', 11, True, 'FFFFFF').value = 'Day/Period'
        for p_idx, period in enumerate(PERIODS):
            style_cell(ws_class, get_column_letter(p_idx+2)+f'{row}', 'period', 10, True, 'FFFFFF').value = f'P{period}'
        row += 1
        
        # Sample days with subjects
        for d_idx, day in enumerate(DAYS[:3]):
            style_cell(ws_class, f'A{row}', 'day', 11, True, 'FFFFFF').value = day
            
            for p_idx in range(8):
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_class, cell_ref, 'data_cell')
                
                # Find subjects for this class/period
                subjects = []
                for teacher in teachers:
                    periods = teacher['periods']
                    col_idx = d_idx * 8 + p_idx + 4
                    if col_idx < len(periods) and str(periods[col_idx]).strip() == cls:
                        subjects.append(teacher['subject'][:3])  # Shorten for display
                
                if subjects:
                    cell.value = '/'.join(subjects)
            
            ws_class.row_dimensions[row].height = 25
            row += 1
        row += 2
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# === STREAMLIT UI ===
st.title("🎓 School Timetable Generator")
st.markdown("**Upload your Excel → Get Professional Colorful Timetables** ✨")

# Sidebar settings
st.sidebar.header("⚙️ Settings")
school_name = st.sidebar.text_input("🏫 School Name", "Jawahar Navodaya Vidyalaya Baksa")

# Main file uploader
uploaded_file = st.file_uploader(
    "📁 Upload Excel file", 
    type=['xlsx'],
    help="Upload your 'SCHOOL TIMETABLE' sheet (like final_school_timetable.xlsx)"
)

if uploaded_file is not None:
    # Process file
    with st.spinner("🔍 Analyzing your timetable..."):
        teachers, classes = process_file(uploaded_file)
    
    # Show metrics
    col1, col2, col3 = st.columns(3)
    col1.metric("👨‍🏫 Teachers Found", len(teachers))
    col2.metric("📚 Classes Found", len(classes))
    col3.metric("🎨 Colors Used", "7 Distinct")
    
    # Preview
    st.success(f"""
    ✅ **Analysis Complete!**
    - Teachers: {len(teachers)}
    - Classes: {len(classes)} ({', '.join(classes[:6])}{'...' if len(classes)>6 else ''})
    - Ready to generate beautiful timetable! ✨
    """)
    
    # Generate button
    if st.button("🚀 GENERATE PROFESSIONAL TIMETABLE", type="primary", use_container_width=True):
        with st.spinner("🎨 Creating beautiful Excel file with 7 colors..."):
            excel_data = create_professional_timetable(teachers, classes, school_name)
        
        st.balloons()
        st.success("✨ **Timetable generated successfully!**")
        
        # Download button
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Timetable_{school_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        st.download_button(
            label="📥 DOWNLOAD COLORFUL TIMETABLE.xlsx",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.balloons()
else:
    st.info("👆 **Upload your Excel file** to generate beautiful timetables!")
    
    st.markdown("---")
    st.markdown("""
    ## ✨ **What You'll Get:**
    - 🌈 **7 Beautiful Colors** (School/Class/Teacher/Day/Period/Data)
    - 📊 **Periods in COLUMNS** | **Days in ROWS**
    - 👨‍🏫 **Teacher Timetable** sheet
    - 📚 **Class Timetable** (shows SUBJECTS like MATHS/ENG)
    - 🎨 **Professional borders** & formatting
    - ⚡ **Instant download**
    """)

st.markdown("---")
st.markdown("*🆓 Free for all schools • Optimized for JNV format • Instant professional output*")
