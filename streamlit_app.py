import streamlit as st
import xlsxwriter
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Facility Management System Generator",
    page_icon="🏫",
    layout="wide"
)

# Title
st.title("🏫 Facility Management System Generator")
st.markdown("### Generate Your Complete Excel Workbook")
st.markdown("---")

# Information
st.info("""
**This tool generates a complete Facility Management System with:**
- ✅ Dashboard with Live KPIs
- ✅ Complaint Register (Auto-tracking)
- ✅ Maintenance Log (Auto-calculations)
- ✅ Preventive Maintenance Schedule
- ✅ Inventory Tracker
- ✅ Generator & Diesel Log
- ✅ Vendor/Technician Tracker
- ✅ Daily Report Sheet
- ✅ Settings & Controls
""")

st.markdown("---")

# School details input
st.subheader("📝 Customize Your System (Optional)")

col1, col2 = st.columns(2)

with col1:
    school_name = st.text_input("School Name", "School Name")
    facility_officer = st.text_input("Facility Officer Name", "Facility Officer")

with col2:
    location = st.text_input("Location", "Lagos, Nigeria")
    year = st.text_input("Academic Year", "2024/2025")

st.markdown("---")

# Generate button
if st.button("🚀 Generate Excel Workbook", type="primary", use_container_width=True):
    
    with st.spinner("⏳ Creating your Facility Management System..."):
        
        # Create workbook in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        # ═══════════════════════════════════════════════════════
        # FORMAT DEFINITIONS
        # ═══════════════════════════════════════════════════════
        
        header_base = {
            'bold': True,
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        }
        
        format_dashboard = workbook.add_format({**header_base, 'bg_color': '#1F4E78', 'font_size': 11})
        format_complaint = workbook.add_format({**header_base, 'bg_color': '#C00000'})
        format_maintenance = workbook.add_format({**header_base, 'bg_color': '#70AD47'})
        format_schedule = workbook.add_format({**header_base, 'bg_color': '#7030A0'})
        format_inventory = workbook.add_format({**header_base, 'bg_color': '#00B0F0'})
        format_generator = workbook.add_format({**header_base, 'bg_color': '#FF6600'})
        format_vendor = workbook.add_format({**header_base, 'bg_color': '#833C0C'})
        format_daily = workbook.add_format({**header_base, 'bg_color': '#002060'})
        format_settings = workbook.add_format({**header_base, 'bg_color': '#595959'})
        
        currency_format = workbook.add_format({'num_format': '₦#,##0.00', 'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        time_format = workbook.add_format({'num_format': 'hh:mm AM/PM', 'border': 1})
        number_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
        border_format = workbook.add_format({'border': 1, 'valign': 'top'})
        wrap_format = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
        
        metric_title = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left', 'bg_color': '#D9E1F2', 'border': 1})
        metric_value = workbook.add_format({'font_size': 20, 'bold': True, 'align': 'center', 'border': 1, 'num_format': '#,##0'})
        metric_good = workbook.add_format({'font_size': 20, 'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
        metric_warning = workbook.add_format({'font_size': 20, 'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
        metric_alert = workbook.add_format({'font_size': 20, 'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        
        # ═══════════════════════════════════════════════════════
        # SETTINGS SHEET
        # ═══════════════════════════════════════════════════════
        
        ws_settings = workbook.add_worksheet('Settings')
        ws_settings.set_tab_color('#595959')
        
        settings_data = {
            'Status_List': ['Open', 'In Progress', 'Completed', 'Closed', 'Pending'],
            'Priority_List': ['URGENT', 'High', 'Medium', 'Low'],
            'Location_List': ['Admin Block', 'Primary Block', 'Secondary Block', 'Science Lab', 'Computer Lab', 
                              'Library', 'Assembly Hall', 'Sports Field', 'Generator Room', 'Water Tank Area', 
                              'Staff Room', 'Canteen', 'Toilets', 'Parking Area', 'Main Gate'],
            'Issue_Category': ['Electrical', 'Plumbing', 'AC/Cooling', 'Furniture', 'Doors/Windows', 'Painting',
                               'Roofing', 'Drainage', 'Security', 'Cleaning', 'Generator', 'Water Supply', 'Other'],
            'Maintenance_Type': ['Corrective', 'Preventive', 'Emergency', 'Routine', 'Inspection'],
            'Frequency_List': ['Daily', 'Weekly', 'Bi-Weekly', 'Monthly', 'Quarterly', 'Bi-Annual', 'Annual'],
            'Condition_List': ['Excellent', 'Good', 'Fair', 'Poor', 'Needs Replacement'],
            'Staff_Names': ['Mr. Adeleke', 'Mr. Okafor', 'Mrs. Bello', 'Mr. Ibrahim', 'External Technician'],
            'Vendor_Type': ['Electrician', 'Plumber', 'AC Technician', 'Generator Technician', 'Carpenter',
                            'Painter', 'Diesel Supplier', 'General Maintenance', 'Security', 'Cleaning Service']
        }
        
        col = 0
        for list_name, items in settings_data.items():
            ws_settings.write(0, col, list_name, format_settings)
            ws_settings.set_column(col, col, 20)
            for row, item in enumerate(items, start=1):
                ws_settings.write(row, col, item, border_format)
            col += 1
        
        # ═══════════════════════════════════════════════════════
        # COMPLAINT REGISTER
        # ═══════════════════════════════════════════════════════
        
        cols_complaint = [
            'Complaint ID', 'Date Reported', 'Time', 'Reported By', 'Department', 'Location',
            'Category', 'Complaint Details', 'Priority', 'Assigned To', 'Status',
            'Date Resolved', 'Days Open', 'Alert', 'Action Taken', 'Remarks'
        ]
        
        ws_complaint = workbook.add_worksheet('Complaint_Register')
        ws_complaint.set_tab_color('#C00000')
        ws_complaint.freeze_panes(1, 0)
        ws_complaint.autofilter(0, 0, 0, len(cols_complaint) - 1)
        
        for col_num, col_name in enumerate(cols_complaint):
            ws_complaint.write(0, col_num, col_name, format_complaint)
        
        widths = [15, 12, 10, 20, 20, 18, 15, 40, 12, 20, 12, 12, 10, 10, 40, 30]
        for col_num, width in enumerate(widths):
            ws_complaint.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_complaint.write_formula(row, 0, f'=IF(B{row+1}="","","CMP-"&TEXT(ROW()-1,"000"))', border_format)
            ws_complaint.write_formula(row, 12, f'=IF(B{row+1}="","",IF(L{row+1}<>"",L{row+1}-B{row+1},TODAY()-B{row+1}))', number_format)
            ws_complaint.write_formula(row, 13, f'=IF(B{row+1}="","",IF(AND(K{row+1}<>"Closed",M{row+1}>3),"OVERDUE","OK"))', border_format)
        
        ws_complaint.data_validation('F2:F101', {'validate': 'list', 'source': '=Settings!$C$2:$C$16'})
        ws_complaint.data_validation('G2:G101', {'validate': 'list', 'source': '=Settings!$D$2:$D$14'})
        ws_complaint.data_validation('I2:I101', {'validate': 'list', 'source': '=Settings!$B$2:$B$5'})
        ws_complaint.data_validation('J2:J101', {'validate': 'list', 'source': '=Settings!$H$2:$H$6'})
        ws_complaint.data_validation('K2:K101', {'validate': 'list', 'source': '=Settings!$A$2:$A$6'})
        
        # ═══════════════════════════════════════════════════════
        # MAINTENANCE LOG
        # ═══════════════════════════════════════════════════════
        
        cols_maint = [
            'Work Order ID', 'Date Reported', 'Type', 'Category', 'Location', 'Description',
            'Priority', 'Assigned To', 'Status', 'Date Started', 'Date Completed',
            'Duration (Days)', 'Parts Used', 'Cost (₦)', 'Vendor', 'Alert', 'Notes'
        ]
        
        ws_maint = workbook.add_worksheet('Maintenance_Log')
        ws_maint.set_tab_color('#70AD47')
        ws_maint.freeze_panes(1, 0)
        ws_maint.autofilter(0, 0, 0, len(cols_maint) - 1)
        
        for col_num, col_name in enumerate(cols_maint):
            ws_maint.write(0, col_num, col_name, format_maintenance)
        
        widths = [15, 12, 15, 15, 18, 40, 12, 20, 12, 12, 12, 12, 30, 15, 20, 10, 40]
        for col_num, width in enumerate(widths):
            ws_maint.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_maint.write_formula(row, 0, f'=IF(B{row+1}="","","WO-"&TEXT(ROW()-1,"0000"))', border_format)
            ws_maint.write_formula(row, 11, f'=IF(K{row+1}="","",IF(J{row+1}="","",K{row+1}-J{row+1}))', number_format)
            ws_maint.write_formula(row, 15, f'=IF(B{row+1}="","",IF(AND(I{row+1}<>"Completed",TODAY()-B{row+1}>5),"DELAYED","OK"))', border_format)
        
        ws_maint.set_column('N:N', 15, currency_format)
        
        ws_maint.data_validation('C2:C101', {'validate': 'list', 'source': '=Settings!$E$2:$E$6'})
        ws_maint.data_validation('D2:D101', {'validate': 'list', 'source': '=Settings!$D$2:$D$14'})
        ws_maint.data_validation('E2:E101', {'validate': 'list', 'source': '=Settings!$C$2:$C$16'})
        ws_maint.data_validation('G2:G101', {'validate': 'list', 'source': '=Settings!$B$2:$B$5'})
        ws_maint.data_validation('H2:H101', {'validate': 'list', 'source': '=Settings!$H$2:$H$6'})
        ws_maint.data_validation('I2:I101', {'validate': 'list', 'source': '=Settings!$A$2:$A$6'})
        
        # ═══════════════════════════════════════════════════════
        # MAINTENANCE SCHEDULE
        # ═══════════════════════════════════════════════════════
        
        cols_schedule = [
            'Task ID', 'Equipment/System', 'Location', 'Task Description', 'Frequency',
            'Last Completed', 'Next Due Date', 'Days Until Due', 'Assigned To',
            'Estimated Cost (₦)', 'Status', 'Actual Cost (₦)', 'Date Completed', 'Notes', 'Alert'
        ]
        
        ws_schedule = workbook.add_worksheet('Maintenance_Schedule')
        ws_schedule.set_tab_color('#7030A0')
        ws_schedule.freeze_panes(1, 0)
        ws_schedule.autofilter(0, 0, 0, len(cols_schedule) - 1)
        
        for col_num, col_name in enumerate(cols_schedule):
            ws_schedule.write(0, col_num, col_name, format_schedule)
        
        widths = [12, 25, 18, 40, 12, 12, 12, 12, 20, 15, 12, 15, 12, 40, 10]
        for col_num, width in enumerate(widths):
            ws_schedule.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_schedule.write_formula(row, 0, f'=IF(B{row+1}="","","PM-"&TEXT(ROW()-1,"000"))', border_format)
            ws_schedule.write_formula(row, 6, f'=IF(F{row+1}="","",IF(E{row+1}="Daily",F{row+1}+1,IF(E{row+1}="Weekly",F{row+1}+7,IF(E{row+1}="Bi-Weekly",F{row+1}+14,IF(E{row+1}="Monthly",EDATE(F{row+1},1),IF(E{row+1}="Quarterly",EDATE(F{row+1},3),IF(E{row+1}="Bi-Annual",EDATE(F{row+1},6),IF(E{row+1}="Annual",EDATE(F{row+1},12),"")))))))', date_format)
            ws_schedule.write_formula(row, 7, f'=IF(G{row+1}="","",G{row+1}-TODAY())', number_format)
            ws_schedule.write_formula(row, 14, f'=IF(G{row+1}="","",IF(H{row+1}<0,"OVERDUE",IF(H{row+1}<=7,"DUE SOON","OK")))', border_format)
        
        ws_schedule.set_column('J:J', 15, currency_format)
        ws_schedule.set_column('L:L', 15, currency_format)
        
        ws_schedule.data_validation('C2:C101', {'validate': 'list', 'source': '=Settings!$C$2:$C$16'})
        ws_schedule.data_validation('E2:E101', {'validate': 'list', 'source': '=Settings!$F$2:$F$8'})
        ws_schedule.data_validation('I2:I101', {'validate': 'list', 'source': '=Settings!$H$2:$H$6'})
        ws_schedule.data_validation('K2:K101', {'validate': 'list', 'source': '=Settings!$A$2:$A$6'})
        
        sample_tasks = [
            ['Generator', 'Generator Room', 'Full service: oil change, filter replacement, spark plug check', 'Monthly', '2024-12-01'],
            ['Main AC Units', 'Admin Block', 'Filter cleaning and refrigerant check', 'Monthly', '2024-12-05'],
            ['Water Tank', 'Roof', 'Complete tank cleaning and disinfection', 'Quarterly', '2024-10-15'],
            ['Fire Extinguishers', 'All Locations', 'Pressure check and inspection', 'Monthly', '2024-12-10'],
            ['External Drains', 'Compound', 'Clear all drains and gutters', 'Weekly', '2024-12-15'],
            ['Emergency Lights', 'All Buildings', 'Test and battery check', 'Monthly', '2024-12-01'],
            ['Water Pumps', 'Pump Room', 'Lubrication and performance check', 'Bi-Weekly', '2024-12-08'],
            ['School Bus', 'Parking Area', 'Service and maintenance check', 'Monthly', '2024-11-30'],
        ]
        
        for i, task in enumerate(sample_tasks, start=1):
            for j, val in enumerate(task):
                ws_schedule.write(i, j + 1, val, border_format)
        
        # ═══════════════════════════════════════════════════════
        # INVENTORY TRACKER
        # ═══════════════════════════════════════════════════════
        
        cols_inventory = [
            'Asset ID', 'Category', 'Item Name', 'Brand/Model', 'Location', 'Serial Number',
            'Quantity', 'Condition', 'Unit Cost (₦)', 'Total Value (₦)', 'Purchase Date',
            'Warranty Expiry', 'Last Service Date', 'Next Service Due', 'Supplier', 'Alert', 'Notes'
        ]
        
        ws_inv = workbook.add_worksheet('Inventory_Tracker')
        ws_inv.set_tab_color('#00B0F0')
        ws_inv.freeze_panes(1, 0)
        ws_inv.autofilter(0, 0, 0, len(cols_inventory) - 1)
        
        for col_num, col_name in enumerate(cols_inventory):
            ws_inv.write(0, col_num, col_name, format_inventory)
        
        widths = [12, 20, 25, 20, 18, 18, 10, 15, 15, 15, 12, 12, 12, 12, 20, 15, 40]
        for col_num, width in enumerate(widths):
            ws_inv.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_inv.write_formula(row, 0, f'=IF(B{row+1}="","","AST-"&TEXT(ROW()-1,"0000"))', border_format)
            ws_inv.write_formula(row, 9, f'=IF(AND(G{row+1}<>"",I{row+1}<>""),G{row+1}*I{row+1},"")', currency_format)
            ws_inv.write_formula(row, 15, f'=IF(B{row+1}="","",IF(OR(H{row+1}="Needs Replacement",H{row+1}="Poor"),"REPLACE",IF(L{row+1}<TODAY(),"WARRANTY EXPIRED","OK")))', border_format)
        
        ws_inv.set_column('I:J', 15, currency_format)
        
        ws_inv.data_validation('B2:B101', {'validate': 'list', 'source': '=Settings!$D$2:$D$14'})
        ws_inv.data_validation('E2:E101', {'validate': 'list', 'source': '=Settings!$C$2:$C$16'})
        ws_inv.data_validation('H2:H101', {'validate': 'list', 'source': '=Settings!$G$2:$G$6'})
        
        # ═══════════════════════════════════════════════════════
        # GENERATOR & DIESEL LOG
        # ═══════════════════════════════════════════════════════
        
        cols_gen = [
            'Date', 'Day', 'Opening Level (L)', 'Diesel Added (L)', 'Cost Per Liter (₦)',
            'Total Cost (₦)', 'Closing Level (L)', 'Fuel Consumed (L)', 'Start Time', 'Stop Time',
            'Runtime (Hours)', 'Efficiency (L/Hr)', 'PHCN Hours', 'Maintenance Done', 'Issues/Faults', 'Remarks'
        ]
        
        ws_gen = workbook.add_worksheet('Generator_Diesel_Log')
        ws_gen.set_tab_color('#FF6600')
        ws_gen.freeze_panes(1, 0)
        ws_gen.autofilter(0, 0, 0, len(cols_gen) - 1)
        
        for col_num, col_name in enumerate(cols_gen):
            ws_gen.write(0, col_num, col_name, format_generator)
        
        widths = [12, 10, 15, 15, 15, 15, 15, 15, 12, 12, 12, 12, 12, 30, 30, 30]
        for col_num, width in enumerate(widths):
            ws_gen.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_gen.write_formula(row, 1, f'=IF(A{row+1}="","",TEXT(A{row+1},"dddd"))', border_format)
            ws_gen.write_formula(row, 5, f'=IF(AND(D{row+1}<>"",E{row+1}<>""),D{row+1}*E{row+1},"")', currency_format)
            ws_gen.write_formula(row, 7, f'=IF(C{row+1}="","",C{row+1}+D{row+1}-G{row+1})', number_format)
            ws_gen.write_formula(row, 10, f'=IF(AND(I{row+1}<>"",J{row+1}<>""),(J{row+1}-I{row+1})*24,"")', number_format)
            ws_gen.write_formula(row, 11, f'=IF(K{row+1}=0,"",IF(H{row+1}="","",H{row+1}/K{row+1}))', number_format)
        
        ws_gen.set_column('E:F', 15, currency_format)
        
        # ═══════════════════════════════════════════════════════
        # VENDOR/TECHNICIAN TRACKER
        # ═══════════════════════════════════════════════════════
        
        cols_vendor = [
            'Vendor ID', 'Company/Name', 'Service Type', 'Contact Person', 'Phone',
            'Email', 'Address', 'Services Offered', 'Rating (1-5)', 'Jobs Completed',
            'Total Spent (₦)', 'Last Used', 'Average Response Time', 'Performance', 'Contract Status', 'Notes'
        ]
        
        ws_vendor = workbook.add_worksheet('Vendor_Technician_Tracker')
        ws_vendor.set_tab_color('#833C0C')
        ws_vendor.freeze_panes(1, 0)
        ws_vendor.autofilter(0, 0, 0, len(cols_vendor) - 1)
        
        for col_num, col_name in enumerate(cols_vendor):
            ws_vendor.write(0, col_num, col_name, format_vendor)
        
        widths = [12, 25, 20, 20, 15, 25, 30, 40, 12, 12, 15, 12, 20, 15, 15, 40]
        for col_num, width in enumerate(widths):
            ws_vendor.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_vendor.write_formula(row, 0, f'=IF(B{row+1}="","","VEN-"&TEXT(ROW()-1,"000"))', border_format)
            ws_vendor.write_formula(row, 13, f'=IF(I{row+1}="","",IF(I{row+1}>=4,"Excellent",IF(I{row+1}>=3,"Good",IF(I{row+1}>=2,"Fair","Poor"))))', border_format)
        
        ws_vendor.set_column('K:K', 15, currency_format)
        
        ws_vendor.data_validation('C2:C101', {'validate': 'list', 'source': '=Settings!$I$2:$I$11'})
        ws_vendor.data_validation('I2:I101', {'validate': 'integer', 'criteria': 'between', 'minimum': 1, 'maximum': 5})
        ws_vendor.data_validation('O2:O101', {'validate': 'list', 'source': ['Active', 'Inactive', 'Contract', 'Blacklisted']})
        
        # ═══════════════════════════════════════════════════════
        # DAILY REPORT
        # ═══════════════════════════════════════════════════════
        
        cols_daily = [
            'Date', 'Day', 'Officer Name', 'Weather', 'PHCN Hours', 'Generator Hours',
            'Diesel Used (L)', 'Water Supply Status', 'Complaints Received', 'Complaints Resolved',
            'Maintenance Jobs Done', 'Technicians On-Site', 'Inspections Done', 'Safety Issues',
            'Materials Used', 'Total Cost (₦)', 'Pending Issues', 'Recommendations', 'Overall Status'
        ]
        
        ws_daily = workbook.add_worksheet('Daily_Report')
        ws_daily.set_tab_color('#002060')
        ws_daily.freeze_panes(1, 0)
        ws_daily.autofilter(0, 0, 0, len(cols_daily) - 1)
        
        for col_num, col_name in enumerate(cols_daily):
            ws_daily.write(0, col_num, col_name, format_daily)
        
        widths = [12, 10, 20, 15, 12, 12, 12, 20, 15, 15, 18, 20, 15, 30, 30, 15, 40, 40, 15]
        for col_num, width in enumerate(widths):
            ws_daily.set_column(col_num, col_num, width)
        
        for row in range(1, 101):
            ws_daily.write_formula(row, 1, f'=IF(A{row+1}="","",TEXT(A{row+1},"dddd"))', border_format)
            ws_daily.write_formula(row, 18, f'=IF(A{row+1}="","",IF(N{row+1}<>"","Critical",IF(I{row+1}>5,"Busy",IF(K{row+1}>3,"Active","Normal"))))', border_format)
        
        ws_daily.set_column('P:P', 15, currency_format)
        
        ws_daily.data_validation('D2:D101', {'validate': 'list', 'source': ['Sunny', 'Rainy', 'Cloudy', 'Stormy']})
        ws_daily.data_validation('H2:H101', {'validate': 'list', 'source': ['Stable', 'Intermittent', 'Low Pressure', 'No Water']})
        
        # ═══════════════════════════════════════════════════════
        # DASHBOARD
        # ═══════════════════════════════════════════════════════
        
        ws_dash = workbook.add_worksheet('Dashboard')
        ws_dash.set_tab_color('#1F4E78')
        
        workbook.worksheets_objs.remove(ws_dash)
        workbook.worksheets_objs.insert(0, ws_dash)
        
        ws_dash.set_column('A:A', 3)
        ws_dash.set_column('B:B', 30)
        ws_dash.set_column('C:C', 20)
        ws_dash.set_column('D:D', 3)
        ws_dash.set_column('E:E', 30)
        ws_dash.set_column('F:F', 20)
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': '#1F4E78',
            'align': 'center',
            'valign': 'vcenter'
        })
        ws_dash.merge_range('B2:F2', f'🏫 {school_name.upper()} - FACILITY MANAGEMENT DASHBOARD', title_format)
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'center',
            'italic': True,
            'font_color': '#595959'
        })
        ws_dash.merge_range('B3:F3', f'Real-Time Overview | {location} | {year}', subtitle_format)
        
        section_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'left'
        })
        
        ws_dash.merge_range('B5:C5', '📋 COMPLAINTS OVERVIEW', section_format)
        ws_dash.write('B6', 'Total Complaints:', metric_title)
        ws_dash.write_formula('C6', '=COUNTA(Complaint_Register!B:B)-1', metric_value)
        ws_dash.write('B7', 'Open Complaints:', metric_title)
        ws_dash.write_formula('C7', '=COUNTIFS(Complaint_Register!K:K,"Open",Complaint_Register!B:B,"<>"&"")', metric_warning)
        ws_dash.write('B8', 'Overdue Complaints:', metric_title)
        ws_dash.write_formula('C8', '=COUNTIF(Complaint_Register!N:N,"OVERDUE")', metric_alert)
        ws_dash.write('B9', 'Resolved This Month:', metric_title)
        ws_dash.write_formula('C9', '=COUNTIFS(Complaint_Register!L:L,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Complaint_Register!K:K,"Closed")', metric_good)
        ws_dash.write('B10', 'Avg Resolution Time (Days):', metric_title)
        ws_dash.write_formula('C10', '=IFERROR(AVERAGEIF(Complaint_Register!K:K,"Closed",Complaint_Register!M:M),0)', number_format)
        
        ws_dash.merge_range('E5:F5', '🔧 MAINTENANCE STATUS', section_format)
        ws_dash.write('E6', 'Total Work Orders:', metric_title)
        ws_dash.write_formula('F6', '=COUNTA(Maintenance_Log!B:B)-1', metric_value)
        ws_dash.write('E7', 'In Progress:', metric_title)
        ws_dash.write_formula('F7', '=COUNTIFS(Maintenance_Log!I:I,"In Progress",Maintenance_Log!B:B,"<>"&"")', metric_warning)
        ws_dash.write('E8', 'Delayed Jobs:', metric_title)
        ws_dash.write_formula('F8', '=COUNTIF(Maintenance_Log!P:P,"DELAYED")', metric_alert)
        ws_dash.write('E9', 'Completed This Month:', metric_title)
        ws_dash.write_formula('F9', '=COUNTIFS(Maintenance_Log!K:K,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Maintenance_Log!I:I,"Completed")', metric_good)
        ws_dash.write('E10', 'Total Spent This Month (₦):', metric_title)
        ws_dash.write_formula('F10', '=SUMIFS(Maintenance_Log!N:N,Maintenance_Log!K:K,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', currency_format)
        
        ws_dash.merge_range('B12:C12', '📅 PREVENTIVE MAINTENANCE', section_format)
        ws_dash.write('B13', 'Total Scheduled Tasks:', metric_title)
        ws_dash.write_formula('C13', '=COUNTA(Maintenance_Schedule!B:B)-1', metric_value)
        ws_dash.write('B14', 'Overdue Tasks:', metric_title)
        ws_dash.write_formula('C14', '=COUNTIF(Maintenance_Schedule!O:O,"OVERDUE")', metric_alert)
        ws_dash.write('B15', 'Due This Week:', metric_title)
        ws_dash.write_formula('C15', '=COUNTIF(Maintenance_Schedule!O:O,"DUE SOON")', metric_warning)
        ws_dash.write('B16', 'Completion Rate:', metric_title)
        ws_dash.write_formula('C16', '=IFERROR(COUNTIFS(Maintenance_Schedule!K:K,"Completed")/(COUNTA(Maintenance_Schedule!B:B)-1),0)', percent_format)
        
        ws_dash.merge_range('E12:F12', '⚡ GENERATOR & DIESEL', section_format)
        ws_dash.write('E13', 'Total Diesel Used This Month (L):', metric_title)
        ws_dash.write_formula('F13', '=SUMIFS(Generator_Diesel_Log!H:H,Generator_Diesel_Log!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', number_format)
        ws_dash.write('E14', 'Total Diesel Cost This Month (₦):', metric_title)
        ws_dash.write_formula('F14', '=SUMIFS(Generator_Diesel_Log!F:F,Generator_Diesel_Log!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', currency_format)
        ws_dash.write('E15', 'Avg Efficiency (L/Hr):', metric_title)
        ws_dash.write_formula('F15', '=IFERROR(AVERAGE(Generator_Diesel_Log!L:L),0)', number_format)
        ws_dash.write('E16', 'Total Runtime This Month (Hrs):', metric_title)
        ws_dash.write_formula('F16', '=SUMIFS(Generator_Diesel_Log!K:K,Generator_Diesel_Log!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', number_format)
        
        ws_dash.merge_range('B18:C18', '📦 INVENTORY ALERTS', section_format)
        ws_dash.write('B19', 'Total Assets:', metric_title)
        ws_dash.write_formula('C19', '=COUNTA(Inventory_Tracker!B:B)-1', metric_value)
        ws_dash.write('B20', 'Items Needing Replacement:', metric_title)
        ws_dash.write_formula('C20', '=COUNTIF(Inventory_Tracker!P:P,"REPLACE")', metric_alert)
        ws_dash.write('B21', 'Warranty Expired:', metric_title)
        ws_dash.write_formula('C21', '=COUNTIF(Inventory_Tracker!P:P,"WARRANTY EXPIRED")', metric_warning)
        ws_dash.write('B22', 'Total Asset Value (₦):', metric_title)
        ws_dash.write_formula('C22', '=SUM(Inventory_Tracker!J:J)', currency_format)
        
        ws_dash.merge_range('E18:F18', '⚠️ CRITICAL ALERTS', section_format)
        ws_dash.write('E19', 'URGENT Priorities:', metric_title)
        ws_dash.write_formula('F19', '=COUNTIF(Complaint_Register!I:I,"URGENT")+COUNTIF(Maintenance_Log!G:G,"URGENT")', metric_alert)
        ws_dash.write('E20', 'Issues Older Than 7 Days:', metric_title)
        ws_dash.write_formula('F20', '=COUNTIFS(Complaint_Register!M:M,">7",Complaint_Register!K:K,"<>Closed")', metric_alert)
        ws_dash.write('E21', 'Vendors on Blacklist:', metric_title)
        ws_dash.write_formula('F21', '=COUNTIF(Vendor_Technician_Tracker!O:O,"Blacklisted")', metric_alert)
        
        ws_dash.merge_range('B24:F24', '📖 HOW TO USE THIS SYSTEM', section_format)
        
        instructions_format = workbook.add_format({
            'text_wrap': True,
            'border': 1,
            'align': 'left',
            'valign': 'top',
            'font_size': 9
        })
        
        instructions_text = [
            ("1. Complaint Register", "Log all complaints here. System auto-tracks days open and alerts you if overdue."),
            ("2. Maintenance Log", "Record all repairs and work orders. Tracks duration and costs automatically."),
            ("3. Maintenance Schedule", "Preventive tasks auto-calculate next due dates. Check weekly for overdue items."),
            ("4. Inventory Tracker", "Track all school assets. System alerts for replacements and expired warranties."),
            ("5. Generator Log", "Daily diesel tracking with auto-calculated efficiency and consumption."),
            ("6. Vendor Tracker", "Database of all service providers with performance ratings."),
            ("7. Daily Report", "End-of-day summary sheet for management reporting."),
            ("8. Settings", "Control all dropdown lists from this sheet. Update as needed."),
        ]
        
        row = 25
        for sheet_name, description in instructions_text:
            ws_dash.write(row, 1, sheet_name, workbook.add_format({'bold': True, 'border': 1}))
            ws_dash.merge_range(row, 2, row, 5, description, instructions_format)
            row += 1
        
        ws_dash.write('B34', 'System Generated:', workbook.add_format({'italic': True}))
        ws_dash.write('C34', datetime.now().strftime('%d/%m/%Y %I:%M %p'), workbook.add_format({'italic': True}))
        
        ws_dash.write('B35', 'Facility Officer:', workbook.add_format({'italic': True}))
        ws_dash.write('C35', facility_officer, workbook.add_format({'italic': True}))
        
        # Close and prepare download
        workbook.close()
        output.seek(0)
        
        # Success message
        st.success("✅ Your Facility Management System has been generated successfully!")
        
        # Download button
        st.download_button(
            label="📥 Download Excel Workbook",
            data=output.getvalue(),
            file_name=f"Facility_Management_System_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        # Features summary
        st.markdown("---")
        st.markdown("### ✅ Your System Includes:")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            - ✅ Live Dashboard
            - ✅ Complaint Register
            - ✅ Maintenance Log
            """)
        
        with col2:
            st.markdown("""
            - ✅ Preventive Schedule
            - ✅ Inventory Tracker
            - ✅ Generator Log
            """)
        
        with col3:
            st.markdown("""
            - ✅ Vendor Tracker
            - ✅ Daily Report
            - ✅ Smart Alerts
            """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p><strong>Facility Management System Generator</strong></p>
    <p>Professional Excel Workbook for School Operations</p>
    <p style='font-size: 0.8em;'>All formulas and automations included • No additional software required</p>
</div>
""", unsafe_allow_html=True)
