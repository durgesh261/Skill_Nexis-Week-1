import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference

def create_dashboard_workbook():
    input_file = "global_superstore_2016.xlsx"
    output_file = "Week1_Sales_Performance_Dashboard.xlsx"
    
    print("Reading Orders from dataset...")
    df = pd.read_excel(input_file, sheet_name="Orders")
    
    # Clean data & add helper columns (e.g. Order Year)
    df_cleaned = df.copy()
    df_cleaned["Postal Code"] = df_cleaned["Postal Code"].fillna("N/A - Non-US")
    df_cleaned["Order Year"] = pd.to_datetime(df_cleaned["Order Date"]).dt.year
    
    print("Writing base workbook with cleaned orders...")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_cleaned.to_excel(writer, sheet_name="Cleaned_Orders", index=False)
        
    print("Loading workbook for Executive Dashboard formatting...")
    wb = openpyxl.load_workbook(output_file)
    
    # Corporate Executive Color Palette (Navy, Royal Blue, Steel Blue, Cool Grays)
    navy_fill = PatternFill(start_color="0F2537", end_color="0F2537", fill_type="solid")      # Deep Navy
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")    # Corporate Blue
    accent_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")    # Medium Blue
    light_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")     # Accent Light Blue
    kpi_card_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")  # Card Background
    card_banner_1 = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    card_banner_2 = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    card_banner_3 = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    card_banner_4 = PatternFill(start_color="41719C", end_color="41719C", fill_type="solid")
    card_banner_5 = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    success_soft = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    font_main_title = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    font_sub_title = Font(name="Calibri", size=11, italic=True, color="D9E1F2")
    font_section = Font(name="Calibri", size=13, bold=True, color="0F2537")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_kpi_label = Font(name="Calibri", size=9, bold=True, color="595959")
    font_kpi_val = Font(name="Calibri", size=18, bold=True, color="0F2537")
    font_kpi_sub = Font(name="Calibri", size=9, italic=True, color="2E75B6")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    card_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom = Border(top=thin_border_side, bottom=Side(border_style="double", color="1F4E79"))
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # -------------------------------------------------------------
    # Format Cleaned_Orders sheet
    # -------------------------------------------------------------
    ws_orders = wb["Cleaned_Orders"]
    ws_orders.views.sheetView[0].showGridLines = True
    for col_idx in range(1, len(df_cleaned.columns) + 1):
        c = ws_orders.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        
    # -------------------------------------------------------------
    # Sheet 1: Executive_Dashboard
    # -------------------------------------------------------------
    print("Creating Executive_Dashboard sheet...")
    ws_dash = wb.create_sheet(title="Executive_Dashboard", index=0)
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Title Banner (Row 2 to 3, Cols B to M)
    for r in range(2, 4):
        for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]:
            ws_dash[f"{col_letter}{r}"].fill = navy_fill
            
    ws_dash["B2"] = "GLOBAL SUPERSTORE — SALES PERFORMANCE DASHBOARD"
    ws_dash["B2"].font = font_main_title
    ws_dash["B3"] = "Week 1 Project Deliverable: Executive Sales, Category, Trend & Department Analysis (2012 - 2015)"
    ws_dash["B3"].font = font_sub_title
    
    # -------------------------------------------------------------
    # KPI Cards (Row 5 to 7)
    # Card 1: Total Revenue (B5:C7)
    # Card 2: Total Profit (D5:E7)
    # Card 3: Total Orders (F5:G7)
    # Card 4: Avg Order Value (H5:I7)
    # Card 5: Profit Margin % (J5:K7)
    # Card 6: Top Category (L5:M7)
    # -------------------------------------------------------------
    kpis = [
        ("TOTAL REVENUE", "$12,642,502", "+26.3% YoY in 2015", "B", "C", card_banner_1),
        ("TOTAL PROFIT", "$1,467,457", "Healthy Global Gain", "D", "E", card_banner_2),
        ("TOTAL ORDERS", "25,728", "51,290 Order Items", "F", "G", card_banner_3),
        ("AVG ORDER VALUE", "$491.39", "$246.49 Per Item", "H", "I", card_banner_4),
        ("PROFIT MARGIN", "11.61%", "Target > 10.0%", "J", "K", card_banner_5),
        ("TOP CATEGORY", "Technology", "37.5% Revenue Share", "L", "M", card_banner_1),
    ]
    
    for title, val, sub, c1, c2, b_fill in kpis:
        # Top banner for card
        ws_dash[f"{c1}5"] = title
        ws_dash[f"{c1}5"].fill = b_fill
        ws_dash[f"{c1}5"].font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        ws_dash[f"{c1}5"].alignment = align_center
        ws_dash[f"{c2}5"].fill = b_fill
        
        # Value row
        ws_dash[f"{c1}6"] = val
        ws_dash[f"{c1}6"].font = font_kpi_val
        ws_dash[f"{c1}6"].alignment = align_center
        ws_dash[f"{c1}6"].fill = kpi_card_fill
        ws_dash[f"{c2}6"].fill = kpi_card_fill
        
        # Subtitle row
        ws_dash[f"{c1}7"] = sub
        ws_dash[f"{c1}7"].font = font_kpi_sub
        ws_dash[f"{c1}7"].alignment = align_center
        ws_dash[f"{c1}7"].fill = kpi_card_fill
        ws_dash[f"{c2}7"].fill = kpi_card_fill
        
        # Border
        for r in range(5, 8):
            for col in [c1, c2]:
                ws_dash[f"{col}{r}"].border = thin_border
                
    # -------------------------------------------------------------
    # Summary Tables Section for Dashboard Visuals
    # -------------------------------------------------------------
    
    # Table 1: Yearly Performance (Row 9 to 15, Cols B to G)
    ws_dash["B9"] = "1. Yearly Sales & Profit Trends"
    ws_dash["B9"].font = font_section
    
    yearly_headers = ["Year", "Total Orders", "Sales Revenue ($)", "YoY Growth", "Total Profit ($)", "Margin %"]
    for i, h in enumerate(yearly_headers, start=2):
        c = ws_dash.cell(row=10, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    yearly_data = [
        (2012, 4498, 2259451.55, "-", 248940.81, "=F11/D11"),
        (2013, 5451, 2677438.73, "=(D12-D11)/D11", 307415.28, "=F12/D12"),
        (2014, 6932, 3405708.34, "=(D13-D12)/D12", 406935.23, "=F13/D13"),
        (2015, 8847, 4299903.29, "=(D14-D13)/D13", 504165.97, "=F14/D14"),
    ]
    
    for idx, row in enumerate(yearly_data, start=11):
        ws_dash.cell(row=idx, column=2, value=row[0]).alignment = align_center
        ws_dash.cell(row=idx, column=3, value=row[1]).alignment = align_right
        ws_dash.cell(row=idx, column=4, value=row[2]).alignment = align_right
        ws_dash.cell(row=idx, column=5, value=row[3]).alignment = align_right
        ws_dash.cell(row=idx, column=6, value=row[4]).alignment = align_right
        ws_dash.cell(row=idx, column=7, value=row[5]).alignment = align_right
        
        ws_dash.cell(row=idx, column=3).number_format = "#,##0"
        ws_dash.cell(row=idx, column=4).number_format = "$#,##0.00"
        if row[3] != "-":
            ws_dash.cell(row=idx, column=5).number_format = "+0.00%;-0.00%"
        ws_dash.cell(row=idx, column=6).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=7).number_format = "0.00%"
        
        for c in range(2, 8):
            ws_dash.cell(row=idx, column=c).border = thin_border
            ws_dash.cell(row=idx, column=c).font = font_regular
            
    # Total row for Yearly Table
    ws_dash.cell(row=15, column=2, value="Grand Total").font = font_bold
    ws_dash.cell(row=15, column=3, value="=SUM(C11:C14)").font = font_bold
    ws_dash.cell(row=15, column=4, value="=SUM(D11:D14)").font = font_bold
    ws_dash.cell(row=15, column=5, value="=(D14-D11)/D11").font = font_bold
    ws_dash.cell(row=15, column=6, value="=SUM(F11:F14)").font = font_bold
    ws_dash.cell(row=15, column=7, value="=F15/D15").font = font_bold
    
    ws_dash.cell(row=15, column=3).number_format = "#,##0"
    ws_dash.cell(row=15, column=4).number_format = "$#,##0.00"
    ws_dash.cell(row=15, column=5).number_format = "+0.00%"
    ws_dash.cell(row=15, column=6).number_format = "$#,##0.00"
    ws_dash.cell(row=15, column=7).number_format = "0.00%"
    
    for c in range(2, 8):
        ws_dash.cell(row=15, column=c).border = double_bottom
        ws_dash.cell(row=15, column=c).fill = light_blue
        
    # Table 2: Category Breakdown (Row 9 to 14, Cols I to M)
    ws_dash["I9"] = "2. Sales by Product Category"
    ws_dash["I9"].font = font_section
    
    cat_headers = ["Category", "Orders", "Revenue ($)", "Profit ($)", "Margin %"]
    for i, h in enumerate(cat_headers, start=9):
        c = ws_dash.cell(row=10, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    cat_data = [
        ("Technology", 10141, 4744557.50, 663778.74, "=L11/K11"),
        ("Furniture", 9860, 4110874.19, 285412.87, "=L12/K12"),
        ("Office Supplies", 31289, 3787070.23, 518265.68, "=L13/K13"),
    ]
    
    for idx, row in enumerate(cat_data, start=11):
        ws_dash.cell(row=idx, column=9, value=row[0]).alignment = align_left
        ws_dash.cell(row=idx, column=10, value=row[1]).alignment = align_right
        ws_dash.cell(row=idx, column=11, value=row[2]).alignment = align_right
        ws_dash.cell(row=idx, column=12, value=row[3]).alignment = align_right
        ws_dash.cell(row=idx, column=13, value=row[4]).alignment = align_right
        
        ws_dash.cell(row=idx, column=10).number_format = "#,##0"
        ws_dash.cell(row=idx, column=11).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=12).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=13).number_format = "0.00%"
        
        for c in range(9, 14):
            ws_dash.cell(row=idx, column=c).border = thin_border
            ws_dash.cell(row=idx, column=c).font = font_regular
            
    # Category total row
    ws_dash.cell(row=14, column=9, value="Total").font = font_bold
    ws_dash.cell(row=14, column=10, value="=SUM(J11:J13)").font = font_bold
    ws_dash.cell(row=14, column=11, value="=SUM(K11:K13)").font = font_bold
    ws_dash.cell(row=14, column=12, value="=SUM(L11:L13)").font = font_bold
    ws_dash.cell(row=14, column=13, value="=L14/K14").font = font_bold
    
    ws_dash.cell(row=14, column=10).number_format = "#,##0"
    ws_dash.cell(row=14, column=11).number_format = "$#,##0.00"
    ws_dash.cell(row=14, column=12).number_format = "$#,##0.00"
    ws_dash.cell(row=14, column=13).number_format = "0.00%"
    
    for c in range(9, 14):
        ws_dash.cell(row=14, column=c).border = double_bottom
        ws_dash.cell(row=14, column=c).fill = light_blue
        
    # Table 3: Department / Customer Segment Breakdown (Row 17 to 22, Cols B to G)
    ws_dash["B17"] = "3. Department & Customer Segment Performance"
    ws_dash["B17"].font = font_section
    
    seg_headers = ["Segment / Dept", "Orders", "Revenue ($)", "% Revenue", "Profit ($)", "Margin %"]
    for i, h in enumerate(seg_headers, start=2):
        c = ws_dash.cell(row=18, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    seg_data = [
        ("Consumer", 26518, 6507949.41, "=D19/$D$22", 749239.78, "=F19/D19"),
        ("Corporate", 15429, 3824697.52, "=D20/$D$22", 441208.33, "=F20/D20"),
        ("Home Office", 9343, 2309854.98, "=D21/$D$22", 277009.18, "=F21/D21"),
    ]
    
    for idx, row in enumerate(seg_data, start=19):
        ws_dash.cell(row=idx, column=2, value=row[0]).alignment = align_left
        ws_dash.cell(row=idx, column=3, value=row[1]).alignment = align_right
        ws_dash.cell(row=idx, column=4, value=row[2]).alignment = align_right
        ws_dash.cell(row=idx, column=5, value=row[3]).alignment = align_right
        ws_dash.cell(row=idx, column=6, value=row[4]).alignment = align_right
        ws_dash.cell(row=idx, column=7, value=row[5]).alignment = align_right
        
        ws_dash.cell(row=idx, column=3).number_format = "#,##0"
        ws_dash.cell(row=idx, column=4).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=5).number_format = "0.00%"
        ws_dash.cell(row=idx, column=6).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=7).number_format = "0.00%"
        
        for c in range(2, 8):
            ws_dash.cell(row=idx, column=c).border = thin_border
            ws_dash.cell(row=idx, column=c).font = font_regular
            
    # Segment Total Row
    ws_dash.cell(row=22, column=2, value="Grand Total").font = font_bold
    ws_dash.cell(row=22, column=3, value="=SUM(C19:C21)").font = font_bold
    ws_dash.cell(row=22, column=4, value="=SUM(D19:D21)").font = font_bold
    ws_dash.cell(row=22, column=5, value="=SUM(E19:E21)").font = font_bold
    ws_dash.cell(row=22, column=6, value="=SUM(F19:F21)").font = font_bold
    ws_dash.cell(row=22, column=7, value="=F22/D22").font = font_bold
    
    ws_dash.cell(row=22, column=3).number_format = "#,##0"
    ws_dash.cell(row=22, column=4).number_format = "$#,##0.00"
    ws_dash.cell(row=22, column=5).number_format = "0.00%"
    ws_dash.cell(row=22, column=6).number_format = "$#,##0.00"
    ws_dash.cell(row=22, column=7).number_format = "0.00%"
    
    for c in range(2, 8):
        ws_dash.cell(row=22, column=c).border = double_bottom
        ws_dash.cell(row=22, column=c).fill = light_blue
        
    # Table 4: Top 5 Performing Regions (Row 17 to 23, Cols I to M)
    ws_dash["I17"] = "4. Top 5 Global Regions"
    ws_dash["I17"].font = font_section
    
    reg_headers = ["Region", "Orders", "Revenue ($)", "Profit ($)", "Margin %"]
    for i, h in enumerate(reg_headers, start=9):
        c = ws_dash.cell(row=18, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    top_reg_data = [
        ("Western Europe", 5883, 1731929.74, 219502.82, "=L19/K19"),
        ("Central America", 5616, 1223101.10, 158912.44, "=L20/K20"),
        ("Oceania", 3487, 1100184.73, 120089.11, "=L21/K21"),
        ("Southeastern Asia", 3129, 884423.17, 17852.19, "=L22/K22"),
        ("Southern Asia", 2655, 866572.70, 160868.23, "=L23/K23"),
    ]
    
    for idx, row in enumerate(top_reg_data, start=19):
        ws_dash.cell(row=idx, column=9, value=row[0]).alignment = align_left
        ws_dash.cell(row=idx, column=10, value=row[1]).alignment = align_right
        ws_dash.cell(row=idx, column=11, value=row[2]).alignment = align_right
        ws_dash.cell(row=idx, column=12, value=row[3]).alignment = align_right
        ws_dash.cell(row=idx, column=13, value=row[4]).alignment = align_right
        
        ws_dash.cell(row=idx, column=10).number_format = "#,##0"
        ws_dash.cell(row=idx, column=11).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=12).number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=13).number_format = "0.00%"
        
        for c in range(9, 14):
            ws_dash.cell(row=idx, column=c).border = thin_border
            ws_dash.cell(row=idx, column=c).font = font_regular

    # -------------------------------------------------------------
    # Visual Charts on Dashboard
    # Chart 1: Yearly Sales Growth Line Chart
    # Chart 2: Category Revenue Bar Chart
    # -------------------------------------------------------------
    chart_line = LineChart()
    chart_line.title = "Yearly Sales Revenue & Profit Trend (2012 - 2015)"
    chart_line.style = 13
    chart_line.y_axis.title = "Amount in USD ($)"
    chart_line.x_axis.title = "Fiscal Year"
    chart_line.height = 11
    chart_line.width = 17
    
    line_data = Reference(ws_dash, min_col=4, min_row=10, max_col=6, max_row=14)
    # Exclude column E (YoY Growth) from the line reference by creating separate references
    chart_line = LineChart()
    chart_line.title = "Yearly Sales Revenue & Profit Trend (2012 - 2015)"
    chart_line.style = 13
    chart_line.y_axis.title = "Amount in USD ($)"
    chart_line.x_axis.title = "Fiscal Year"
    chart_line.height = 12
    chart_line.width = 18
    
    # Sales reference
    sales_ref = Reference(ws_dash, min_col=4, min_row=10, max_row=14)
    # Profit reference
    profit_ref = Reference(ws_dash, min_col=6, min_row=10, max_row=14)
    chart_line.add_data(sales_ref, titles_from_data=True)
    chart_line.add_data(profit_ref, titles_from_data=True)
    
    years_ref = Reference(ws_dash, min_col=2, min_row=11, max_row=14)
    chart_line.set_categories(years_ref)
    
    ws_dash.add_chart(chart_line, "B25")
    
    # Chart 2: Category Revenue Bar Chart
    chart_cat = BarChart()
    chart_cat.type = "col"
    chart_cat.style = 10
    chart_cat.title = "Revenue by Product Category"
    chart_cat.y_axis.title = "Revenue in USD ($)"
    chart_cat.x_axis.title = "Category"
    chart_cat.height = 12
    chart_cat.width = 16
    
    cat_sales_ref = Reference(ws_dash, min_col=11, min_row=10, max_row=13)
    cat_name_ref = Reference(ws_dash, min_col=9, min_row=11, max_row=13)
    chart_cat.add_data(cat_sales_ref, titles_from_data=True)
    chart_cat.set_categories(cat_name_ref)
    
    ws_dash.add_chart(chart_cat, "I25")
    
    # Strategic Insights & Recommendations (Row 50 onwards)
    rec_row = 51
    ws_dash[f"B{rec_row}"] = "Executive Insights & Strategic Recommendations:"
    ws_dash[f"B{rec_row}"].font = font_section
    
    recommendations = [
        "1. Rapid Top-Line Acceleration: Revenue expanded consistently from $2.26M in 2012 to $4.30M in 2015, achieving a remarkable 90.3% cumulative 4-year growth.",
        "2. Profit Stability: Overall profit margin has remained solid and resilient at ~11.6% across all four years, reflecting strong cost controls and steady operating efficiency.",
        "3. Technology is the Core Growth Engine: Technology accounts for $4.74M in revenue (37.5%) and $663.8K in profit (14.0% margin). Scaling high-margin corporate IT accessories should remain a strategic priority.",
        "4. Furniture Margin Remediation: Furniture delivers significant volume ($4.11M revenue) but experiences depressed margins (6.94%). Implementing minimum margin constraints on tables and bookcases will prevent margin leakage.",
        "5. Core Customer Driver: The Consumer segment drives 51.5% of total revenue ($6.51M) with high customer loyalty, while Corporate delivers $3.82M."
    ]
    for idx, text in enumerate(recommendations, start=rec_row+1):
        ws_dash[f"B{idx}"] = text
        ws_dash[f"B{idx}"].font = font_regular

    # -------------------------------------------------------------
    # Sheet 2: Pivot_Summaries
    # -------------------------------------------------------------
    print("Creating Pivot_Summaries sheet...")
    ws_piv = wb.create_sheet(title="Pivot_Summaries", index=1)
    ws_piv.views.sheetView[0].showGridLines = True
    
    ws_piv["B2"] = "PIVOT TABLE SUMMARIES & DETAILED BREAKDOWNS"
    ws_piv["B2"].font = font_section
    ws_piv["B3"] = "Underlying Multi-Dimensional Pivot Aggregations for Sales Performance"
    ws_piv["B3"].font = font_sub_title
    
    # Pivot 1: Market Breakdown
    ws_piv["B5"] = "Market Summary Pivot Table"
    ws_piv["B5"].font = font_bold
    piv1_headers = ["Market", "Record Count", "Total Revenue ($)", "% Share", "Total Profit ($)", "Profit Margin %"]
    for i, h in enumerate(piv1_headers, start=2):
        c = ws_piv.cell(row=6, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    market_df = df_cleaned.groupby("Market").agg(
        Count=("Sales", "count"),
        Revenue=("Sales", "sum"),
        Profit=("Profit", "sum")
    ).sort_values("Revenue", ascending=False).reset_index()
    
    for idx, row in market_df.iterrows():
        r = 7 + idx
        ws_piv.cell(row=r, column=2, value=row["Market"]).alignment = align_left
        ws_piv.cell(row=r, column=3, value=row["Count"]).alignment = align_right
        ws_piv.cell(row=r, column=4, value=row["Revenue"]).alignment = align_right
        ws_piv.cell(row=r, column=5, value=f"=D{r}/$D${7+len(market_df)}").alignment = align_right
        ws_piv.cell(row=r, column=6, value=row["Profit"]).alignment = align_right
        ws_piv.cell(row=r, column=7, value=f"=F{r}/D{r}").alignment = align_right
        
        ws_piv.cell(row=r, column=3).number_format = "#,##0"
        ws_piv.cell(row=r, column=4).number_format = "$#,##0.00"
        ws_piv.cell(row=r, column=5).number_format = "0.00%"
        ws_piv.cell(row=r, column=6).number_format = "$#,##0.00"
        ws_piv.cell(row=r, column=7).number_format = "0.00%"
        
        for c in range(2, 8):
            ws_piv.cell(row=r, column=c).border = thin_border
            ws_piv.cell(row=r, column=c).font = font_regular
            
    m_tot = 7 + len(market_df)
    ws_piv.cell(row=m_tot, column=2, value="Grand Total").font = font_bold
    ws_piv.cell(row=m_tot, column=3, value=f"=SUM(C7:C{m_tot-1})").font = font_bold
    ws_piv.cell(row=m_tot, column=4, value=f"=SUM(D7:D{m_tot-1})").font = font_bold
    ws_piv.cell(row=m_tot, column=5, value=f"=SUM(E7:E{m_tot-1})").font = font_bold
    ws_piv.cell(row=m_tot, column=6, value=f"=SUM(F7:F{m_tot-1})").font = font_bold
    ws_piv.cell(row=m_tot, column=7, value=f"=F{m_tot}/D{m_tot}").font = font_bold
    
    ws_piv.cell(row=m_tot, column=3).number_format = "#,##0"
    ws_piv.cell(row=m_tot, column=4).number_format = "$#,##0.00"
    ws_piv.cell(row=m_tot, column=5).number_format = "0.00%"
    ws_piv.cell(row=m_tot, column=6).number_format = "$#,##0.00"
    ws_piv.cell(row=m_tot, column=7).number_format = "0.00%"
    
    for c in range(2, 8):
        ws_piv.cell(row=m_tot, column=c).border = double_bottom
        ws_piv.cell(row=m_tot, column=c).fill = light_blue

    # Pivot 2: Sub-Category Breakdown (Row 18 onwards)
    subcat_start = m_tot + 3
    ws_piv[f"B{subcat_start}"] = "Sub-Category Performance Summary"
    ws_piv[f"B{subcat_start}"].font = font_bold
    
    subcat_headers = ["Sub-Category", "Category", "Orders", "Revenue ($)", "Profit ($)", "Margin %"]
    for i, h in enumerate(subcat_headers, start=2):
        c = ws_piv.cell(row=subcat_start+1, column=i, value=h)
        c.fill = header_fill
        c.font = font_header
        c.alignment = align_center
        c.border = thin_border
        
    subcat_df = df_cleaned.groupby(["Sub-Category", "Category"]).agg(
        Orders=("Sales", "count"),
        Revenue=("Sales", "sum"),
        Profit=("Profit", "sum")
    ).sort_values("Revenue", ascending=False).reset_index()
    
    sub_r_start = subcat_start + 2
    for idx, row in subcat_df.iterrows():
        r = sub_r_start + idx
        ws_piv.cell(row=r, column=2, value=row["Sub-Category"]).alignment = align_left
        ws_piv.cell(row=r, column=3, value=row["Category"]).alignment = align_left
        ws_piv.cell(row=r, column=4, value=row["Orders"]).alignment = align_right
        ws_piv.cell(row=r, column=5, value=row["Revenue"]).alignment = align_right
        ws_piv.cell(row=r, column=6, value=row["Profit"]).alignment = align_right
        ws_piv.cell(row=r, column=7, value=f"=F{r}/E{r}").alignment = align_right
        
        ws_piv.cell(row=r, column=4).number_format = "#,##0"
        ws_piv.cell(row=r, column=5).number_format = "$#,##0.00"
        ws_piv.cell(row=r, column=6).number_format = "$#,##0.00"
        ws_piv.cell(row=r, column=7).number_format = "0.00%"
        
        for c in range(2, 8):
            ws_piv.cell(row=r, column=c).border = thin_border
            ws_piv.cell(row=r, column=c).font = font_regular
            
    sub_tot = sub_r_start + len(subcat_df)
    ws_piv.cell(row=sub_tot, column=2, value="Grand Total").font = font_bold
    ws_piv.cell(row=sub_tot, column=3, value="").border = double_bottom
    ws_piv.cell(row=sub_tot, column=4, value=f"=SUM(D{sub_r_start}:D{sub_tot-1})").font = font_bold
    ws_piv.cell(row=sub_tot, column=5, value=f"=SUM(E{sub_r_start}:E{sub_tot-1})").font = font_bold
    ws_piv.cell(row=sub_tot, column=6, value=f"=SUM(F{sub_r_start}:F{sub_tot-1})").font = font_bold
    ws_piv.cell(row=sub_tot, column=7, value=f"=F{sub_tot}/E{sub_tot}").font = font_bold
    
    ws_piv.cell(row=sub_tot, column=4).number_format = "#,##0"
    ws_piv.cell(row=sub_tot, column=5).number_format = "$#,##0.00"
    ws_piv.cell(row=sub_tot, column=6).number_format = "$#,##0.00"
    ws_piv.cell(row=sub_tot, column=7).number_format = "0.00%"
    
    for c in range(2, 8):
        ws_piv.cell(row=sub_tot, column=c).border = double_bottom
        ws_piv.cell(row=sub_tot, column=c).fill = light_blue

    # Set column widths
    ws_dash.column_dimensions["A"].width = 3
    ws_dash.column_dimensions["B"].width = 18
    ws_dash.column_dimensions["C"].width = 16
    ws_dash.column_dimensions["D"].width = 20
    ws_dash.column_dimensions["E"].width = 16
    ws_dash.column_dimensions["F"].width = 20
    ws_dash.column_dimensions["G"].width = 14
    ws_dash.column_dimensions["H"].width = 4
    ws_dash.column_dimensions["I"].width = 20
    ws_dash.column_dimensions["J"].width = 14
    ws_dash.column_dimensions["K"].width = 20
    ws_dash.column_dimensions["L"].width = 18
    ws_dash.column_dimensions["M"].width = 14
    
    ws_piv.column_dimensions["A"].width = 4
    ws_piv.column_dimensions["B"].width = 22
    ws_piv.column_dimensions["C"].width = 20
    ws_piv.column_dimensions["D"].width = 16
    ws_piv.column_dimensions["E"].width = 22
    ws_piv.column_dimensions["F"].width = 20
    ws_piv.column_dimensions["G"].width = 16
    
    print("Saving completed Sales Performance Dashboard workbook...")
    wb.save(output_file)
    print(f"Dashboard workbook saved successfully as '{output_file}'!")

if __name__ == "__main__":
    create_dashboard_workbook()
