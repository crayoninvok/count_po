import io
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import plotly.express as px
import plotly.graph_objects as go

def create_excel_simple(stats_df, counts, vendor_name, df_raw=None, po_stats_df=None, po_counts=None, po_breakdown_df=None):
    """
    Membuat file Excel dengan data tabel, grafik, dan vendor report.
    """
    output = io.BytesIO()
    
    # Fix data types untuk menghindari Arrow serialization error
    if df_raw is not None:
        df_raw = df_raw.copy()
        for col in df_raw.select_dtypes(include=['object']).columns:
            df_raw[col] = df_raw[col].astype(str)
    
    if po_breakdown_df is not None:
        po_breakdown_df = po_breakdown_df.copy()
        for col in po_breakdown_df.select_dtypes(include=['object']).columns:
            po_breakdown_df[col] = po_breakdown_df[col].astype(str)
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Format definitions
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        
        title_format = workbook.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#2E5090', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        
        subtitle_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'bg_color': '#5B9BD5', 'font_color': 'white',
            'align': 'left', 'valign': 'vcenter', 'border': 1
        })
        
        data_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        number_format = workbook.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0'})
        currency_format = workbook.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter', 'num_format': 'Rp #,##0'})
        percentage_format = workbook.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter', 'num_format': '0.00%'})
        date_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'num_format': 'yyyy-mm-dd'})
        
        # ===== Sheet 1: Dashboard =====
        worksheet_dashboard = workbook.add_worksheet('Dashboard')
        worksheet_dashboard.set_column('A:A', 30)
        worksheet_dashboard.set_column('B:D', 20)
        
        worksheet_dashboard.merge_range('A1:D1', 'DASHBOARD ANALISIS TRANSAKSI PO', title_format)
        worksheet_dashboard.write('A2', 'Vendor', header_format)
        worksheet_dashboard.merge_range('B2:D2', vendor_name, data_format)
        
        worksheet_dashboard.write('A4', 'KEY METRICS', subtitle_format)
        worksheet_dashboard.write('A5', 'Metrik', header_format)
        worksheet_dashboard.write('B5', 'Nilai', header_format)
        
        metrics = [
            ('Total Transaksi', counts[0], number_format),
            ('Total Nilai', counts[1], currency_format),
            ('Rata-rata Transaksi', counts[1]/counts[0] if counts[0] > 0 else 0, currency_format),
        ]
        
        if po_counts is not None:
            metrics.append(('Total PO Unik', po_counts[0], number_format))
        
        if df_raw is not None and 'jumlah' in df_raw.columns:
            metrics.append(('Transaksi Terbesar', df_raw['jumlah'].max(), currency_format))
            metrics.append(('Transaksi Terkecil', df_raw['jumlah'].min(), currency_format))
        
        row = 6
        for label, value, fmt in metrics:
            worksheet_dashboard.write(row, 0, label, data_format)
            worksheet_dashboard.write(row, 1, value, fmt)
            row += 1
        
        # Charts
        chart_column = workbook.add_chart({'type': 'column'})
        chart_column.add_series({
            'name': 'Jumlah Transaksi',
            'categories': f'=\'Statistik Detail\'!$A$4:$A${3+len(stats_df)}',
            'values': f'=\'Statistik Detail\'!$B$4:$B${3+len(stats_df)}',
            'fill': {'color': '#4472C4'},
            'data_labels': {'value': True, 'position': 'outside_end'}
        })
        chart_column.set_title({'name': 'Distribusi Transaksi Berdasarkan Rentang'})
        chart_column.set_style(11)
        chart_column.set_size({'width': 720, 'height': 400})
        worksheet_dashboard.insert_chart('A13', chart_column)
        
        chart_pie = workbook.add_chart({'type': 'pie'})
        chart_pie.add_series({
            'categories': f'=\'Statistik Detail\'!$A$4:$A${3+len(stats_df)}',
            'values': f'=\'Statistik Detail\'!$B$4:$B${3+len(stats_df)}',
            'data_labels': {'percentage': True, 'category': True},
        })
        chart_pie.set_title({'name': 'Proporsi Transaksi (%)'})
        chart_pie.set_style(10)
        chart_pie.set_size({'width': 720, 'height': 400})
        worksheet_dashboard.insert_chart('A35', chart_pie)
        
        # ===== Sheet 2: Statistik Detail =====
        worksheet2 = workbook.add_worksheet('Statistik Detail')
        worksheet2.set_column('A:A', 30)
        worksheet2.set_column('B:C', 20)
        
        worksheet2.merge_range('A1:C1', 'STATISTIK DETAIL BERDASARKAN RENTANG', title_format)
        
        for col, header in enumerate(['Rentang Transaksi', 'Jumlah', 'Persentase']):
            worksheet2.write(2, col, header, header_format)
        
        row = 3
        for idx in range(len(stats_df)):
            worksheet2.write(row, 0, f"Rp {stats_df.iloc[idx, 0]}", data_format)
            worksheet2.write(row, 1, stats_df.iloc[idx, 1], number_format)
            worksheet2.write(row, 2, stats_df.iloc[idx, 2]/100, percentage_format)
            row += 1
        
        worksheet2.write(row, 0, 'TOTAL', header_format)
        worksheet2.write(row, 1, counts[0], number_format)
        worksheet2.write(row, 2, 1.0, percentage_format)
        
        # ===== Sheet 3: PO Unik =====
        if po_stats_df is not None and po_counts is not None:
            worksheet_po = workbook.add_worksheet('PO Unik')
            worksheet_po.set_column('A:A', 30)
            worksheet_po.set_column('B:C', 20)
            
            worksheet_po.merge_range('A1:C1', 'STATISTIK PO UNIK BERDASARKAN RENTANG', title_format)
            worksheet_po.write('A2', 'Total PO Unik', header_format)
            worksheet_po.write('B2', po_counts[0], number_format)
            
            for col, header in enumerate(['Rentang Transaksi', 'Jumlah PO', 'Persentase']):
                worksheet_po.write(3, col, header, header_format)
            
            row = 4
            for idx in range(len(po_stats_df)):
                worksheet_po.write(row, 0, f"Rp {po_stats_df.iloc[idx, 0]}", data_format)
                worksheet_po.write(row, 1, po_stats_df.iloc[idx, 1], number_format)
                worksheet_po.write(row, 2, po_stats_df.iloc[idx, 2]/100, percentage_format)
                row += 1
            
            worksheet_po.write(row, 0, 'TOTAL', header_format)
            worksheet_po.write(row, 1, po_counts[0], number_format)
            worksheet_po.write(row, 2, 1.0, percentage_format)
            
            chart_po = workbook.add_chart({'type': 'column'})
            chart_po.add_series({
                'name': 'Jumlah PO Unik',
                'categories': f'=\'PO Unik\'!$A$5:$A${4+len(po_stats_df)}',
                'values': f'=\'PO Unik\'!$B$5:$B${4+len(po_stats_df)}',
                'fill': {'color': '#70AD47'},
                'data_labels': {'value': True, 'position': 'outside_end'}
            })
            chart_po.set_title({'name': 'Distribusi PO Unik Berdasarkan Rentang'})
            chart_po.set_style(11)
            chart_po.set_size({'width': 720, 'height': 400})
            worksheet_po.insert_chart(f'A{row+3}', chart_po)
        
        # ===== Sheet 4: Breakdown PO Detail =====
        if po_breakdown_df is not None and not po_breakdown_df.empty:
            worksheet_po_detail = workbook.add_worksheet('Breakdown PO Detail')
            worksheet_po_detail.set_column('A:A', 25)
            worksheet_po_detail.set_column('B:B', 20)
            worksheet_po_detail.set_column('C:C', 25)
            worksheet_po_detail.set_column('D:D', 25)
            
            worksheet_po_detail.merge_range('A1:D1', 'BREAKDOWN DETAIL SETIAP PO', title_format)
            worksheet_po_detail.write('A2', f'Total PO: {len(po_breakdown_df)}', data_format)
            
            for col, header in enumerate(['PO Code', 'Jumlah Transaksi', 'Total Nilai', 'Rentang']):
                worksheet_po_detail.write(3, col, header, header_format)
            
            row = 4
            for idx in range(len(po_breakdown_df)):
                worksheet_po_detail.write(row, 0, str(po_breakdown_df.iloc[idx, 0]), data_format)
                worksheet_po_detail.write(row, 1, po_breakdown_df.iloc[idx, 1], number_format)
                worksheet_po_detail.write(row, 2, po_breakdown_df.iloc[idx, 2], currency_format)
                worksheet_po_detail.write(row, 3, str(po_breakdown_df.iloc[idx, 3]), data_format)
                row += 1
            
            worksheet_po_detail.write(row, 0, 'TOTAL', header_format)
            worksheet_po_detail.write(row, 1, po_breakdown_df['Jumlah Transaksi'].sum(), number_format)
            worksheet_po_detail.write(row, 2, po_breakdown_df['Total Nilai'].sum(), currency_format)
            worksheet_po_detail.write(row, 3, '', data_format)
        
        # ===== Sheet 5: Data Mentah =====
        if df_raw is not None and not df_raw.empty:
            df_clean = df_raw.copy()
            
            for col in df_clean.columns:
                if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
                    df_clean[col] = df_clean[col].apply(lambda x: None if pd.isna(x) else x)
            
            worksheet4 = workbook.add_worksheet('Data Mentah')
            worksheet4.merge_range(0, 0, 0, len(df_clean.columns)-1, 'DATA TRANSAKSI LENGKAP', title_format)
            
            for col_num, col_name in enumerate(df_clean.columns):
                worksheet4.write(1, col_num, col_name, header_format)
            
            for idx, col in enumerate(df_clean.columns):
                max_len = max(df_clean[col].astype(str).apply(len).max(), len(str(col))) + 2
                worksheet4.set_column(idx, idx, min(max_len, 50))
            
            for row_num, row_data in enumerate(df_clean.values):
                for col_num, cell_value in enumerate(row_data):
                    col_name = df_clean.columns[col_num]
                    
                    if cell_value is None or pd.isna(cell_value):
                        worksheet4.write(row_num + 2, col_num, '', data_format)
                    elif col_name == 'jumlah':
                        worksheet4.write(row_num + 2, col_num, float(cell_value), currency_format)
                    elif pd.api.types.is_datetime64_any_dtype(df_clean[col_name]):
                        worksheet4.write(row_num + 2, col_num, cell_value, date_format)
                    else:
                        worksheet4.write(row_num + 2, col_num, str(cell_value), data_format)
            
            if 'jumlah' in df_clean.columns:
                jumlah_col_idx = df_clean.columns.get_loc('jumlah')
                total_row = len(df_clean) + 2
                
                if jumlah_col_idx > 0:
                    worksheet4.write(total_row, jumlah_col_idx - 1, 'TOTAL:', header_format)
                
                worksheet4.write(total_row, jumlah_col_idx, df_clean['jumlah'].sum(), currency_format)
    
    output.seek(0)
    return output


def create_pdf_report(stats_df, counts, vendor_name, df_raw=None):
    """Membuat file PDF untuk presentasi."""
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), 
                           leftMargin=0.5*inch, rightMargin=0.5*inch,
                           topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'],
        fontSize=18, textColor=colors.HexColor('#2E5090'),
        spaceAfter=12, alignment=TA_CENTER, fontName='Helvetica-Bold')
    
    heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'],
        fontSize=14, textColor=colors.HexColor('#4472C4'),
        spaceAfter=10, spaceBefore=12, fontName='Helvetica-Bold')
    
    story.append(Paragraph('LAPORAN ANALISIS TRANSAKSI PO', title_style))
    story.append(Paragraph(f'Vendor: {vendor_name}', styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph('RINGKASAN TRANSAKSI', heading_style))
    
    ringkasan_data = [
        ['Metrik', 'Nilai'],
        ['Total Transaksi', f"{counts[0]:,}".replace(",", ".")],
        ['Total Nilai', f"Rp {counts[1]:,.0f}".replace(",", ".")],
        ['Rata-rata Transaksi', f"Rp {counts[1]/counts[0]:,.0f}".replace(",", ".") if counts[0] > 0 else "Rp 0"],
    ]
    
    if df_raw is not None and 'jumlah' in df_raw.columns:
        ringkasan_data.append(['Transaksi Terbesar', f"Rp {df_raw['jumlah'].max():,.0f}".replace(",", ".")])
        ringkasan_data.append(['Transaksi Terkecil', f"Rp {df_raw['jumlah'].min():,.0f}".replace(",", ".")])
    
    ringkasan_table = Table(ringkasan_data, colWidths=[3*inch, 3*inch])
    ringkasan_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
    ]))
    
    story.append(ringkasan_table)
    story.append(Spacer(1, 0.4*inch))
    story.append(Paragraph('STATISTIK DETAIL BERDASARKAN RENTANG', heading_style))
    
    stats_data = [['Rentang Transaksi', 'Jumlah', 'Persentase']]
    for idx in range(len(stats_df)):
        stats_data.append([
            f"Rp {stats_df.iloc[idx, 0]}",
            f"{stats_df.iloc[idx, 1]:,}".replace(",", "."),
            f"{stats_df.iloc[idx, 2]:.2f}%"
        ])
    stats_data.append(['TOTAL', f"{counts[0]:,}".replace(",", "."), '100.00%'])
    
    stats_table = Table(stats_data, colWidths=[3.5*inch, 2*inch, 2*inch])
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E7E6E6')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
    ]))
    
    story.append(stats_table)
    story.append(PageBreak())
    story.append(Paragraph('VISUALISASI DATA', heading_style))
    story.append(Spacer(1, 0.2*inch))
    
    try:
        fig_bar = px.bar(stats_df, x='Rentang', y='Jumlah',
            title='Jumlah Transaksi Berdasarkan Rentang',
            color='Jumlah', color_continuous_scale='Blues', text='Jumlah')
        fig_bar.update_traces(textposition='outside')
        fig_bar.update_layout(showlegend=False, height=450, width=1000,
            xaxis_tickangle=-45, title=dict(font=dict(size=16, color='#2E5090')))
        
        img_bytes_bar = fig_bar.to_image(format="png", engine="kaleido", scale=2)
        story.append(Image(io.BytesIO(img_bytes_bar), width=9*inch, height=4*inch))
        story.append(Spacer(1, 0.3*inch))
    except:
        pass
    
    try:
        fig_pie = go.Figure(data=[go.Pie(
            labels=stats_df['Rentang'], values=stats_df['Jumlah'], hole=0.4,
            textposition='inside', textinfo='percent+label')])
        fig_pie.update_layout(title='Distribusi Transaksi (%)', height=450, width=1000)
        
        img_bytes_pie = fig_pie.to_image(format="png", engine="kaleido", scale=2)
        story.append(Image(io.BytesIO(img_bytes_pie), width=9*inch, height=4*inch))
    except:
        pass
    
    doc.build(story)
    output.seek(0)
    return output