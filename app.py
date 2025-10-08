import streamlit as st
import pandas as pd
import locale
from transaction_utils import (count_transactions, get_transaction_dataframe, 
                                get_range_statistics, count_unique_po_by_range, 
                                get_unique_po_statistics, get_po_breakdown)
from excel_exporter import create_excel_simple, create_pdf_report
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Set locale ke Indonesia untuk format mata uang
try:
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'id_ID')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, '')

# Konfigurasi halaman
st.set_page_config(
    page_title="Analisis Transaksi PO",
    page_icon="📊",
    layout="wide"
)

def format_currency(amount):
    """Format jumlah sebagai Rupiah Indonesia"""
    try:
        formatted = locale.currency(amount, grouping=True, symbol=False)
        return f"Rp {formatted}"
    except:
        return f"Rp {amount:,.0f}".replace(",", ".")

def main():
    st.title("📊 Analisis Monitoring Transaksi PO")
    st.markdown("---")
    
    # Sidebar untuk upload file dan filter
    with st.sidebar:
        st.header("Upload & Filter")
        uploaded_file = st.file_uploader(
            "Upload File Excel Purchase Order", 
            type=["xls", "xlsx"],
            help="Upload file Excel yang berisi data transaksi PO"
        )
        
        if uploaded_file is not None:
            st.success("✅ File berhasil diupload!")
    
    if uploaded_file is not None:
        try:
            # Muat dataset untuk mendapatkan vendor unik
            df = pd.read_excel(uploaded_file, sheet_name=0)
            
            vendor_column = 'vendor_name'
            if vendor_column not in df.columns:
                st.error(f"❌ Kolom '{vendor_column}' tidak ditemukan dalam file.")
                st.info(f"Kolom yang tersedia: {', '.join(df.columns)}")
                return
            
            # Dapatkan vendor unik dan urutkan
            vendors = df[vendor_column].dropna().unique()
            vendors = sorted(vendors)
            
            # Filter vendor di sidebar
            with st.sidebar:
                st.markdown("---")
                selected_vendor = st.selectbox(
                    "🏢 Pilih Vendor", 
                    options=["Semua Vendor"] + list(vendors),
                    help="Filter transaksi berdasarkan vendor"
                )
                
                st.info(f"📋 Total vendor: {len(vendors)}")
            
            # Dapatkan data
            counts = count_transactions(
                uploaded_file, 
                vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
            )
            stats_df = get_range_statistics(
                uploaded_file,
                vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
            )
            
            # Dapatkan data PO unik
            try:
                po_counts = count_unique_po_by_range(
                    uploaded_file,
                    vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
                )
                po_stats_df = get_unique_po_statistics(
                    uploaded_file,
                    vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
                )
                po_breakdown_df = get_po_breakdown(
                    uploaded_file,
                    vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
                )
            except Exception as e:
                po_counts = None
                po_stats_df = None
                po_breakdown_df = None
                st.warning(f"⚠️ Data PO unik tidak tersedia: {e}")
            
            # Header
            if selected_vendor != "Semua Vendor":
                st.subheader(f"📈 Analisis untuk: {selected_vendor}")
            else:
                st.subheader("📈 Analisis untuk: Semua Vendor")
            
            # Metrik Utama
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(
                    label="Total Transaksi",
                    value=f"{counts[0]:,}".replace(",", "."),
                    help="Jumlah total purchase order"
                )
            
            with col2:
                st.metric(
                    label="Total Nilai",
                    value=format_currency(counts[1]),
                    help="Jumlah total nilai transaksi"
                )
            
            with col3:
                avg_transaction = counts[1] / counts[0] if counts[0] > 0 else 0
                st.metric(
                    label="Rata-rata Transaksi",
                    value=format_currency(avg_transaction),
                    help="Nilai rata-rata transaksi"
                )
            
            with col4:
                df_filtered = get_transaction_dataframe(
                    uploaded_file,
                    vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
                )
                max_transaction = df_filtered['jumlah'].max()
                st.metric(
                    label="Transaksi Terbesar",
                    value=format_currency(max_transaction),
                    help="Nilai transaksi tertinggi"
                )
            
            st.markdown("---")
            
            # Tabs
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 Distribusi", 
                "📈 Statistik", 
                "📋 Detail", 
                "🔢 PO Unik",
                "📁 Data Mentah",
                "ℹ️ Info"
            ])
            
            with tab1:
                col1, col2 = st.columns(2)
                
                with col1:
                    fig_bar = px.bar(
                        stats_df,
                        x='Rentang',
                        y='Jumlah',
                        title='Jumlah Transaksi Berdasarkan Rentang',
                        labels={'Jumlah': 'Jumlah Transaksi', 'Rentang': 'Rentang Transaksi'},
                        color='Jumlah',
                        color_continuous_scale='Blues',
                        text='Jumlah'
                    )
                    fig_bar.update_traces(textposition='outside', textfont_size=11)
                    fig_bar.update_layout(
                        showlegend=False, 
                        height=500,
                        xaxis_tickangle=-45,
                        xaxis=dict(tickfont=dict(size=10), automargin=True),
                        yaxis=dict(tickfont=dict(size=11)),
                        margin=dict(l=50, r=50, t=80, b=120),
                        font=dict(size=11),
                        title=dict(font=dict(size=14))
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with col2:
                    fig_pie = px.pie(
                        stats_df,
                        values='Jumlah',
                        names='Rentang',
                        title='Distribusi Transaksi (%)',
                        hole=0.4
                    )
                    fig_pie.update_traces(
                        textposition='inside', 
                        textinfo='percent+label',
                        textfont_size=10
                    )
                    fig_pie.update_layout(
                        height=500,
                        margin=dict(l=20, r=20, t=80, b=20),
                        font=dict(size=10),
                        title=dict(font=dict(size=14))
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                st.markdown("---")
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    try:
                        excel_buffer = create_excel_simple(
                            stats_df, counts, selected_vendor, 
                            df_raw=df_filtered,
                            po_stats_df=po_stats_df,
                            po_counts=po_counts,
                            po_breakdown_df=po_breakdown_df
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Distribusi_{selected_vendor.replace(' ', '_')}_{timestamp}.xlsx"
                        st.download_button(
                            label="📥 Download Excel",
                            data=excel_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download laporan dalam format Excel",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file Excel: {e}")
                
                with col_dl2:
                    try:
                        pdf_buffer = create_pdf_report(
                            stats_df, counts, selected_vendor, df_raw=df_filtered
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Distribusi_{selected_vendor.replace(' ', '_')}_{timestamp}.pdf"
                        st.download_button(
                            label="📄 Download PDF (Presentasi)",
                            data=pdf_buffer,
                            file_name=filename,
                            mime="application/pdf",
                            help="Download laporan untuk presentasi dalam format PDF",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file PDF: {e}")
            
            with tab2:
                st.subheader("📊 Statistik Detail Berdasarkan Rentang")
                
                detailed_stats = stats_df.copy()
                detailed_stats['Jumlah'] = detailed_stats['Jumlah'].apply(lambda x: f"{x:,}".replace(",", "."))
                detailed_stats['Persentase'] = detailed_stats['Persentase'].apply(lambda x: f"{x:.2f}%")
                detailed_stats = detailed_stats.reset_index(drop=True)
                
                st.dataframe(detailed_stats, use_container_width=True)
                
                st.markdown("### 💡 Insight")
                col1, col2 = st.columns(2)
                
                with col1:
                    small_transactions = sum(counts[2:4])
                    small_pct = (small_transactions / counts[0] * 100) if counts[0] > 0 else 0
                    st.info(f"**Transaksi Kecil (< 500 Ribu):** {small_transactions:,} ({small_pct:.1f}%)".replace(",", "."))
                    
                    medium_transactions = sum(counts[4:6])
                    medium_pct = (medium_transactions / counts[0] * 100) if counts[0] > 0 else 0
                    st.info(f"**Transaksi Menengah (500 Ribu-5 Juta):** {medium_transactions:,} ({medium_pct:.1f}%)".replace(",", "."))
                
                with col2:
                    large_transactions = sum(counts[6:8])
                    large_pct = (large_transactions / counts[0] * 100) if counts[0] > 0 else 0
                    st.info(f"**Transaksi Besar (5 Juta-50 Juta):** {large_transactions:,} ({large_pct:.1f}%)".replace(",", "."))
                    
                    xlarge_transactions = counts[8]
                    xlarge_pct = (xlarge_transactions / counts[0] * 100) if counts[0] > 0 else 0
                    st.info(f"**Transaksi Sangat Besar (> 50 Juta):** {xlarge_transactions:,} ({xlarge_pct:.1f}%)".replace(",", "."))
                
                st.markdown("---")
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    try:
                        excel_buffer = create_excel_simple(
                            stats_df, counts, selected_vendor, df_raw=df_filtered
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Statistik_{selected_vendor.replace(' ', '_')}_{timestamp}.xlsx"
                        st.download_button(
                            label="📥 Download Excel",
                            data=excel_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download statistik dalam format Excel",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file Excel: {e}")
                
                with col_dl2:
                    try:
                        pdf_buffer = create_pdf_report(
                            stats_df, counts, selected_vendor, df_raw=df_filtered
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Statistik_{selected_vendor.replace(' ', '_')}_{timestamp}.pdf"
                        st.download_button(
                            label="📄 Download PDF (Presentasi)",
                            data=pdf_buffer,
                            file_name=filename,
                            mime="application/pdf",
                            help="Download statistik untuk presentasi dalam format PDF",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file PDF: {e}")
            
            with tab3:
                st.subheader("📋 Detail Rentang Transaksi")
                
                ranges_info = [
                    ("0 - 100.000", counts[2], "🔵"),
                    ("100.001 - 500.000", counts[3], "🟢"),
                    ("500.001 - 1.000.000", counts[4], "🟡"),
                    ("1.000.001 - 5.000.000", counts[5], "🟠"),
                    ("5.000.001 - 10.000.000", counts[6], "🔴"),
                    ("10.000.001 - 50.000.000", counts[7], "🟣"),
                    ("Di atas 50.000.000", counts[8], "⭐")
                ]
                
                for range_label, count, icon in ranges_info:
                    percentage = (count / counts[0] * 100) if counts[0] > 0 else 0
                    formatted_count = f"{count:,}".replace(",", ".")
                    st.markdown(f"{icon} **Rp {range_label}**: {formatted_count} transaksi ({percentage:.2f}%)")
                
                st.markdown("---")
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    try:
                        excel_buffer = create_excel_simple(
                            stats_df, counts, selected_vendor, df_raw=df_filtered
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Detail_{selected_vendor.replace(' ', '_')}_{timestamp}.xlsx"
                        st.download_button(
                            label="📥 Download Excel",
                            data=excel_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download detail lengkap dalam format Excel",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file Excel: {e}")
                
                with col_dl2:
                    try:
                        pdf_buffer = create_pdf_report(
                            stats_df, counts, selected_vendor, df_raw=df_filtered
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Detail_{selected_vendor.replace(' ', '_')}_{timestamp}.pdf"
                        st.download_button(
                            label="📄 Download PDF (Presentasi)",
                            data=pdf_buffer,
                            file_name=filename,
                            mime="application/pdf",
                            help="Download detail untuk presentasi dalam format PDF",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file PDF: {e}")
            
            with tab4:
                st.subheader("🔢 Analisis PO Unik (Berdasarkan po_code)")
                
                if po_counts is None or po_stats_df is None:
                    st.error("❌ Kolom 'po_code' tidak ditemukan dalam file. Tab ini memerlukan kolom po_code untuk analisis.")
                else:
                    st.info(f"📌 Total PO Unik: **{po_counts[0]:,}** PO".replace(",", "."))
                    st.markdown("*PO dengan kode yang sama dihitung sebagai 1 PO*")
                    st.markdown("---")
                    
                    # Grafik PO Unik
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig_bar_po = px.bar(
                            po_stats_df,
                            x='Rentang',
                            y='Jumlah PO',
                            title='Jumlah PO Unik Berdasarkan Rentang',
                            labels={'Jumlah PO': 'Jumlah PO', 'Rentang': 'Rentang Nilai PO'},
                            color='Jumlah PO',
                            color_continuous_scale='Greens',
                            text='Jumlah PO'
                        )
                        fig_bar_po.update_traces(textposition='outside', textfont_size=11)
                        fig_bar_po.update_layout(
                            showlegend=False,
                            height=500,
                            xaxis_tickangle=-45,
                            xaxis=dict(tickfont=dict(size=10), automargin=True),
                            yaxis=dict(tickfont=dict(size=11)),
                            margin=dict(l=50, r=50, t=80, b=120),
                            font=dict(size=11),
                            title=dict(font=dict(size=14))
                        )
                        st.plotly_chart(fig_bar_po, use_container_width=True)
                    
                    with col2:
                        fig_pie_po = px.pie(
                            po_stats_df,
                            values='Jumlah PO',
                            names='Rentang',
                            title='Distribusi PO Unik (%)',
                            hole=0.4
                        )
                        fig_pie_po.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            textfont_size=10
                        )
                        fig_pie_po.update_layout(
                            height=500,
                            margin=dict(l=20, r=20, t=80, b=20),
                            font=dict(size=10),
                            title=dict(font=dict(size=14))
                        )
                        st.plotly_chart(fig_pie_po, use_container_width=True)
                    
                    st.markdown("---")
                    st.subheader("📊 Detail PO Unik per Rentang")
                    
                    # Tabel detail
                    detailed_po_stats = po_stats_df.copy()
                    detailed_po_stats['Jumlah PO'] = detailed_po_stats['Jumlah PO'].apply(lambda x: f"{x:,}".replace(",", "."))
                    detailed_po_stats['Persentase'] = detailed_po_stats['Persentase'].apply(lambda x: f"{x:.2f}%")
                    detailed_po_stats = detailed_po_stats.reset_index(drop=True)
                    
                    st.dataframe(detailed_po_stats, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("### 📋 Breakdown PO Unik")
                    
                    ranges_info_po = [
                        ("0 - 100.000", po_counts[1], "🔵"),
                        ("100.001 - 500.000", po_counts[2], "🟢"),
                        ("500.001 - 1.000.000", po_counts[3], "🟡"),
                        ("1.000.001 - 5.000.000", po_counts[4], "🟠"),
                        ("5.000.001 - 10.000.000", po_counts[5], "🔴"),
                        ("10.000.001 - 50.000.000", po_counts[6], "🟣"),
                        ("Di atas 50.000.000", po_counts[7], "⭐")
                    ]
                    
                    for range_label, count, icon in ranges_info_po:
                        percentage = (count / po_counts[0] * 100) if po_counts[0] > 0 else 0
                        formatted_count = f"{count:,}".replace(",", ".")
                        st.markdown(f"{icon} **Rp {range_label}**: {formatted_count} PO ({percentage:.2f}%)")
                    
                    # Breakdown detail tiap PO
                    if po_breakdown_df is not None and not po_breakdown_df.empty:
                        st.markdown("---")
                        st.subheader("📦 Detail Breakdown Setiap PO")
                        st.markdown("*Menampilkan jumlah transaksi dan total nilai per PO*")
                        
                        # Format breakdown dataframe
                        po_breakdown_display = po_breakdown_df.copy()
                        po_breakdown_display['Total Nilai'] = po_breakdown_display['Total Nilai'].apply(
                            lambda x: format_currency(x)
                        )
                        po_breakdown_display['Jumlah Transaksi'] = po_breakdown_display['Jumlah Transaksi'].apply(
                            lambda x: f"{x:,}".replace(",", ".")
                        )
                        
                        # Tambahkan search box
                        search_po = st.text_input("🔍 Cari PO Code", placeholder="Ketik kode PO...")
                        
                        if search_po:
                            filtered_breakdown = po_breakdown_display[
                                po_breakdown_display['PO Code'].str.contains(search_po, case=False, na=False)
                            ]
                            st.info(f"Ditemukan {len(filtered_breakdown)} PO")
                            st.dataframe(filtered_breakdown, use_container_width=True, height=400)
                        else:
                            st.dataframe(po_breakdown_display, use_container_width=True, height=400)
                        
                        # Download breakdown
                        col1, col2 = st.columns(2)
                        with col1:
                            csv = po_breakdown_df.to_csv(index=False).encode('utf-8')
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            st.download_button(
                                label="📥 Download Breakdown CSV",
                                data=csv,
                                file_name=f"PO_Breakdown_{selected_vendor.replace(' ', '_')}_{timestamp}.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
            
            with tab5:
                st.subheader("📁 Data Transaksi Mentah")
                
                df_display = get_transaction_dataframe(
                    uploaded_file,
                    vendor=None if selected_vendor == "Semua Vendor" else selected_vendor
                )
                
                formatted_count = f"{len(df_display):,}".replace(",", ".")
                st.info(f"Menampilkan {formatted_count} transaksi")
                st.dataframe(df_display, use_container_width=True, height=400)
                
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    csv = df_display.to_csv(index=False).encode('utf-8')
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📥 Download CSV",
                        data=csv,
                        file_name=f"transaksi_{selected_vendor.replace(' ', '_')}_{timestamp}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    try:
                        excel_buffer = create_excel_simple(
                            stats_df, counts, selected_vendor, df_raw=df_display
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Lengkap_{selected_vendor.replace(' ', '_')}_{timestamp}.xlsx"
                        st.download_button(
                            label="📥 Download Excel",
                            data=excel_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download data lengkap dalam format Excel",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file Excel: {e}")
                
                with col3:
                    try:
                        pdf_buffer = create_pdf_report(
                            stats_df, counts, selected_vendor, df_raw=df_display
                        )
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Laporan_Lengkap_{selected_vendor.replace(' ', '_')}_{timestamp}.pdf"
                        st.download_button(
                            label="📄 Download PDF",
                            data=pdf_buffer,
                            file_name=filename,
                            mime="application/pdf",
                            help="Download laporan untuk presentasi dalam format PDF",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Error membuat file PDF: {e}")
            
            with tab6:
                st.subheader("ℹ️ Informasi Aplikasi")
                
                st.markdown("""
                ### 📊 Tentang Aplikasi
                
                Aplikasi ini dirancang untuk menganalisis data Purchase Order (PO) dengan berbagai perspektif:
                
                **Tab yang Tersedia:**
                
                1. **📊 Distribusi**: Visualisasi grafik distribusi transaksi
                2. **📈 Statistik**: Breakdown statistik detail dan insight
                3. **📋 Detail**: Detail rentang transaksi per kategori
                4. **🔢 PO Unik**: Analisis berdasarkan PO code dengan breakdown detail per PO
                5. **📁 Data Mentah**: Dataset lengkap yang dapat diunduh
                6. **ℹ️ Info**: Halaman ini
                
                ---
                
                ### 🎯 Perbedaan Analisis
                
                **Analisis Transaksi (Tab 1-3):**
                - Menghitung setiap baris data sebagai 1 transaksi
                - Cocok untuk melihat volume transaksi detail
                
                **Analisis PO Unik (Tab 4):**
                - Menggabungkan transaksi dengan `po_code` yang sama
                - PO yang sama dihitung sebagai 1 PO
                - Menampilkan breakdown detail setiap PO (jumlah transaksi & total nilai)
                - Cocok untuk melihat jumlah PO aktual
                
                ---
                
                ### 📋 Format Excel yang Diperlukan
                
                File Excel harus memiliki kolom:
                - `vendor_name`: Nama vendor
                - `jumlah`: Nilai transaksi (numerik dalam Rupiah)
                - `po_status_approval`: Status approval
                - `po_code`: Kode PO (opsional, diperlukan untuk tab PO Unik)
                
                **Catatan:** Transaksi dengan status "Not Yet Approved" otomatis dikecualikan dari analisis.
                
                ---
                
                ### 📈 Rentang Transaksi
                
                - 🔵 Rp 0 - 100.000
                - 🟢 Rp 100.001 - 500.000
                - 🟡 Rp 500.001 - 1.000.000
                - 🟠 Rp 1.000.001 - 5.000.000
                - 🔴 Rp 5.000.001 - 10.000.000
                - 🟣 Rp 10.000.001 - 50.000.000
                - ⭐ Di atas Rp 50.000.000
                """)
            
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}")
            st.exception(e)
    
    else:
        st.info("👈 Silakan upload file Excel dari sidebar untuk memulai analisis")
        
        st.markdown("""
        ### 📖 Cara menggunakan aplikasi ini:
        
        1. **Upload** file Excel Purchase Order Anda menggunakan sidebar
        2. **Pilih** vendor untuk filter (atau lihat semua vendor)
        3. **Jelajahi** berbagai tab:
           - 📊 **Distribusi**: Grafik visual yang menunjukkan pola transaksi
           - 📈 **Statistik**: Breakdown detail dan insight
           - 📋 **Detail**: Jumlah transaksi berdasarkan rentang
           - 🔢 **PO Unik**: Analisis berdasarkan PO code dengan breakdown detail per PO
           - 📁 **Data Mentah**: Lihat dan unduh dataset lengkap
           - ℹ️ **Info**: Informasi lengkap tentang aplikasi
        4. **Download** laporan dalam format Excel atau PDF untuk presentasi!
        
        ### 📋 Format Excel yang Diperlukan:
        
        File Excel Anda harus berisi kolom-kolom ini:
        - `vendor_name`: Nama vendor
        - `jumlah`: Jumlah transaksi (numerik dalam Rupiah)
        - `po_status_approval`: Status approval (transaksi "Not Yet Approved" akan dikecualikan)
        - `po_code`: Kode PO (opsional, diperlukan untuk analisis PO Unik)
        """)

if __name__ == "__main__":
    main()