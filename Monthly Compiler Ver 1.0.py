# -*- coding: utf-8 -*-
"""
================================================================================
Monthly Media Monitoring Report Compiler
================================================================================
* Author      : Hizkia Gerald Garibaldi
* Email       : hgeraldgaribaldi@gmail.com
* Version     : 1.0.0
* Last Update : 8 Juli 2025]
* Description : Script to compile media monitoring daily reports from 
* Word file into one monthly report Excel File.
*
* Developed with guidance and assistance from Google's AI, Gemini.
*
* This software is licensed under the MIT License.
* Copyright (c) [2025] Hizkia Gerald Garibaldi
================================================================================
"""


import os
import pandas as pd
from docx import Document
import re
from docx.oxml.ns import qn
from datetime import datetime

# ==============================================================================
# "PABRIK" EKSTRAKSI DATA 
# ==============================================================================
def extract_data_from_docx(docx_path):
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"  - Gagal membuka file: {os.path.basename(docx_path)}. Error: {e}")
        return None

    data = {
        'Category': [], 'Date': [], 'Title': [], 'Media': [],
        'Journalist': [], 'Page Number': [], 'Link': []
    }
    
    category_map = {
        "Client News": "Client News", "Corporate News": "Corporate News",
        "Industry & Regulatory News": "Industry & Regulatory News",
        "Rental & Autopool Industry": "Industry & Regulatory News",
        "Logistic & Express Courier Industry": "Industry & Regulatory News",
        "Car Auction & Selling Industry": "Industry & Regulatory News",
        "Ammonia News": "Industry & Regulatory News",
        "LPG News": "Industry & Regulatory News"
    }
    valid_categories = list(category_map.keys())

    def get_text_from_element(element):
        text = ''
        for t in element.iter():
            if t.tag == qn('w:t') and t.text:
                text += t.text
        return text.strip()

    for table in doc.tables:
        if not table.rows:
            continue
        first_cell_text = get_text_from_element(table.rows[0].cells[0]._element)
        if first_cell_text in valid_categories:
            current_category = category_map[first_cell_text]
        else:
            continue

        for row in table.rows[1:]:
            cells_text = [get_text_from_element(cell._element) for cell in row.cells]
            if len(cells_text) < 2 or cells_text[0] == "No." or cells_text[0] == "":
                continue
            
            # --- LOGIKA PARSING "DETEKTIF 2.0" ---
            title, media, journalist, date, link, page_number = "", "-", "-", None, "-", "-"
            combined_text = cells_text[1].strip()
            
            link_start_index = combined_text.find('http')
            
            if link_start_index != -1:
                link = combined_text[link_start_index:]
                text_before_link = combined_text[:link_start_index]
                if len(text_before_link) >= 8 and text_before_link[-8:].isdigit():
                    extracted_date = text_before_link[-8:]
                    content_string = text_before_link[:-8]
                else:
                    content_string = text_before_link
            else:
                if len(combined_text) >= 8 and combined_text[-8:].isdigit():
                    extracted_date = combined_text[-8:]
                    content_string = combined_text[:-8]
                else:
                    content_string = combined_text

            if extracted_date:
                try:
                    date = datetime.strptime(extracted_date, '%Y%m%d').date()
                except ValueError:
                    date = None
            
            if content_string.endswith('_'):
                content_string = content_string[:-1]

            if content_string:
                content_parts = []
                parts = content_string.split('_')
                for part in parts:
                    if part.lower().startswith("page"):
                        page_number = part.split(" ")[-1]
                    else:
                        content_parts.append(part)
                
                if content_parts:
                    title = content_parts[0]
                    if len(content_parts) > 1: media = content_parts[1]
                    if len(content_parts) > 2: journalist = content_parts[2]

            if not title and not media and not journalist and not date:
                title = combined_text

            # --- PERBAIKAN DI SINI ---
            data['Category'].append(current_category)
            data['Date'].append(date)
            data['Title'].append(title)
            # Menggunakan .title() untuk kapitalisasi setiap awal kata
            data['Media'].append(media.title()) 
            data['Journalist'].append(journalist)
            data['Page Number'].append(page_number)
            data['Link'].append(link)

    return pd.DataFrame(data)

# ==============================================================================
# MAIN LOGIC
# ==============================================================================

# 1. Root Folder Path
root_folder_path = r"C:\Users\DELL\OneDrive\Work\Media Monitoring\1. ASSA\6. Juni"

# 2. Empty "In-Tray"
all_dataframes = []

print(f"Memulai Misi Pencarian di folder: '{root_folder_path}'")
# 3. Jalankan Misi Pencarian
for dirpath, dirnames, filenames in os.walk(root_folder_path):
    for filename in filenames:
        if filename.lower().endswith('.docx'):
            docx_path = os.path.join(dirpath, filename)
            print(f"  -> Menemukan dan memproses file: {filename}")
            daily_df = extract_data_from_docx(docx_path)
            if daily_df is not None and not daily_df.empty:
                all_dataframes.append(daily_df)

# 4. Checking Data
if all_dataframes:
    print("\nSemua file telah diproses. Menggabungkan data...")
    master_df = pd.concat(all_dataframes, ignore_index=True)

    master_df['Date'] = pd.to_datetime(master_df['Date'], errors='coerce').dt.date
    
    # 5. Decide names and location of final excel file 
    output_filename = os.path.basename(root_folder_path) + "_Compiled_Report.xlsx"
    excel_output_path = os.path.join(root_folder_path, output_filename)
    
    # 6. Saving Process
    with pd.ExcelWriter(excel_output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- ---
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',      
            'valign': 'vcenter',    
            'fg_color': '#FFFF00',
            'border': 1
        })
        
        date_format = workbook.add_format({'num_format': 'yyyy/mm/dd', 'border': 1, 'valign': 'vcenter'})
        
        cell_format = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'vcenter'})
        
        client_news_columns = [
            'Tanggal', 'Judul', 'Media', 'Page Number', 'Journalist',
            'Narsum', 'Jabatan', 'Narsum 2', 'Jabatan 2', 'Link'
        ]
        default_columns = [
            'Tanggal', 'Judul', 'Media', 'Page Number', 'Journalist', 'Link'
        ]
        
        master_df.rename(columns={'Date': 'Tanggal', 'Title': 'Judul'}, inplace=True)
        
        grouped_by_category = master_df.groupby('Category')
        print("Mulai menyimpan ke sheet Excel...")
        
        for category_name, group_data in grouped_by_category:
            safe_sheet_name = re.sub(r'[\\/*?:"<>|]', '', category_name)
            safe_sheet_name = safe_sheet_name[:31]
            print(f"  -> Menyimpan sheet: {safe_sheet_name}")
            
            if category_name == 'Client News':
                final_columns = client_news_columns
            else:
                final_columns = default_columns
            
            sheet_data = group_data.reindex(columns=final_columns)
            
            sheet_data.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False, startrow=1)
            
            worksheet = writer.sheets[safe_sheet_name]
            
            for col_num, value in enumerate(sheet_data.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            for idx, col_name in enumerate(sheet_data.columns):
                if col_name == 'Tanggal':
                    col_format = date_format
                else:
                    col_format = cell_format
                
                series = sheet_data[col_name]
                max_len = max((
                    series.astype(str).map(len).max(),
                    len(str(series.name))
                )) + 2
                final_width = max(min(max_len, 50), 10)
                worksheet.set_column(idx, idx, final_width, col_format)

    print(f"\n(SELESAI) Laporan bulanan berhasil dibuat di: '{excel_output_path}'")
else:

    print("\n(Peringatan) Tidak ada data yang berhasil diekstrak dari semua file Word yang ditemukan.")
