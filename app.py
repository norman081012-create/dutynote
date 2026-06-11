```python
    # === 替換 build_word_and_check_overflow 函數 ===
    def build_word_and_check_overflow(p_stations, p_new, p_out, lines, selected_date, selected_doc):
        """生成 Word 檔案，並在發生溢出時，將超出內容及簽章區移至新頁面。"""
        if not os.path.exists(TEMPLATE_PATH): raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。")
        doc = Document(TEMPLATE_PATH)
        
        # 1. 基礎數據填充 (日期、簽章、各護理站動態、出入院名單) --- 保持不變 ---
        roc_year = selected_date.year - 1911
        date_str = f"日期： {roc_year} 年 {selected_date.month:02d} 月 {selected_date.day:02d} 日"
        
        def apply_signature(p_element, doc_name):
            p_element.text = "" 
            run_label = p_element.add_run("值班醫師：")
            if doc_name:
                run_name = p_element.add_run(f"  {doc_name}")
                run_name.font.size = Pt(16)
                run_name.bold = True
                run_name.font.name = '標楷體'
                rPr_name = run_name._element.get_or_add_rPr()
                rPr_name.get_or_add_rFonts().set(qn('w:eastAsia'), '標楷體')
            p_element.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for p in doc.paragraphs:
            txt = p.text.replace(" ", "")
            if "日期" in txt and ("年" in txt or "月" in txt): p.text = date_str
            elif "值班醫師" in txt: apply_signature(p, selected_doc)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "值班醫師" in p.text.replace(" ", ""): apply_signature(p, selected_doc)
        
        for table in doc.tables:
            for row in table.rows:
                u_cells = get_unique_cells(row)
                row_txt = "".join([c.text for c in u_cells]).replace(" ", "")
                matched_st = None
                for kn in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                    if kn in row_txt: matched_st = kn; break
                if matched_st and matched_st in p_stations:
                    for idx, c in enumerate(u_cells):
                        if matched_st in re.sub(r'[\r\n\t]', '', c.text.replace(" ", "").replace(" ", "")) and idx+3 < len(u_cells):
                            safe_fill_cell(u_cells[idx+1], p_stations[matched_st][0], font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
                            safe_fill_cell(u_cells[idx+2], p_stations[matched_st][1], font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
                            safe_fill_cell(u_cells[idx+3], p_stations[matched_st][2], font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

        # 處理出入院病人 (刪除空白行) --- 保持不變 ---
        for table in doc.tables:
            blank_new_rows, blank_out_rows, section = [], [], None
            name_col_new, name_col_out = 0, 0
            for row in table.rows:
                u_cells = get_unique_cells(row)
                row_txt = "".join(c.text for c in u_cells).replace(" ", "")
                if "姓名" in row_txt and "病歷" in row_txt:
                    if "燈號" in row_txt or "強制" in row_txt: 
                        section, name_col_new = "new", next((i for i, c in enumerate(u_cells) if "姓名" in c.text.replace(" ", "")), 0)
                    else: 
                        section, name_col_out = "out", next((i for i, c in enumerate(u_cells) if "姓名" in c.text.replace(" ", "")), 0)
                elif "出院病人" in row_txt or "危險評估" in row_txt: section = None
                elif section == "new" and name_col_new < len(u_cells):
                    if re.sub(r'[\r\n\t\s_0]', '', u_cells[name_col_new].text) == "": blank_new_rows.append((row, name_col_new))
                elif section == "out" and name_col_out < len(u_cells):
                    if re.sub(r'[\r\n\t\s_0]', '', u_cells[name_col_out].text) == "": blank_out_rows.append((row, name_col_out))

            # 用戶圖片顯示「自殺顧慮」表格，這裡刪除空白行 --- 保持不變 ---
            while len(p_new) > len(blank_new_rows) and blank_new_rows:
                last_row, col = blank_new_rows[-1]
                new_tr = deepcopy(last_row._tr)
                last_row._tr.addnext(new_tr)
                blank_new_rows.append((_Row(new_tr, last_row._parent), col))
                
            while len(p_out) > len(blank_out_rows) and blank_out_rows:
                last_row, col = blank_out_rows[-1]
                new_tr = deepcopy(last_row._tr)
                last_row._tr.addnext(new_tr)
                blank_out_rows.append((_Row(new_tr, last_row._parent), col))

            for i, (row, col_idx) in enumerate(blank_new_rows):
                if i < len(p_new):
                    u_cells = get_unique_cells(row)
                    for k in range(min(len(p_new[i]), len(u_cells))):
                        tc = col_idx + k if k < 6 else col_idx + k + 1
                        if tc < len(u_cells): safe_fill_cell(u_cells[tc], p_new[i][k], font_size=10)
                else:
                    try: row._element.getparent().remove(row._element)
                    except: pass
                    
            for i, (row, col_idx) in enumerate(blank_out_rows):
                if i < len(p_out):
                    u_cells = get_unique_cells(row)
                    for k in range(min(len(p_out[i]), len(u_cells))):
                        if col_idx + k < len(u_cells): safe_fill_cell(u_cells[col_idx + k], p_out[i][k], font_size=10)
                else:
                    try: row._element.getparent().remove(row._element)
                    except: pass

        # 2. 交班事項排版更改核心邏輯 ---
        
        # 計算交班事項字串 (事項塊)
        all_chunks_to_fill = []
        for line in lines:
            if line == "": all_chunks_to_fill.append("")
            else: all_chunks_to_fill.extend(visual_smart_chunker(line, max_visual_width=78))

        # 找出目標表格、「病房特殊狀況及處理」標頭行、以及「討論與講評」行
        target_table, header_row_idx, discuss_row_idx = None, -1, -1
        for table in doc.tables:
            for idx, row in enumerate(table.rows):
                u_cells = get_unique_cells(row)
                if not u_cells: continue
                row_txt = u_cells[0].text.replace(" ", "")
                if "項次" in row_txt and "病房特殊狀況及處理" in row_txt: header_row_idx = idx
                if "討論與講評" in row_txt: discuss_row_idx = idx
                if header_row_idx != -1 and discuss_row_idx != -1: target_table = table; break
            if target_table: break

        if not target_table or header_row_idx == -1 or discuss_row_idx == -1:
            raise ValueError("無法在 Word 模板中找到完整的交班表格結構 (標頭或討論行)。")

        # 保存（使用 deepcopy）「討論與講評」及以下所有行的 XML，並從原始表格中刪除
        discuss_rows_xml_str = []
        rows_to_save = []
        for r_idx in range(discuss_row_idx, len(target_table.rows)):
            discuss_rows_xml_str.append(deepcopy(target_table.rows[r_idx]._element))
            rows_to_save.append(target_table.rows[r_idx]._element)
        
        for r_element in rows_to_save:
            target_table._element.remove(r_element)

        # 計算原始表格容量 (不加新行，直到「討論與講評」行之前客滿)
        start_row_idx = header_row_idx + 1
        capacity = discuss_row_idx - start_row_idx
        
        is_overflow = False
        chunks_on_first_page = all_chunks_to_fill
        chunks_to_move_to_new_pages = []

        if len(all_chunks_to_fill) > capacity:
            is_overflow = True
            chunks_on_first_page = all_chunks_to_fill[:capacity]
            chunks_to_move_to_new_pages = all_chunks_to_fill[capacity:]

        # 填充原始表格客滿 (第一頁)
        for i, chunk_text in enumerate(chunks_on_first_page):
            target_cell = get_unique_cells(target_table.rows[start_row_idx + i])[0]
            safe_fill_cell(target_cell, chunk_text, font_size=12)

        final_table = target_table  # 預設最後一頁表格是原始表格
        
        # 如果發生溢出，將內容移動到新頁面的新表格中
        if is_overflow:
            # 獲取標頭行的 XML 副本 (例如：項次、病房特殊狀況及處理)
            header_tr_xml = deepcopy(target_table.rows[header_row_idx]._element)
            
            # 對溢出事項塊進行迭代，每頁填充 10 塊或在最後一個塊
            for chunk_idx, chunk_text in enumerate(chunks_to_move_to_new_pages):
                # 每頁 10 塊或在第一個溢出塊時添加新頁 (用戶圖片排版需求)
                if chunk_idx % 10 == 0:
                    doc.add_page_break()
                    
                    # 創建一個與原始交班表格具有相同結構的新表格 (具有相同的列標頭)
                    # 我們使用 2 列，這需要根據您的 template.docx 實際列數進行調整
                    new_table = doc.add_table(rows=1, cols=len(header_tr_xml.xpath('.//w:tc'))) 
                    # 複製標頭行 XML
                    new_table.rows[0]._element.getparent().replace(new_table.rows[0]._element, header_tr_xml)
                    final_table = new_table # 更新最後一頁表格

                # 創建事項行，填充事项块内容
                # 獲取標頭行的 XML 副本 (使用深拷貝)
                header_tr_xml = deepcopy(target_table.rows[header_row_idx]._element)
                
                # 拷貝一個事項行 XML 結構 (從原始事項行中獲取)
                source_row_xml = deepcopy(target_table.rows[start_row_idx]._element)
                final_table._element.append(source_row_xml)
                
                # 將事項填充到新行中
                target_cell = get_unique_cells(final_table.rows[-1])[0]
                safe_fill_cell(target_cell, chunk_text, font_size=12)

        # 3. 最後一頁表格處理：附加上「討論與講評」和所有簽章行 (如圖片所示) ---
        
        # 迭代保存的「討論與講評」和簽章行的 XML，並將其副本添加到最後一個表格的末尾
        for row_xml in discuss_rows_xml_str:
            final_table._element.append(deepcopy(row_xml))

        # 4. 返回文檔串流 ---
        stream = io.BytesIO(); doc.save(stream); stream.seek(0)
        return stream, is_overflow

```
