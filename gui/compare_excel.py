import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os
import subprocess
import stat
import sys
import queue
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
import datetime

# 全局队列：用于子线程与GUI线程通信
log_queue = queue.Queue()
progress_queue = queue.Queue()

# 默认主题设置
DEFAULT_APPEARANCE_MODE = "light"  # "dark", "light", "system"
DEFAULT_COLOR_THEME = "blue"     # "blue", "green", "dark-blue"

# 初始化主题
ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

def compare_excel_files(baseline_path, compare_path, output_baseline_path, output_compare_path, results_folder, original_filename, timestamp, header_row=3, key_fields=None, stop_event=None):
    # 获取文件夹名称用于标识
    baseline_folder = os.path.basename(os.path.dirname(baseline_path))
    compare_folder = os.path.basename(os.path.dirname(compare_path))
    
    # 定义颜色样式
    fill_changed = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色：数值变化
    fill_added = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")      # 绿色：新增（在基准基础上）
    fill_deleted = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")    # 红色：删除（在基准基础上）

    log_queue.put(f"正在加载文件: {baseline_path} 和 {compare_path} ...")
    
    try:
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        # my文件夹是基准文件
        wb_baseline = openpyxl.load_workbook(baseline_path, data_only=True)  # 只加载数据，不加载公式
        wb_compare = openpyxl.load_workbook(compare_path, data_only=True)
    except FileNotFoundError as e:
        log_queue.put(f"错误：找不到文件 - {e}")
        return False
    except Exception as e:
        log_queue.put(f"加载文件时出错: {e}")
        return False

    # 1. 选择工作表
    log_queue.put(f"\n【{baseline_folder}文件夹】工作表列表: {wb_baseline.sheetnames}")
    log_queue.put(f"【{compare_folder}文件夹】工作表列表: {wb_compare.sheetnames}")
    
    # 默认使用第一个工作表
    ws_baseline = wb_baseline.active
    ws_compare = wb_compare.active
    log_queue.put(f"默认比较第一个工作表: {ws_baseline.title} ({baseline_folder}) vs {ws_compare.title} ({compare_folder})")

    # 2. 获取实际使用的范围
    baseline_max_row = ws_baseline.max_row
    baseline_max_col = ws_baseline.max_column
    compare_max_row = ws_compare.max_row
    compare_max_col = ws_compare.max_column

    log_queue.put(f"开始比较 ({baseline_folder}文件夹: {baseline_max_row}行 x {baseline_max_col}列, {compare_folder}文件夹: {compare_max_row}行 x {compare_max_col}列)...")
    
    # 检查列数是否一致
    if baseline_max_col != compare_max_col:
        log_queue.put(f"警告：两个文件的列数不一致！基准文件：{baseline_max_col}列，比较文件：{compare_max_col}列")

    # 3. 预先获取所有单元格值
    cells_baseline = {}
    cells_compare = {}
    
    # 获取基准文件所有单元格值
    for r in range(1, baseline_max_row + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        for c in range(1, baseline_max_col + 1):
            cells_baseline[(r, c)] = ws_baseline.cell(row=r, column=c).value
    
    # 获取比较文件所有单元格值
    for r in range(1, compare_max_row + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        for c in range(1, compare_max_col + 1):
            cells_compare[(r, c)] = ws_compare.cell(row=r, column=c).value
    
    # 4. 基于关键字段的行匹配算法
    def get_col_content(col_num, cells, max_row):
        """获取一列的所有单元格内容，作为比较的键"""
        return tuple(cells.get((r, col_num), None) for r in range(1, max_row + 1))
    
    # 如果没有提供关键字段，默认使用前三列作为特征列
    if not key_fields:
        # 获取表头行的列名
        header_values = [cells_baseline.get((header_row, c), "").strip() for c in range(1, min(baseline_max_col + 1, 4))]
        key_fields = [v for v in header_values if v]  # 过滤空值
        if len(key_fields) < 3:
            # 如果表头不足3个有效列名，使用默认列名
            key_fields = [f"列{c}" for c in range(1, min(baseline_max_col + 1, 4))]
    
    # 从指定表头行获取关键字段的列索引
    def find_key_columns(cells, max_col, header_row_num, key_field_names):
        """从指定行查找关键字段的列索引"""
        key_cols = {}
        # 先获取表头行的所有列名
        header_values = {}
        for col in range(1, max_col + 1):
            cell_value = cells.get((header_row_num, col), "").strip()
            header_values[cell_value] = col
        
        # 查找关键字段的列索引
        for field in key_field_names:
            if field in header_values:
                key_cols[field] = header_values[field]
            else:
                # 如果找不到字段名，尝试直接使用列索引
                try:
                    col_idx = int(field.replace("列", ""))
                    if 1 <= col_idx <= max_col:
                        key_cols[field] = col_idx
                except ValueError:
                    pass
        return key_cols
    
    # 查找基准文件和比较文件的关键字段列索引
    key_cols_baseline = find_key_columns(cells_baseline, baseline_max_col, header_row, key_fields)
    key_cols_compare = find_key_columns(cells_compare, compare_max_col, header_row, key_fields)
    
    log_queue.put(f"\n基准文件关键字段列索引: {key_cols_baseline}")
    log_queue.put(f"比较文件关键字段列索引: {key_cols_compare}")
    
    # 检查是否找到所有关键字段
    has_all_keys_baseline = all(field in key_cols_baseline for field in key_fields)
    has_all_keys_compare = all(field in key_cols_compare for field in key_fields)
    
    # 行匹配：基准行号 -> 比较行号
    row_mapping = {}
    
    if has_all_keys_baseline and has_all_keys_compare:
        log_queue.put("\n使用关键字段进行行匹配...")
        
        # 构建行关键字映射：关键字 -> 行号
        def build_row_key_map(cells, max_row, key_cols, data_start_row):
            row_key_map = {}
            for row in range(data_start_row, max_row + 1):  # 从数据行开始
                key_values = tuple(cells.get((row, key_cols[field]), None) for field in key_fields)
                # 只有当所有关键字段都有值时才进行映射
                if all(v is not None for v in key_values):
                    row_key_map[key_values] = row
            return row_key_map
        
        # 数据行从表头行的下一行开始
        data_start_row = header_row + 1
        row_key_map_baseline = build_row_key_map(cells_baseline, baseline_max_row, key_cols_baseline, data_start_row)
        row_key_map_compare = build_row_key_map(cells_compare, compare_max_row, key_cols_compare, data_start_row)
        
        # 建立行映射：基准行 -> 比较行
        for key in row_key_map_baseline:
            if key in row_key_map_compare:
                row_baseline = row_key_map_baseline[key]
                row_compare = row_key_map_compare[key]
                row_mapping[row_baseline] = row_compare
        
        log_queue.put(f"基于关键字段匹配到 {len(row_mapping)} 行")
    else:
        log_queue.put("\n无法找到所有关键字段，使用默认行匹配...")
        # 原来的行匹配逻辑
        def get_row_content(row_num, cells, max_col):
            """获取一行的所有单元格内容，作为比较的键"""
            return tuple(cells.get((row_num, c), None) for c in range(1, max_col + 1))
        
        # 构建行内容映射
        row_contents_baseline = {r: get_row_content(r, cells_baseline, baseline_max_col) for r in range(1, baseline_max_row + 1)}
        row_contents_compare = {r: get_row_content(r, cells_compare, compare_max_col) for r in range(1, compare_max_row + 1)}
        
        # 先找到完全匹配的行
        for row_baseline in row_contents_baseline:
            # 检查是否需要停止
            if stop_event and stop_event.is_set():
                log_queue.put("操作已取消")
                return False
                
            content_baseline = row_contents_baseline[row_baseline]
            for row_compare in row_contents_compare:
                if row_compare not in row_mapping.values() and content_baseline == row_contents_compare[row_compare]:
                    row_mapping[row_baseline] = row_compare
                    break
        
        # 如果没有找到足够的匹配，使用简单的索引映射
        if len(row_mapping) < min(baseline_max_row, compare_max_row) // 2:
            min_rows = min(baseline_max_row, compare_max_row)
            row_mapping = {r: r for r in range(1, min_rows + 1)}
    
    # 5. 比较单元格
    changes_count = 0
    
    # 定义关键字段列索引集合，避免重新计算
    key_col_set_baseline = set(key_cols_baseline.values()) if has_all_keys_baseline else set()
    key_col_set_compare = set(key_cols_compare.values()) if has_all_keys_compare else set()
    
    # 只比较匹配的行（基于关键字段匹配的行）
    log_queue.put("\n开始比较匹配行的单元格差异...")
    
    # 创建列映射（基于列名匹配）
    def create_col_name_map(header_row_num):
        col_name_map = {}
        # 先获取基准文件的列名映射
        baseline_col_names = {}
        for col_b in range(1, baseline_max_col + 1):
            col_name_b = cells_baseline.get((header_row_num, col_b), "").strip()
            if col_name_b:
                baseline_col_names[col_name_b] = col_b
        
        # 然后在比较文件中查找相同列名
        for col_c in range(1, compare_max_col + 1):
            col_name_c = cells_compare.get((header_row_num, col_c), "").strip()
            if col_name_c in baseline_col_names:
                col_name_map[baseline_col_names[col_name_c]] = col_c
        
        # 如果没有找到足够的匹配，使用简单的索引映射
        if len(col_name_map) < min(baseline_max_col, compare_max_col) // 2:
            min_cols = min(baseline_max_col, compare_max_col)
            col_name_map = {c: c for c in range(1, min_cols + 1)}
        
        return col_name_map
    
    col_name_map = create_col_name_map(header_row)
    
    for row_baseline in row_mapping:
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        row_compare = row_mapping[row_baseline]
        
        # 比较匹配的列
        for col_baseline, col_compare in col_name_map.items():
            # 跳过关键字段列（它们已经匹配，不需要比较）
            if col_baseline in key_col_set_baseline or col_compare in key_col_set_compare:
                continue
            
            val_baseline = cells_baseline.get((row_baseline, col_baseline), None)
            val_compare = cells_compare.get((row_compare, col_compare), None)
            
            # 只在值不同时标记为黄色（数值变化）
            if val_baseline != val_compare:
                ws_baseline.cell(row=row_baseline, column=col_baseline).fill = fill_changed
                ws_compare.cell(row=row_compare, column=col_compare).fill = fill_changed
                changes_count += 1
    
    # 6. 标记新增行和删除行
    log_queue.put("\n开始标记新增行和删除行...")
    
    # 获取所有数据行的关键字映射
    def get_all_row_keys(cells, max_row, key_cols, data_start_row):
        """获取所有数据行的关键字映射"""
        all_row_keys = {}
        for row in range(data_start_row, max_row + 1):  # 从数据行开始
            key_values = tuple(cells.get((row, key_cols[field]), None) for field in key_fields)
            if all(v is not None for v in key_values):
                all_row_keys[key_values] = row
        return all_row_keys
    
    if has_all_keys_baseline and has_all_keys_compare:
        # 获取所有数据行的关键字映射
        data_start_row = header_row + 1
        all_baseline_keys = get_all_row_keys(cells_baseline, baseline_max_row, key_cols_baseline, data_start_row)
        all_compare_keys = get_all_row_keys(cells_compare, compare_max_row, key_cols_compare, data_start_row)
        
        # 标记删除行（基准文件中有，比较文件中没有）
        deleted_rows = 0
        for key, row_baseline in all_baseline_keys.items():
            # 检查是否需要停止
            if stop_event and stop_event.is_set():
                log_queue.put("操作已取消")
                return False
                
            if key not in all_compare_keys:
                # 标记整行为绿色
                for col in range(1, baseline_max_col + 1):
                    ws_baseline.cell(row=row_baseline, column=col).fill = fill_added
                changes_count += 1
                deleted_rows += 1
        log_queue.put(f"已标记 {deleted_rows} 行删除（绿色）")
        
        # 标记新增行（比较文件中有，基准文件中没有）
        added_rows = 0
        for key, row_compare in all_compare_keys.items():
            # 检查是否需要停止
            if stop_event and stop_event.is_set():
                log_queue.put("操作已取消")
                return False
                
            if key not in all_baseline_keys:
                # 标记整行为红色
                for col in range(1, compare_max_col + 1):
                    ws_compare.cell(row=row_compare, column=col).fill = fill_deleted
                changes_count += 1
                added_rows += 1
        log_queue.put(f"已标记 {added_rows} 行新增（红色）")
    else:
        # 使用简单的行匹配来标记新增和删除行
        log_queue.put("\n使用简单匹配标记新增和删除行...")
        
        # 标记删除行（基准文件中有，比较文件中没有对应的行）
        deleted_rows = 0
        for row_baseline in range(1, baseline_max_row + 1):
            # 检查是否需要停止
            if stop_event and stop_event.is_set():
                log_queue.put("操作已取消")
                return False
                
            if row_baseline not in row_mapping:
                # 标记整行为绿色
                for col in range(1, baseline_max_col + 1):
                    ws_baseline.cell(row=row_baseline, column=col).fill = fill_added
                changes_count += 1
                deleted_rows += 1
        log_queue.put(f"已标记 {deleted_rows} 行删除（绿色）")
        
        # 标记新增行（比较文件中有，基准文件中没有对应的行）
        added_rows = 0
        mapped_compare_rows = set(row_mapping.values())
        for row_compare in range(1, compare_max_row + 1):
            # 检查是否需要停止
            if stop_event and stop_event.is_set():
                log_queue.put("操作已取消")
                return False
                
            if row_compare not in mapped_compare_rows:
                # 标记整行为红色
                for col in range(1, compare_max_col + 1):
                    ws_compare.cell(row=row_compare, column=col).fill = fill_deleted
                changes_count += 1
                added_rows += 1
        log_queue.put(f"已标记 {added_rows} 行新增（红色）")

    # 6. 保存my和from的比较结果文件
    log_queue.put("\n正在保存结果文件...")
    
    try:
        wb_baseline.save(output_baseline_path)
        wb_compare.save(output_compare_path)
    except Exception as e:
        log_queue.put(f"保存结果文件时出错: {e}")
        return False
    
    # 7. 生成差异结果文件
    log_queue.put("\n正在生成差异结果文件...")
    
    try:
        # 直接使用保存后的my_比较结果文件作为差异结果的基础
        # 这样可以确保格式完全一致
        wb_diff = openpyxl.load_workbook(output_baseline_path)
        ws_diff = wb_diff.active
        ws_diff.title = "差异比较结果"
        
        # 重新加载保存后的文件以获取准确的格式信息
        wb_baseline_saved = openpyxl.load_workbook(output_baseline_path)
        ws_baseline_saved = wb_baseline_saved.active
        
        wb_compare_saved = openpyxl.load_workbook(output_compare_path)
        ws_compare_saved = wb_compare_saved.active
    except Exception as e:
        log_queue.put(f"加载保存后的文件时出错: {e}")
        return False
    
    # 创建一个字典来快速查找基准行是否存在
    baseline_key_set = set()
    key_to_row = {}
    
    # 获取基准文件中所有行的关键字段值
    for row_baseline in range(4, ws_baseline_saved.max_row + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        key_values = tuple(ws_baseline_saved.cell(row=row_baseline, column=key_cols_baseline[field]).value for field in key_fields)
        if all(v is not None for v in key_values):
            baseline_key_set.add(key_values)
            key_to_row[key_values] = row_baseline
    
    # 收集比较文件中的新增行（红色行）
    added_rows = []
    for row_compare in range(4, ws_compare_saved.max_row + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        # 获取当前行的关键字段值
        key_values = tuple(ws_compare_saved.cell(row=row_compare, column=key_cols_compare[field]).value for field in key_fields)
        if not all(v is not None for v in key_values):
            continue
        
        # 检查是否为新增行（红色）
        is_added_row = False
        first_cell = ws_compare_saved.cell(row=row_compare, column=1)
        if first_cell.fill.start_color.rgb == fill_deleted.start_color.rgb:
            is_added_row = True
        
        if is_added_row:
            # 获取当前行在比较文件中的上一行关键字段值
            prev_key_values = None
            if row_compare > 4:
                prev_key_values = tuple(ws_compare_saved.cell(row=row_compare - 1, column=key_cols_compare[field]).value for field in key_fields)
            
            added_rows.append((key_values, row_compare, prev_key_values))
    
    # 计算需要插入的行数，提前插入空白行
    for i in range(len(added_rows)):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        # 在差异结果文件末尾插入一行
        ws_diff.append(['' for _ in range(baseline_max_col)])
    
    # 将新增行插入到正确位置
    for key_values, row_compare, prev_key_values in added_rows:
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        # 找到插入位置
        insert_row = ws_diff.max_row
        if prev_key_values and prev_key_values in key_to_row:
            insert_row = key_to_row[prev_key_values] + 1
        
        # 插入空白行
        ws_diff.insert_rows(insert_row)
        
        # 更新key_to_row字典
        for k, v in key_to_row.items():
            if v >= insert_row:
                key_to_row[k] = v + 1
        
        # 使用基准文件的第4行作为模板，复制其格式
        from openpyxl.styles import Font, Border, Side, Alignment
        template_row = 4
        
        # 先复制模板行的格式到新插入的行
        for col in range(1, baseline_max_col + 1):
            # 获取模板单元格
            template_cell = ws_baseline_saved.cell(row=template_row, column=col)
            new_cell = ws_diff.cell(row=insert_row, column=col)
            
            # 复制数字格式
            new_cell.number_format = template_cell.number_format
            
            # 复制字体样式（创建新的Font对象）
            template_font = template_cell.font
            new_font = Font(
                name=template_font.name,
                size=template_font.size,
                bold=template_font.bold,
                italic=template_font.italic,
                vertAlign=template_font.vertAlign,
                underline=template_font.underline,
                strike=template_font.strike,
                color=template_font.color
            )
            new_cell.font = new_font
            
            # 复制边框样式（创建新的Border对象）
            template_border = template_cell.border
            new_border = Border(
                left=template_border.left,
                right=template_border.right,
                top=template_border.top,
                bottom=template_border.bottom,
                diagonal=template_border.diagonal,
                diagonalUp=template_border.diagonalUp,
                diagonalDown=template_border.diagonalDown
            )
            new_cell.border = new_border
            
            # 复制对齐方式（创建新的Alignment对象）
            template_alignment = template_cell.alignment
            new_alignment = Alignment(
                horizontal=template_alignment.horizontal,
                vertical=template_alignment.vertical,
                textRotation=template_alignment.textRotation,
                wrapText=template_alignment.wrapText,
                shrinkToFit=template_alignment.shrinkToFit,
                indent=template_alignment.indent,
                relativeIndent=template_alignment.relativeIndent,
                justifyLastLine=template_alignment.justifyLastLine,
                readingOrder=template_alignment.readingOrder,
                text_rotation=template_alignment.text_rotation,
                wrap_text=template_alignment.wrap_text,
                shrink_to_fit=template_alignment.shrink_to_fit
            )
            new_cell.alignment = new_alignment
        
        # 然后填入新增行的数据
        for col in range(1, baseline_max_col + 1):
            # 获取基准文件中对应的列名
            col_name_b = ws_baseline_saved.cell(row=3, column=col).value.strip() if ws_baseline_saved.cell(row=3, column=col).value else ""
            if not col_name_b:
                continue
            
            # 在比较文件中查找对应的列
            for c in range(1, ws_compare_saved.max_column + 1):
                col_name_c = ws_compare_saved.cell(row=3, column=c).value.strip() if ws_compare_saved.cell(row=3, column=c).value else ""
                if col_name_c == col_name_b:
                    # 填入数据
                    value = ws_compare_saved.cell(row=row_compare, column=c).value
                    ws_diff.cell(row=insert_row, column=col, value=value)
                    break
        
        # 最后将整行设置为红色填充
        for col in range(1, baseline_max_col + 1):
            cell = ws_diff.cell(row=insert_row, column=col)
            cell.fill = fill_deleted
    
    # 4. 复制基准文件的列宽设置，确保格式完全一致
    for col in range(1, ws_baseline_saved.max_column + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        col_letter = get_column_letter(col)
        if col_letter in ws_baseline_saved.column_dimensions:
            ws_diff.column_dimensions[col_letter].width = ws_baseline_saved.column_dimensions[col_letter].width
    
    # 5. 复制基准文件的行高设置
    for row in range(1, ws_baseline_saved.max_row + 1):
        # 检查是否需要停止
        if stop_event and stop_event.is_set():
            log_queue.put("操作已取消")
            return False
            
        if row in ws_baseline_saved.row_dimensions:
            ws_diff.row_dimensions[row].height = ws_baseline_saved.row_dimensions[row].height
    
    # 保存差异结果文件
    diff_output_path = os.path.join(results_folder, f"{original_filename}_差异结果_{timestamp}.xlsx")
    try:
        wb_diff.save(diff_output_path)
    except Exception as e:
        log_queue.put(f"保存差异结果文件时出错: {e}")
        return False
    
    # 8. 设置文件为只读
    log_queue.put("\n正在设置文件只读属性...")
    try:
        # 获取当前文件权限
        baseline_stat = os.stat(output_baseline_path)
        compare_stat = os.stat(output_compare_path)
        diff_stat = os.stat(diff_output_path)
        
        # 在Windows上设置只读属性
        if os.name == 'nt':
            # 使用Windows命令设置只读
            subprocess.run(['attrib', '+r', output_baseline_path], check=True)
            subprocess.run(['attrib', '+r', output_compare_path], check=True)
            subprocess.run(['attrib', '+r', diff_output_path], check=True)
        else:
            # 在Linux/macOS上设置只读
            os.chmod(output_baseline_path, baseline_stat.st_mode & ~stat.S_IWUSR & ~stat.S_IWGRP & ~stat.S_IWOTH)
            os.chmod(output_compare_path, compare_stat.st_mode & ~stat.S_IWUSR & ~stat.S_IWGRP & ~stat.S_IWOTH)
            os.chmod(diff_output_path, diff_stat.st_mode & ~stat.S_IWUSR & ~stat.S_IWGRP & ~stat.S_IWOTH)
        
        log_queue.put("结果文件已设置为只读属性")
    except Exception as e:
        log_queue.put(f"设置只读属性时出错: {e}")
    
    log_queue.put(f"\n比较完成！共发现 {changes_count} 处差异。")
    log_queue.put(f"已生成带颜色标记的文件至: {output_baseline_path}")
    log_queue.put(f"已生成带颜色标记的文件至: {output_compare_path}")
    log_queue.put(f"已生成差异结果文件至: {diff_output_path}")
    
    # 7. 自动打开文件
    log_queue.put("\n正在打开生成的文件...")
    try:
        subprocess.Popen(['start', '', output_baseline_path], shell=True)
        subprocess.Popen(['start', '', output_compare_path], shell=True)
        subprocess.Popen(['start', '', diff_output_path], shell=True)
        log_queue.put("文件已成功打开！")
    except Exception as e:
        log_queue.put(f"打开文件时出错: {e}")
    
    return True



class StdoutRedirector:
    """重定向stdout到GUI的Text组件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        log_queue.put(message)

    def flush(self):
        pass

class ExcelCompareGUI(ctk.CTk):
    """Excel文件比较工具GUI界面"""
    def __init__(self):
        super().__init__()
        self.title("Excel文件比较工具")
        self.geometry("1200x800")
        self.minsize(1000, 700)
        
        # 配置变量
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.parent_dir = os.path.dirname(self.current_dir)
        self.results_folder = os.path.join(self.parent_dir, "results")
        os.makedirs(self.results_folder, exist_ok=True)
        
        self.baseline_file = ""
        self.compare_file = ""
        self.running = False
        self.stop_event = threading.Event()
        self.worker_thread = None
        
        # 初始化界面
        self._init_widgets()
        
        # 重定向stdout
        self._redirect_stdout()
        
        # 启动队列监听
        self._listen_queues()
    
    def _init_widgets(self):
        """初始化GUI组件"""
        # 创建主容器
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill="both", expand=True, padx=0, pady=0)
        
        # 顶部标题栏
        header_frame = ctk.CTkFrame(main_container, fg_color=("gray90", "gray20"), height=100)
        header_frame.pack(fill="x", padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # 标题和主题选择
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(fill="x", padx=20, pady=10)
        
        # 标题
        title_label = ctk.CTkLabel(
            title_frame, 
            text="Excel文件比较工具",
            font=("微软雅黑", 26, "bold"),
            text_color=("#1f77b4", "#64b5f6")
        )
        title_label.pack(anchor="w", side="left")
        
        # 主题选择
        theme_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        theme_frame.pack(anchor="e", side="right")
        
        ctk.CTkLabel(
            theme_frame,
            text="主题:",
            font=("微软雅黑", 12),
            text_color=("gray50", "gray70")
        ).pack(side="left", padx=(0, 10))
        
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            theme_frame,
            values=["light", "dark", "system"],
            command=self._change_appearance_mode_event,
            font=("微软雅黑", 12),
            width=120
        )
        self.appearance_mode_optionemenu.set(DEFAULT_APPEARANCE_MODE)
        self.appearance_mode_optionemenu.pack(side="left", padx=(0, 10))
        
        self.color_theme_optionemenu = ctk.CTkOptionMenu(
            theme_frame,
            values=["blue", "green", "dark-blue"],
            command=self._change_color_theme_event,
            font=("微软雅黑", 12),
            width=120
        )
        self.color_theme_optionemenu.set(DEFAULT_COLOR_THEME)
        self.color_theme_optionemenu.pack(side="left")
        
        # 主内容区
        content_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 左侧面板（文件选择和操作）
        left_panel = ctk.CTkFrame(content_frame, fg_color=("gray86", "gray17"))
        left_panel.pack(side="left", fill="y", expand=False, padx=(0, 10))
        left_panel.configure(width=300)
        
        # 文件选择区
        file_section = ctk.CTkFrame(left_panel, fg_color="transparent")
        file_section.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(
            file_section, 
            text="文件选择", 
            font=("微软雅黑", 16, "bold")
        ).pack(anchor="w", pady=(0, 10))
        
        # 基准文件选择
        baseline_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        baseline_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            baseline_frame, 
            text="基准文件:", 
            width=100,
            font=("微软雅黑", 12)
        ).pack(side="left", anchor="center")
        
        self.baseline_entry = ctk.CTkEntry(baseline_frame, font=("微软雅黑", 12))
        self.baseline_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkButton(
            baseline_frame, 
            text="浏览", 
            width=60,
            font=("微软雅黑", 12),
            command=self._browse_baseline_file
        ).pack(side="left", padx=5)
        
        # 比较文件选择
        compare_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        compare_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            compare_frame, 
            text="比较文件:", 
            width=100,
            font=("微软雅黑", 12)
        ).pack(side="left", anchor="center")
        
        self.compare_entry = ctk.CTkEntry(compare_frame, font=("微软雅黑", 12))
        self.compare_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkButton(
            compare_frame, 
            text="浏览", 
            width=60,
            font=("微软雅黑", 12),
            command=self._browse_compare_file
        ).pack(side="left", padx=5)
        
        # 结果文件夹显示
        result_folder_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        result_folder_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            result_folder_frame, 
            text="结果文件夹:", 
            width=100,
            font=("微软雅黑", 12)
        ).pack(side="left", anchor="center")
        
        self.result_folder_entry = ctk.CTkEntry(result_folder_frame, font=("微软雅黑", 12))
        self.result_folder_entry.insert(0, self.results_folder)
        self.result_folder_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkButton(
            result_folder_frame, 
            text="浏览", 
            width=60,
            font=("微软雅黑", 12),
            command=self._browse_result_folder
        ).pack(side="left", padx=5)
        
        # 配置选项区
        config_section = ctk.CTkFrame(left_panel, fg_color="transparent")
        config_section.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(
            config_section, 
            text="比较配置", 
            font=("微软雅黑", 16, "bold")
        ).pack(anchor="w", pady=(0, 10))
        
        # 表头行号选择
        header_row_frame = ctk.CTkFrame(config_section, fg_color="transparent")
        header_row_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            header_row_frame, 
            text="表头行号:", 
            width=100,
            font=("微软雅黑", 12)
        ).pack(side="left", anchor="center")
        
        self.header_row_var = ctk.StringVar(value="3")
        self.header_row_entry = ctk.CTkEntry(header_row_frame, textvariable=self.header_row_var, font=("微软雅黑", 12), width=60)
        self.header_row_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(
            header_row_frame, 
            text="(表头所在行，从1开始)", 
            font=("微软雅黑", 10),
            text_color="gray50"
        ).pack(side="left", anchor="center", padx=5)
        
        # 特征列选择
        feature_cols_frame = ctk.CTkFrame(config_section, fg_color="transparent")
        feature_cols_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            feature_cols_frame, 
            text="特征列:", 
            width=100,
            font=("微软雅黑", 12)
        ).pack(side="left", anchor="center")
        
        self.feature_cols_var = ctk.StringVar(value="前三列")
        self.feature_cols_entry = ctk.CTkEntry(feature_cols_frame, textvariable=self.feature_cols_var, font=("微软雅黑", 12))
        self.feature_cols_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkButton(
            feature_cols_frame, 
            text="选择", 
            width=60,
            font=("微软雅黑", 12),
            command=self._select_feature_columns
        ).pack(side="left", padx=5)
        
        ctk.CTkLabel(
            config_section, 
            text="提示: 特征列用于判断行的增删变化，默认使用前三列，最多支持6列", 
            font=("微软雅黑", 10),
            text_color="gray50"
        ).pack(anchor="w", pady=(5, 0))
        
        # 操作按钮区
        button_section = ctk.CTkFrame(left_panel, fg_color="transparent")
        button_section.pack(fill="x", padx=15, pady=15)
        
        self.start_button = ctk.CTkButton(
            button_section, 
            text="开始比较", 
            font=("微软雅黑", 16, "bold"),
            height=50,
            fg_color="#4CAF50",
            hover_color="#45a049",
            command=self._start_compare
        )
        self.start_button.pack(fill="x", pady=5)
        
        self.stop_button = ctk.CTkButton(
            button_section, 
            text="停止", 
            font=("微软雅黑", 16, "bold"),
            height=50,
            fg_color="#f44336",
            hover_color="#da190b",
            command=self._stop_compare,
            state="disabled"
        )
        self.stop_button.pack(fill="x", pady=5)
        
        # 右侧面板（日志显示）
        right_panel = ctk.CTkFrame(content_frame, fg_color=("gray86", "gray17"))
        right_panel.pack(side="right", fill="both", expand=True)
        
        # 日志标题
        log_title_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        log_title_frame.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkLabel(
            log_title_frame, 
            text="操作日志", 
            font=("微软雅黑", 16, "bold")
        ).pack(anchor="w")
        
        # 日志显示区域
        log_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        log_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=("微软雅黑", 12),
            wrap="word"
        )
        self.log_text.pack(fill="both", expand=True)
        
        # 滚动条
        scrollbar = ctk.CTkScrollbar(
            log_frame,
            command=self.log_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
    
    def _change_appearance_mode_event(self, new_appearance_mode: str):
        """切换外观模式"""
        ctk.set_appearance_mode(new_appearance_mode)
    
    def _change_color_theme_event(self, new_color_theme: str):
        """切换颜色主题"""
        ctk.set_default_color_theme(new_color_theme)
    
    def _browse_baseline_file(self):
        """浏览基准文件"""
        file_path = filedialog.askopenfilename(
            title="选择基准Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*")]
        )
        if file_path:
            self.baseline_entry.delete(0, ctk.END)
            self.baseline_entry.insert(0, file_path)
            self.baseline_file = file_path
    
    def _browse_compare_file(self):
        """浏览比较文件"""
        file_path = filedialog.askopenfilename(
            title="选择比较Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*")]
        )
        if file_path:
            self.compare_entry.delete(0, ctk.END)
            self.compare_entry.insert(0, file_path)
            self.compare_file = file_path
    
    def _browse_result_folder(self):
        """浏览结果文件夹"""
        folder_path = filedialog.askdirectory(
            title="选择结果保存文件夹"
        )
        if folder_path:
            self.result_folder_entry.delete(0, ctk.END)
            self.result_folder_entry.insert(0, folder_path)
            self.results_folder = folder_path
    
    def _start_compare(self):
        """开始比较"""
        # 检查文件是否选择
        self.baseline_file = self.baseline_entry.get().strip()
        self.compare_file = self.compare_entry.get().strip()
        
        if not self.baseline_file or not self.compare_file:
            messagebox.showerror("错误", "请选择基准文件和比较文件")
            return
        
        if not os.path.exists(self.baseline_file):
            messagebox.showerror("错误", f"基准文件不存在: {self.baseline_file}")
            return
        
        if not os.path.exists(self.compare_file):
            messagebox.showerror("错误", f"比较文件不存在: {self.compare_file}")
            return
        
        # 开始比较
        self.running = True
        self.stop_event.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        
        # 清空日志
        self.log_text.delete("1.0", ctk.END)
        
        # 创建工作线程
        self.worker_thread = threading.Thread(
            target=self._compare_worker,
            daemon=True
        )
        self.worker_thread.start()
    
    def _stop_compare(self):
        """停止比较"""
        self.stop_event.set()
        self.stop_button.configure(state="disabled")
    
    def _select_feature_columns(self):
        """选择特征列"""
        if not self.baseline_file:
            messagebox.showerror("错误", "请先选择基准文件")
            return
        
        try:
            # 加载基准文件获取表头信息
            wb = openpyxl.load_workbook(self.baseline_file, data_only=True)
            ws = wb.active
            
            # 获取用户输入的表头行号
            try:
                header_row = int(self.header_row_var.get())
            except ValueError:
                messagebox.showerror("错误", "表头行号必须是数字")
                return
            
            # 获取表头行的列名
            max_col = ws.max_column
            header_values = []
            for col in range(1, max_col + 1):
                cell_value = ws.cell(row=header_row, column=col).value
                if cell_value:
                    header_values.append(f"{col}: {cell_value.strip()}")
                else:
                    header_values.append(f"{col}: 空")
            
            # 创建特征列选择窗口
            select_window = ctk.CTkToplevel(self)
            select_window.title("选择特征列")
            select_window.geometry("400x300")
            select_window.resizable(False, False)
            
            # 居中显示
            select_window.transient(self)
            select_window.grab_set()
            
            # 创建列表框
            listbox = ctk.CTkScrollableFrame(select_window)
            listbox.pack(fill="both", expand=True, padx=10, pady=10)
            
            # 存储选中的列
            selected_cols = []
            
            # 创建复选框
            checkboxes = []
            for i, header in enumerate(header_values[:20]):  # 最多显示20列
                var = ctk.IntVar()
                checkbox = ctk.CTkCheckBox(listbox, text=header, variable=var)
                checkbox.pack(anchor="w", pady=5)
                checkboxes.append((var, i + 1))  # 列号从1开始
            
            # 选择按钮
            def on_select():
                selected = [col for var, col in checkboxes if var.get() == 1]
                if len(selected) == 0:
                    messagebox.showerror("错误", "请至少选择1列")
                    return
                if len(selected) > 6:
                    messagebox.showerror("错误", "最多只能选择6列")
                    return
                
                # 更新特征列显示
                self.feature_cols_var.set(", ".join(map(str, selected)))
                select_window.destroy()
            
            select_button = ctk.CTkButton(select_window, text="确定", command=on_select, fg_color="#4CAF50")
            select_button.pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
    
    def _compare_worker(self):
        """比较工作线程"""
        try:
            # 获取表头行号
            try:
                header_row = int(self.header_row_var.get())
            except ValueError:
                log_queue.put("\n❌ 错误：表头行号必须是数字")
                return False
            
            # 获取特征列
            feature_cols_str = self.feature_cols_var.get()
            key_fields = None
            if feature_cols_str != "前三列":
                try:
                    # 解析特征列，支持多种格式："1,2,3" 或 "1 2 3" 或 "1-3"
                    feature_cols = []
                    # 处理逗号分隔
                    parts = [p.strip() for p in feature_cols_str.split(",")]
                    for part in parts:
                        # 处理空格分隔
                        sub_parts = [sp.strip() for sp in part.split() if sp.strip()]
                        for sub_part in sub_parts:
                            # 处理范围
                            if "-" in sub_part:
                                start, end = map(int, sub_part.split("-"))
                                feature_cols.extend(range(start, end + 1))
                            else:
                                feature_cols.append(int(sub_part))
                    # 去重并排序
                    feature_cols = sorted(list(set(feature_cols)))
                    # 转换为列名格式
                    key_fields = [f"列{col}" for col in feature_cols]
                except ValueError:
                    log_queue.put("\n❌ 错误：特征列格式无效")
                    return False
            
            # 生成结果文件名
            baseline_folder = os.path.basename(os.path.dirname(self.baseline_file))
            compare_folder = os.path.basename(os.path.dirname(self.compare_file))
            original_filename = os.path.basename(self.baseline_file).replace('.xlsx', '')
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 构建结果文件路径
            result_baseline = os.path.join(
                self.results_folder, 
                f"{original_filename}_{baseline_folder}_比较结果_{timestamp}.xlsx"
            )
            result_compare = os.path.join(
                self.results_folder, 
                f"{original_filename}_{compare_folder}_比较结果_{timestamp}.xlsx"
            )
            
            # 调用比较函数
            success = compare_excel_files(
                self.baseline_file, 
                self.compare_file, 
                result_baseline, 
                result_compare,
                self.results_folder,
                original_filename,
                timestamp,
                header_row,
                key_fields,
                self.stop_event
            )
            
            if success:
                log_queue.put("\n✅ 比较完成！")
            else:
                log_queue.put("\n❌ 比较失败！")
        except Exception as e:
            log_queue.put(f"\n❌ 比较过程中出错: {str(e)}")
        finally:
            # 更新UI状态
            self.running = False
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
    
    def _redirect_stdout(self):
        """重定向标准输出到日志组件"""
        sys.stdout = StdoutRedirector(self.log_text)
    
    def _listen_queues(self):
        """监听日志队列并更新UI"""
        try:
            while not log_queue.empty():
                message = log_queue.get_nowait()
                self.log_text.insert(ctk.END, message)
                self.log_text.see(ctk.END)
        except queue.Empty:
            pass
        finally:
            # 每100ms检查一次队列
            self.after(100, self._listen_queues)

if __name__ == "__main__":
    app = ExcelCompareGUI()
    app.mainloop()