# coding=utf-8
import os
import openpyxl
from openpyxl.utils import get_column_letter

def rgb_to_hex(rgb_str):
    if not rgb_str or str(rgb_str) == '00000000' or str(rgb_str) == 'Values must be of type <class ''str''>': return None
    rgb_str = str(rgb_str)
    if len(rgb_str) == 8:
        return '#' + rgb_str[2:]
    elif len(rgb_str) == 6:
        return '#' + rgb_str
    return None

def excel_to_html(excel_path, html_path, file_name, project):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    html_content = '''<html><head><meta charset="utf-8">
    <style>
    body { font-family:-apple-system, BlinkMacSystemFont, sans-serif; padding:24px; margin:0; background:#f9f9fa; }
    h2 { color:#111; margin:32px 0 16px; font-size:20px; border-bottom:2px solid #222; padding-bottom:8px; display:inline-block; }
    table { table-layout:auto; border-collapse:collapse; width:100%; max-width:1400px; font-size:13.5px; text-align:left; background:#fff; margin-bottom:48px; box-shadow:0 2px 8px rgba(0,0,0,0.04); border-radius:4px; overflow:hidden; }
    th, td { border:1px solid #e0e0e0; padding:10px 14px; word-break:keep-all; vertical-align:middle; line-height:1.6; color:#333; }
    tr:nth-child(even) { background:#fafafa; }
    </style></head><body>'''
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        actual_max_row = 1
        actual_max_col = 1
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None or (cell.fill and cell.fill.start_color and getattr(cell.fill.start_color, 'rgb', None) and str(cell.fill.start_color.rgb) not in ['00000000', 'None']):
                    if cell.row > actual_max_row: actual_max_row = cell.row
                    if cell.column > actual_max_col: actual_max_col = cell.column
        
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            if max_row > actual_max_row: actual_max_row = max_row
            if max_col > actual_max_col: actual_max_col = max_col

        if project == 'planslot' and file_name == 'WBS.xlsx':
            actual_max_row -= 4
            actual_max_col -= 1

        if actual_max_row == 1 and actual_max_col == 1 and ws.cell(1,1).value is None:
            continue
            
        html_content += f'<h2>{sheet_name}</h2>\n'
        html_content += '<table>\n'
        
        html_content += '<colgroup>\n'
        for c in range(1, actual_max_col + 1):
            col_letter = get_column_letter(c)
            width = ws.column_dimensions[col_letter].width
            
            if project == 'planslot' and file_name == '테스트케이스.xlsx' and c == actual_max_col:
                width = 50
                
            if width:
                html_content += f'<col style="width: {int(width * 8)}px;">\n'
            else:
                html_content += '<col>\n'
        html_content += '</colgroup>\n'
        
        merged_cells_map = {}
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            merged_cells_map[(min_row, min_col)] = {
                'rowspan': max_row - min_row + 1,
                'colspan': max_col - min_col + 1
            }
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if (r, c) != (min_row, min_col):
                        merged_cells_map[(r, c)] = 'skip'
                        
        for row in range(1, actual_max_row + 1):
            html_content += '<tr>\n'
            for col in range(1, actual_max_col + 1):
                cell = ws.cell(row, col)
                merge_info = merged_cells_map.get((row, col))
                if merge_info == 'skip':
                    continue
                
                rowspan = merge_info['rowspan'] if merge_info else 1
                colspan = merge_info['colspan'] if merge_info else 1
                
                bg_color = None
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                    bg_color = rgb_to_hex(cell.fill.start_color.rgb)
                    
                style_attr = ''
                if bg_color and bg_color not in ['#000000', '#FFFFFF', '#ffffff']:
                    style_attr = f' style="background-color: {bg_color}; font-weight:600;"'
                    
                attrs = f' rowspan="{rowspan}" colspan="{colspan}"{style_attr}'
                
                val = cell.value
                if val is None:
                    val = ''
                else:
                    val = str(val).replace('\n', '<br>')
                    
                if cell.font and cell.font.bold:
                    val = f'<b>{val}</b>'
                    
                html_content += f'<td{attrs}>{val}</td>\n'
            html_content += '</tr>\n'
            
        html_content += '</table>\n'
        
    html_content += '''
    <script>
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        if (window.parent && window.parent.closeDocModal) {
          window.parent.closeDocModal();
        }
      }
    });
    </script>
    </body></html>'''
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

for project in ['planslot', 'aimatch']:
    docs_path = f'docs/{project}'
    if not os.path.exists(docs_path): continue
    for file in os.listdir(docs_path):
        if file.endswith('.xlsx'):
            try:
                excel_to_html(os.path.join(docs_path, file), os.path.join(docs_path, file.replace('.xlsx', '.html')), file, project)
                print(f'Converted {file}')
            except Exception as e:
                print(f'Error converting {file}: {e}')
