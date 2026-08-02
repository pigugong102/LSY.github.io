import sys

# Fix WBS
with open('docs/planslot/WBS.html', 'r', encoding='utf-8') as f:
    lines = f.read().splitlines()

new_lines = []
for line in lines:
    if line.startswith('<col ') and line == lines[23]: # The last col in colgroup is at index 23 usually. Wait, better to parse.
        pass

# Let's use regex or string manipulation
import re
with open('docs/planslot/WBS.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Remove last 4 rows: <tr>...</tr> 4 times before </table>
html = re.sub(r'(<tr>\s*(<td[^>]*>.*?</td>\s*)+\s*</tr>\s*){4}</table>', '</table>', html)

# Remove last <col>
html = re.sub(r'<col[^>]*>\s*</colgroup>', '</colgroup>', html)

# Remove last <td> in every row
html = re.sub(r'<td[^>]*>(?:<b>)?\s*(?:</b>)?</td>\s*</tr>', '</tr>', html)

with open('docs/planslot/WBS.html', 'w', encoding='utf-8') as f:
    f.write(html)

# Fix Testcase '비고' width
with open('docs/planslot/테스트케이스.html', 'r', encoding='utf-8') as f:
    tc_html = f.read()

# Let's find the colgroup and modify the last col width which is usually '비고'. 
# In testcase, let's just make all <col>s wider or specifically the last one.
# Wait, let's check Testcase columns first.
