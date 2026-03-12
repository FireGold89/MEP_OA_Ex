import sys
import re

def parse_patch(patch_path, target_file):
    with open(patch_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    in_file = False
    hunks = []
    current_hunk = None
    
    for line in lines:
        if line.startswith('+++ b/' + target_file):
            in_file = True
        elif line.startswith('+++ '):
            in_file = False
        elif in_file:
            if line.startswith('@@ '):
                m = re.match(r'@@ -(\d+),?(\d+)? \+(\d+),?(\d+)? @@', line)
                if not m: continue
                minus_start = int(m.group(1))
                minus_count = int(m.group(2)) if m.group(2) else 1
                current_hunk = {
                    'minus_start': minus_start,
                    'minus_count': minus_count,
                    'lines': []
                }
                hunks.append(current_hunk)
            elif current_hunk is not None:
                if line.startswith(('+', '-', ' ')):
                    current_hunk['lines'].append(line)
                    
    return hunks

def apply_hunks(orig_path, hunks):
    with open(orig_path, 'r', encoding='utf-8') as f:
        orig = f.read().splitlines()
        
    for hunk in reversed(hunks):
        delete_start = hunk['minus_start'] - 1
        delete_count = hunk['minus_count']
        
        new_lines = []
        for line in hunk['lines']:
            if line.startswith('+') or line.startswith(' '):
                # remove the leading character but preserve trailing spaces
                content_line = line[1:].rstrip('\r\n')
                new_lines.append(content_line)
                
        orig[delete_start:delete_start+delete_count] = new_lines
        
    with open(orig_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(orig) + '\n')

hunks_css = parse_patch('local_diff_utf8.patch', 'extension/content.css')
if hunks_css: apply_hunks('extension/content.css', hunks_css)

hunks_js = parse_patch('local_diff_utf8.patch', 'extension/content.js')
if hunks_js: apply_hunks('extension/content.js', hunks_js)

print('Patch applied successfully to content.css and content.js')
