import re

fpath = 'generate_ch1_ch5_knowledge_doc.js'
with open(fpath, 'r', encoding='utf-8') as f:
    content = f.read()

result = []
i = 0
in_string = False
n = len(content)

def is_string_opener_context(content, pos):
    """Look back: is the immediately preceding non-whitespace one of: : , ( [ { = + return"""
    j = pos - 1
    while j >= 0 and content[j] in ' \t\n':
        j -= 1
    if j < 0:
        return False
    if content[j] in ':,([{=+':
        return True
    if j >= 5 and content[j-5:j+1] == 'return':
        return True
    return False

def is_string_closer_context(content, pos):
    """Look ahead: is the immediately following non-whitespace one of: , ) ] } + ; :"""
    j = pos + 1
    while j < len(content) and content[j] in ' \t\n':
        j += 1
    if j >= len(content):
        return True
    if content[j] in ',)]}+;:':
        return True
    return False

while i < n:
    ch = content[i]

    # Skip line comments
    if not in_string and ch == '/' and i+1 < n and content[i+1] == '/':
        while i < n and content[i] != '\n':
            result.append(content[i])
            i += 1
        continue

    # Escape sequence inside string
    if ch == chr(92) and in_string:
        result.append(ch)
        if i+1 < n:
            result.append(content[i+1])
            i += 2
        else:
            i += 1
        continue

    if ch == '"':
        if not in_string:
            if is_string_opener_context(content, i):
                in_string = True
                result.append(ch)
                i += 1
            else:
                # outside string but not opener context - keep as smart quote
                result.append(chr(0x201C))  # left double smart
                i += 1
        else:
            if is_string_closer_context(content, i):
                in_string = False
                result.append(ch)
                i += 1
            else:
                # inner - replace with smart quote
                prev_ch = content[i-1] if i > 0 else ''
                if prev_ch and (ord('一') <= ord(prev_ch) <= ord('鿿') or prev_ch.isalpha() or prev_ch in ')]} '):
                    result.append(chr(0x201D))  # right double smart
                else:
                    result.append(chr(0x201C))  # left double smart
                i += 1
    else:
        result.append(ch)
        i += 1

new_content = ''.join(result)
with open(fpath, 'w', encoding='utf-8') as f:
    f.write(new_content)
print(f"Fixed. New length: {len(new_content)}")
