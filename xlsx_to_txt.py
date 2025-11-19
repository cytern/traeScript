import argparse
import os
import sys
import glob

try:
    import openpyxl
except Exception as e:
    sys.stderr.write("需要安装 openpyxl，请运行: pip install openpyxl\n")
    sys.exit(1)

def find_input_xlsx(provided):
    if provided:
        return provided
    files = glob.glob("*.xlsx")
    if len(files) == 1:
        return files[0]
    if len(files) == 0:
        sys.stderr.write("当前目录未找到 xlsx 文件\n")
    else:
        sys.stderr.write("当前目录存在多个 xlsx 文件，请使用 --input 指定\n")
    sys.exit(1)

def normalize(s):
    if s is None:
        return ""
    return str(s).strip()

def parse_answer_labels(raw, option_count):
    if raw is None:
        return set()
    s = str(raw).upper()
    labels = set()
    for ch in s:
        if "A" <= ch <= "Z":
            labels.add(ch)
    if labels:
        return labels
    digits = []
    cur = ""
    for ch in s:
        if ch.isdigit():
            cur += ch
        else:
            if cur:
                digits.append(int(cur))
                cur = ""
    if cur:
        digits.append(int(cur))
    if digits:
        labels = set()
        for d in digits:
            i = d - 1
            if 0 <= i < option_count:
                labels.add(chr(ord("A") + i))
        return labels
    return set()

def detect_columns(headers):
    q_idx = None
    a_idx = None
    opt_map = {}
    for i, h in enumerate(headers):
        hs = normalize(h)
        hs_lower = hs.lower()
        if q_idx is None and ("题目" in hs or "问" in hs or "question" in hs_lower or "标题" in hs):
            q_idx = i
            continue
        if a_idx is None and ("正确" in hs or "答案" in hs or "answer" in hs_lower):
            a_idx = i
            continue
        for letter in ["A","B","C","D","E","F","G","H"]:
            if hs == letter or (letter in hs and ("选项" in hs or hs.startswith(letter))):
                opt_map[letter] = i
                break
    if q_idx is None:
        q_idx = None
    if a_idx is None:
        a_idx = len(headers) - 1
    if not opt_map:
        opt_map = {}
        letters = []
        for j in range(len(headers)):
            if j == q_idx or j == a_idx:
                continue
            letters.append(chr(ord("A") + len(letters)))
            opt_map[letters[-1]] = j
    if q_idx is None:
        candidate_idxs = [i for i in range(len(headers)) if i != a_idx and i not in opt_map.values()]
        if candidate_idxs:
            q_idx = candidate_idxs[0]
        else:
            q_idx = 0
    return q_idx, a_idx, opt_map

def main():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--input", dest="input", default=None)
    parser.add_argument("--output", dest="output", default="questions.txt")
    args = parser.parse_args()

    xlsx_path = find_input_xlsx(args.input)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.stderr.write("xlsx 内容为空\n")
        sys.exit(1)
    headers = rows[0]
    q_idx, ans_idx, opt_map = detect_columns(headers)
    if q_idx is None or True:
        candidate_idxs = [i for i in range(len(headers)) if i != ans_idx and i not in opt_map.values()]
        best_i = None
        best_len = -1.0
        for i in candidate_idxs:
            lengths = []
            for r in rows[1:]:
                if r is None or i >= len(r):
                    continue
                s = normalize(r[i])
                if not s:
                    continue
                lengths.append(len(s))
            avg_len = (sum(lengths) / len(lengths)) if lengths else 0.0
            if avg_len > best_len:
                best_len = avg_len
                best_i = i
        if best_i is not None:
            q_idx = best_i

    lines = []
    letters_order = sorted(opt_map.keys())
    for row in rows[1:]:
        if row is None:
            continue
        question = normalize(row[q_idx] if q_idx < len(row) else None).replace("\r", "").replace("\n", " ")
        if not question:
            continue
        options = {}
        for letter in letters_order:
            idx = opt_map[letter]
            text = normalize(row[idx] if idx < len(row) else None).replace("\r", "").replace("\n", " ")
            options[letter] = text
        raw_answer = row[ans_idx] if ans_idx < len(row) else None
        correct = parse_answer_labels(raw_answer, len(letters_order))
        existing_letters = [l for l in letters_order if options.get(l)]
        correct_list = []
        wrong_list = []
        for letter in existing_letters:
            item = f"[{letter}] {options[letter]}"
            if letter in correct:
                correct_list.append(item)
            else:
                wrong_list.append(item)
        line = f"{question}、正确答案:" + "、".join(correct_list) + "、错误答案:" + "、".join(wrong_list)
        lines.append(line)

    with open(args.output, "w", encoding="utf-8-sig") as f:
        for i, line in enumerate(lines):
            f.write(line)
            if i != len(lines) - 1:
                f.write("\n")

    print(f"已生成: {args.output}")

if __name__ == "__main__":
    main()
