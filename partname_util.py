# partname_util.py
import os
import re
import time
import json

default_parts_path = os.path.normpath("K:\\PARTS")
exclude_dirs = [os.path.join(default_parts_path, "S")]

def extract_description_from_stp(file_path):
    """从 .stp 文件中提取描述信息"""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                match = re.search(r'/\*\s*description\s*\*/\s*\(\s*[\'"](.+?)[\'"]\s*\)', line, re.IGNORECASE)
                if match:
                    return match.group(1)
    except Exception:
        pass
    return ""

def extract_revision(filename):
    """提取版本信息"""
    match = re.search(r'(R+)(\d+)\.stp$', filename, re.IGNORECASE)
    return match.group(1) + match.group(2) if match else ""

def rev_to_int(rev):
    """将版本信息转为整数进行比较"""
    match = re.match(r'R+(\d+)', rev, re.IGNORECASE)
    return int(match.group(1)) if match else -1

def find_highest_revision_file(files):
    """查找目录下的 PDF 文件，返回最新版本"""
    rev_map = {}
    for f in files:
        match = re.match(r"(.+?)R+(\d+)\.(pdf)$", f, re.IGNORECASE)
        if match:
            base, rev, ext = match.groups()
            rev_num = int(rev)
            if base not in rev_map or rev_num > rev_map[base][0]:
                rev_map[base] = (rev_num, f)
        else:
            base = f.rsplit('.', 1)[0]
            if base not in rev_map:
                rev_map[base] = (0, f)
    return [v[1] for v in rev_map.values()]

def generate_partname_dat(callback=None):
    """后台生成 partname.dat 并在完成后调用回调"""

    # 用于计算生成时间，调试完成后可删除
    start_time = time.time()
    result = {}
    for root, dirs, files in os.walk(default_parts_path):
        if any(ex in root for ex in exclude_dirs):
            continue  # 跳过排除目录
        stp_files = [f for f in files if f.lower().endswith('.stp')]
        for stp_file in stp_files:
            match = re.match(r'^(\d{2}[A-Z]\d{6})(R*\d*)\.stp$', stp_file, re.IGNORECASE)
            if not match:
                continue
            part_number = match.group(1)
            revision = extract_revision(stp_file)
            stp_path = os.path.join(root, stp_file)
            description = extract_description_from_stp(stp_path)
            entry = {
                "id": part_number,
                "rev": revision,
                "description": description,
                "stp": stp_path
            }

            # 查找PDF、IPT、IAM
            pdf_files = [f for f in files if f.lower().endswith('.pdf') and f.startswith(part_number)]
            highest_pdf = find_highest_revision_file(pdf_files)
            if highest_pdf:
                entry["pdf"] = os.path.join(root, highest_pdf[0])
            for ext in ("ipt", "iam"):
                candidate = f"{part_number}.{ext}"
                if candidate in files:
                    entry[ext] = os.path.join(root, candidate)

            if part_number in result:
                existing = result[part_number]
                if rev_to_int(revision) > rev_to_int(existing.get("rev", "")):
                    for key in ("pdf", "ipt", "iam"):
                        if key not in entry and key in existing:
                            entry[key] = existing[key]
                    result[part_number] = entry
                elif rev_to_int(revision) == rev_to_int(existing.get("rev", "")):
                    for key in ("pdf", "ipt", "iam"):
                        if key not in existing and key in entry:
                            existing[key] = entry[key]
                    existing["description"] = entry["description"]
                    existing["stp"] = entry["stp"]
                    existing["rev"] = revision
                    result[part_number] = existing
            else:
                result[part_number] = entry

    dat_path = os.path.join(os.path.dirname(__file__), "partname.dat")
    with open(dat_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    # 显示生成时间，调试完成后可删除
    elapsed = time.time() - start_time
    print(f"Partname.dat generated in {elapsed:.2f} seconds, {len(result)} items found.")

    if callback:
        callback()
