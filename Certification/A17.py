import os
import re
import shutil
import time
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
import pdfplumber

# ==================== 尝试导入 rarfile ====================
try:
    import rarfile
    RAR_SUPPORT = True
    if os.name == 'nt':
        unrar_candidates = [
            r"C:\Program Files\WinRAR\UnRAR.exe",
            r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
        ]
        for cand in unrar_candidates:
            if os.path.exists(cand):
                rarfile.UNRAR_TOOL = cand
                break
except ImportError:
    RAR_SUPPORT = False
    rarfile = None

# ==================== 预设牌号列表 ====================
GRADE_LIST = [
    "HC420/780DPD+Z", "HC340/590DP", "HC420/780DP", "DX53D+AS100", "DC54D+AS80",
    "HC300LAD+Z", "HC260LAD+Z", "HC340LAD+Z", "CDCM-SPCC", "BR440/590HE",
    "QSTE500TM-P", "SAPH440-P", "SPFH590-P", "SAPH400-P", "S355MC-P",
    "S550MC-P", "M610L-P", "M420L-P", "QSTE460TM", "QSTE420TM", "QSTE380TM",
    "QStE550TM", "SPHC-MJ", "SPHC-SF", "SPHC-SD", "SPHC-Z", "SPHC-B", "SPHC-L",
    "SPHC-T", "SPHC", "SAPH440", "SAPH400", "SAPH370", "S355MC", "S420MC",
    "S500MC", "S550MC", "S315MC", "SPFH540", "SPFH590", "Grade 80", "IF GRADE",
    "IF-FT", "B35A270-K", "B35A270-A", "B50A350-A", "B50A400-A", "B50AH470-A",
    "B1500HS", "BR1500HS", "HR800CP", "HR420LA", "M25V1300-H", "M30HV1500",
    "WDER600", "DCR650", "BS600MC", "B750L", "B610L", "B510L", "B420L", "B410LA",
    "B280VK", "BFT300", "BFT350", "BJC355", "BLD", "BAC300", "BTC1", "CP800",
    "DCL250", "DCL350", "DC06", "DC01", "DC03", "DC04", "DC54D+Z", "DD13",
    "DD11", "SPCC-XM", "SPCC-4D", "SPCC", "SPCD", "SPHE", "SPHD", "SPHF-F1",
    "SPHF-F", "SAE 1010", "SS400", "ST37-2G", "Q235B", "Q355B", "Q390D",
    "Q345D", "Q345B", "Q195", "25XW1300", "420L", "610L", "S510L", "S610L",
    "S420L", "HC420LA", "HC340LA", "HC460LA", "HC380LA", "HC260LA", "HC500LA",
    "F11", "F18", "CDCM SPCC", "Grade80", "SAE1010", "QStE500TM", "BHG2", "S550MC-P", "QStE380TM", "St37-2G", "HC260LAD+Z", "Q460C"
]


def extract_pdf_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text


def extract_contract_no(pdf_path, filename=None):
    keywords = ['合同', 'CONTRACT', "MILL"]
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            
            for i, word in enumerate(words):
                matched = False
                for lookahead in range(3):
                    if i + lookahead >= len(words):
                        break
                    combined = ''.join(words[i+j]['text'] for j in range(lookahead+1))
                    combined_clean = combined.upper().replace(' ', '').replace('’', "'").replace('`', "'")
                    if any(kw.upper().replace(' ', '') in combined_clean for kw in keywords):
                        matched = True
                        break
                
                if not matched:
                    continue
                
                kx0 = word['x0']
                ktop = word['top']
                kbottom = word['bottom']
                
                for next_word in words[i+1:]:
                    ntext = next_word['text'].strip()
                    if len(ntext) < 10:
                        continue
                    
                    nx0 = next_word['x0']
                    ntop = next_word['top']
                    nbottom = next_word['bottom']
                    
                    is_right = (nx0 > kx0 - 20 and abs(ntop - ktop) < 25)
                    is_below = (ntop > kbottom - 5 and 
                               ntop - kbottom < 60 and 
                               abs(nx0 - kx0) < 250)
                    
                    if (is_right or is_below) and re.match(r'^[A-Z][A-Z0-9]{7,14}$', ntext) and any(c.isdigit() for c in ntext) and any(c.isalpha() for c in ntext) and not ntext.startswith(('BG', 'BA', 'BT', 'IC', 'CT', 'G0', 'LX')):
                        return ntext, "content"
    
    text = extract_pdf_text(pdf_path)
    
    matches = re.findall(r'(?<![A-Za-z0-9])[A-Z][A-Z0-9]{7,14}(?![A-Za-z0-9])', text)
    for m in matches:
        if any(c.isdigit() for c in m) and any(c.isalpha() for c in m):
            if not m.startswith(('BG', 'BA', 'BT', 'IC', 'CT', 'G0', 'LX')):
                return m, "content"
            
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    for i, line in enumerate(lines):
        if 'CONTRACT' in line.upper() or '合同' in line:
            for j in range(max(0, i-3), min(len(lines), i+15)):
                candidate = lines[j]
                if re.match(r'^[A-Z][A-Z0-9]{7,14}$', candidate):
                    return candidate, "content"
    
    if filename:
        base = Path(filename).stem
        m = re.match(r'^([A-Z][A-Z0-9]{7,14})_', base)
        if m and any(c.isdigit() for c in m.group(1)) and any(c.isalpha() for c in m.group(1)):
            return m.group(1), "filename"
            
    return None, None


def extract_grade(text):
    clean_text = ' '.join(text.split())
    no_space_text = clean_text.replace(' ', '')
    for grade in sorted(GRADE_LIST, key=len, reverse=True):
        if grade in clean_text:
            return grade
        no_space_grade = grade.replace(' ', '')
        if no_space_grade in no_space_text:
            return grade
    return None


def extract_entries(text):
    entries = []
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    # ===== 模式1：标准表格行（含 COIL / SHEETS，兼容中间多余字段）=====
    for line in lines:
        m = re.match(
            r"^\s*\d+\s+([A-Z0-9]{10,})\s+\d+\s+\d{1,8}\s*(\d+\.\d+)\s+(\d{3,4})\s+(?:COIL|SHEETS?)(?:\s+\d+)?\s+(\d{4,})",
            line
        )
        if m:
            entries.append({
                "coil_no": m.group(1),
                "thick": m.group(2),
                "width": m.group(3),
                "weight_kg": int(m.group(4))
            })
            
    # ===== 模式2：冷轧钢板横向一行（无 COIL，有长度+张数+重量）=====
    if not entries:
        for line in lines:
            m = re.match(
                r"^\s*\d+\s+([A-Z0-9]{10,})\s+\d+\s+\d+\s+(\d+\.\d+)\s+(\d{3,4})\s+\d+\s+\d+\s+(\d{4,})\b",
                line
            )
            if m:
                entries.append({
                    "coil_no": m.group(1),
                    "thick": m.group(2),
                    "width": m.group(3),
                    "weight_kg": int(m.group(4))
                })
    
    # ===== 模式3：梅钢热轧钢带散落排版（增强版）=====
    if not entries:
        # 预处理：把 "炉号 厚度" 这种组合行拆成独立字段，如 "36259707 10.20"
        processed_lines = []
        for line in lines:
            if re.match(r'^\d+\s+\d+\.\d+$', line):
                processed_lines.extend(line.split())
            else:
                processed_lines.append(line)
        
        # 先尝试查找厚度*宽度*COIL格式（保留原有兼容）
        thick = width = None
        for line in lines:
            m = re.search(r'(\d+\.\d+)\s*[\*×]\s*(\d{3,4})\s*[\*×]\s*(?:coil|COIL)', line)
            if m:
                thick = m.group(1)
                width = m.group(2)
                break
        
        # 如果没找到 *× 格式，从散落行中提取厚度和宽度
        if not thick or not width:
            thick_candidates = [l for l in processed_lines if re.match(r'^\d+\.\d{2}$', l)]
            width_candidates = [l for l in processed_lines if re.match(r'^\d{3,4}$', l) and 1000 <= int(l) <= 2500]
            if thick_candidates:
                thick = thick_candidates[0]
            if width_candidates:
                width_counts = Counter(width_candidates)
                width = width_counts.most_common(1)[0][0]
        
        # 提取钢卷号（10位以上）和重量（通常 >=5000kg）
        if thick and width:
            coil_nos = [l for l in processed_lines if re.match(r'^\d{10,}$', l)]
            weight_candidates = [int(l) for l in processed_lines if re.match(r'^\d{4,5}$', l) and int(l) >= 5000]
            
            if coil_nos:
                weight = Counter(weight_candidates).most_common(1)[0][0] if weight_candidates else 0
                for coil in coil_nos:
                    entries.append({
                        "coil_no": coil,
                        "thick": thick,
                        "width": width,
                        "weight_kg": weight
                    })
            else:
                # 回退到原有按行扫描逻辑
                for line in lines:
                    numbers = re.findall(r'\b\d+\b', line)
                    coil_no = None
                    weight = None
                    for num in numbers:
                        if len(num) >= 10 and not coil_no:
                            coil_no = num
                        elif 4 <= len(num) <= 5 and int(num) != int(width) and not weight:
                            weight = int(num)
                    if coil_no and weight:
                        entries.append({
                            "coil_no": coil_no,
                            "thick": thick,
                            "width": width,
                            "weight_kg": weight
                        })
    
    # ===== 模式4：竖排冷轧钢板老格式 =====
    if not entries:
        coil_lines = []
        for i, line in enumerate(lines):
            if re.match(r'^\d{10,}$', line):
                coil_lines.append((i, line))
        
        if len(coil_lines) >= 1 and 'SHEETS' in text:
            thick = None
            for line in lines:
                m = re.search(r'(\d+\.\d+)\s*mm', line)
                if m:
                    thick = m.group(1)
                    break
            
            width = None
            width_candidates = []
            for line in lines:
                m = re.match(r'^(\d{3,4})$', line)
                if m:
                    width_candidates.append(m.group(1))
            if width_candidates:
                width = Counter(width_candidates).most_common(1)[0][0]
            
            sheets_idx = -1
            for i, line in enumerate(lines):
                if 'SHEETS' in line:
                    sheets_idx = i
                    break
            
            if sheets_idx >= 0:
                weight_candidates = []
                for j in range(sheets_idx + 1, min(len(lines), sheets_idx + 25)):
                    m = re.match(r'^(\d{4,5})$', lines[j])
                    if m:
                        weight_candidates.append(int(m.group(1)))
                
                valid_weights = [w for w in weight_candidates if w >= 1000]
                if valid_weights:
                    weight_val = Counter(valid_weights).most_common(1)[0][0]
                    for _, coil in coil_lines:
                        entries.append({
                            "coil_no": coil,
                            "thick": thick or "0.00",
                            "width": width or "0",
                            "weight_kg": weight_val
                        })
    # ===== 模式5：新钢/散落格式（无COIL/SHEETS）=====
    if not entries:
        for line in lines:
            m = re.match(
                r"^\s*\d+\s+([A-Z0-9]{10,})\s+[A-Z0-9]{1,12}\s+(\d+\.\d+)\s+(\d{3,4})\s+\d+\s+(\d+\.\d+)",
                line
            )
            if m:
                width_val = int(m.group(3))
                weight_raw = float(m.group(4))
                
                # ↓↓↓ 单位自动判断 ↓↓↓
                if weight_raw < 100:
                    weight_kg = int(weight_raw * 1000)   # 吨 → kg
                else:
                    weight_kg = int(weight_raw)          # kg → 保持
                
                entries.append({
                    "coil_no": m.group(1),
                    "thick": m.group(2),
                    "width": m.group(3),
                    "weight_kg": weight_kg
                })
    # ===== 模式6：本钢浦项/士禾格式（数据分散在多行）=====
    if not entries:
        i = 0
        while i < len(lines):
            line = lines[i]
            # 钢卷号行：如 "267W250920100 9490 400 40 24 ..."
            m_coil = re.match(r'^([A-Z0-9]{10,})\s+(\d{4,})', line)
            if m_coil:
                coil_no = m_coil.group(1)
                weight = int(m_coil.group(2))
                # 看下一行是否是数据行：如 "1 1.5 1405 C 1 0 285 ..."
                if i + 1 < len(lines):
                    m_data = re.match(r'^\d+\s+(\d+\.\d+)\s+(\d{3,4})\s+(?:COIL|SHEETS?|C)\b', lines[i + 1])
                    if m_data:
                        thick = m_data.group(1)
                        width = m_data.group(2)
                        entries.append({
                            "coil_no": coil_no,
                            "thick": thick,
                            "width": width,
                            "weight_kg": weight
                        })
                        i += 2
                        continue
            i += 1

    # ===== 模式7：联鑫格式（厚度*宽度粘连，重量为吨）=====
    if not entries:
        for line in lines:
            # 格式：序号 钢卷号 包装 厚度*宽度 重量(吨) C Si Mn P S AlS 屈服 抗拉 伸长率
            # 例：1 ZX2B96601410 普包 3.0000*1500 13.180 40 10 25 8 13 239 336 51
            m = re.match(
                r"^\s*\d+\s+([A-Z0-9]{10,})\s+\S+\s+(\d+\.\d+)\*(\d{3,4})\s+(\d+\.\d+)",
                line
            )
            if m:
                width_val = int(m.group(3))
                weight_raw = float(m.group(4))
                if 100 <= width_val <= 2500:
                    if weight_raw < 100:
                        weight_kg = int(weight_raw * 1000)
                    else:
                        weight_kg = int(weight_raw)
                    entries.append({
                        "coil_no": m.group(1),
                        "thick": m.group(2),
                        "width": m.group(3),
                        "weight_kg": weight_kg
                    })

    return entries


def split_pdf(pdf_path, temp_dir):
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            print("  ⚠️ 缺少PDF拆分库，请先安装: pip install pypdf")
            return None
    # 新增：防止损坏PDF导致程序崩溃
    try:
        reader = PdfReader(str(pdf_path))
        page_count = len(reader.pages)
    except Exception as e:
        print(f"  ⚠️ PDF文件损坏或无法读取，跳过拆分: {Path(pdf_path).name} - {e}")
        return None
    
    if page_count <= 1:
        return [Path(pdf_path)]
    
    temp_dir = Path(temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    output_paths = []
    stem = Path(pdf_path).stem
    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        out_path = temp_dir / f"{stem}_page{i+1}.pdf"
        with open(out_path, 'wb') as f:
            writer.write(f)
        output_paths.append(out_path)
    
    return output_paths


def extract_archive(archive_path, extract_to):
    extract_to = Path(extract_to)
    extract_to.mkdir(parents=True, exist_ok=True)
    suffix = archive_path.suffix.lower()
    extracted_pdfs = []
    
    if suffix == '.zip':
        try:
            with zipfile.ZipFile(archive_path, 'r') as zf:
                for member in zf.namelist():
                    if member.lower().endswith('.pdf'):
                        zf.extract(member, extract_to)
            extracted_pdfs = sorted(extract_to.rglob("*.pdf"))
        except Exception as e:
            print(f"  ⚠️ 解压ZIP失败: {archive_path.name} - {e}")
    
    elif suffix == '.rar':
        if not RAR_SUPPORT:
            print(f"  ⚠️ 跳过RAR（未安装rarfile库）: {archive_path.name}")
        else:
            try:
                with rarfile.RarFile(archive_path, 'r') as rf:
                    for member in rf.namelist():
                        if member.lower().endswith('.pdf'):
                            rf.extract(member, extract_to)
                extracted_pdfs = sorted(extract_to.rglob("*.pdf"))
            except Exception as e:
                print(f"  ⚠️ 解压RAR失败（请确保已安装WinRAR）: {archive_path.name} - {e}")
    
    return extracted_pdfs


def process_one(pdf_path, output_dir, fresh_dir=None):
    text = extract_pdf_text(pdf_path)
    filename = Path(pdf_path).name
    
    contract_no, contract_source = extract_contract_no(pdf_path, filename)
    grade = extract_grade(text)
    entries = extract_entries(text)

    if not contract_no:
        print(f"    ⚠️ 未提取到合同号，跳过: {filename}")
        return "failed", 0, 0, []
    if not grade:
        print(f"    ⚠️ 未提取到牌号，跳过: {filename}")
        return "failed", 0, 0, []
    if not entries:
        print(f"    ⚠️ 未提取到钢卷数据，跳过: {filename}")
        return "failed", 0, 0, []

    if contract_source == "filename":
        print(f"    ℹ️ 合同号从文件名提取: {contract_no}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    if fresh_dir:
        fresh_dir = Path(fresh_dir)
        fresh_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    skip_count = 0
    generated_files = []

    for entry in entries:
        weight_ton = entry["weight_kg"] / 1000
        new_name = (
            f"{contract_no} "
            f"{entry['coil_no']} "
            f"{grade} "
            f"{entry['thick']} "
            f"{entry['width']} "
            f"{weight_ton:.3f}吨.pdf"
        )
        new_name = re.sub(r'[\\/:*?"<>|]', "", new_name)
        
        archive_path = output_dir / new_name
        fresh_path = fresh_dir / new_name if fresh_dir else None
        
        if archive_path.exists():
            print(f"    ⏭️ 同名文件已存在，跳过: {new_name}")
            skip_count += 1
            continue
        
        shutil.copy2(pdf_path, archive_path)
        print(f"    ✅ [存档] {archive_path.name}")
        success_count += 1
        generated_files.append(str(archive_path.name))
        
        if fresh_path:
            shutil.copy2(pdf_path, fresh_path)
            print(f"    🆕 [新鲜] {fresh_path.name}")
    
    if success_count > 0:
        return "success", success_count, skip_count, generated_files
    elif skip_count > 0:
        return "skipped", 0, skip_count, []
    else:
        return "failed", 0, 0, []


def generate_report(report_dir, start_time, end_time, success_files, fail_files, skip_files,
                    fresh_files, problem_files):
    report_dir = Path(report_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = report_dir / f"报告_{timestamp}.txt"
    
    duration = end_time - start_time
    hours = int(duration // 3600)
    minutes = int((duration % 3600) // 60)
    seconds = duration % 60
    
    lines = []
    lines.append("=" * 60)
    lines.append("           质保书批量处理报告")
    lines.append("=" * 60)
    lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"处理耗时: {hours}小时 {minutes}分钟 {seconds:.2f}秒")
    lines.append("")
    lines.append("-" * 60)
    lines.append("【统计汇总】")
    lines.append("-" * 60)
    lines.append(f"  ✅ 成功生成: {len(success_files)} 个文件")
    lines.append(f"  ⏭️ 同名跳过: {len(skip_files)} 个文件")
    lines.append(f"  ⚠️  识别失败: {len(fail_files)} 个文件")
    lines.append(f"  📄 总计处理: {len(success_files) + len(skip_files) + len(fail_files)} 个文件")
    lines.append("")
    lines.append("-" * 60)
    lines.append("【文件归档】")
    lines.append("-" * 60)
    lines.append(f"  🆕 成功识别（新鲜！）: {len(fresh_files)} 个文件")
    lines.append(f"  ❓ 未识别（疑难杂症）: {len(problem_files)} 个原文件")
    lines.append("")
    
    if success_files:
        lines.append("-" * 60)
        lines.append("【成功生成的文件】")
        lines.append("-" * 60)
        for f in success_files:
            lines.append(f"  ✅ {f}")
        lines.append("")
    
    if skip_files:
        lines.append("-" * 60)
        lines.append("【因同名跳过的文件】")
        lines.append("-" * 60)
        for f in skip_files:
            lines.append(f"  ⏭️ {f}")
        lines.append("")
    
    if fail_files:
        lines.append("-" * 60)
        lines.append("【识别失败的页/文件】")
        lines.append("-" * 60)
        for f in fail_files:
            lines.append(f"  ⚠️ {f}")
        lines.append("")
    
    if fresh_files:
        lines.append("-" * 60)
        lines.append("【成功识别的文件（已复制到 新鲜！）】")
        lines.append("-" * 60)
        for f in fresh_files:
            lines.append(f"  🆕 {f}")
        lines.append("")
    
    if problem_files:
        lines.append("-" * 60)
        lines.append("【未识别的原文件（已复制到 疑难杂症）】")
        lines.append("-" * 60)
        for f in problem_files:
            lines.append(f"  ❓ {f}")
        lines.append("")
    
    lines.append("=" * 60)
    lines.append("报告生成完毕")
    lines.append("=" * 60)
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    return str(report_path)


def batch_process(input_dir, output_dir):
    input_folder = Path(input_dir)
    if not input_folder.exists():
        print(f"❌ 输入文件夹不存在: {input_dir}")
        return

    base_dir = Path(output_dir).parent
    temp_dir = base_dir / ".temp_split"
    archive_temp = base_dir / ".temp_archive"
    report_dir = base_dir / "报告"
    fresh_dir = base_dir / "质保书（新鲜！）"
    problem_dir = base_dir / "质保书（疑难杂症）"
    
    fresh_dir.mkdir(parents=True, exist_ok=True)
    problem_dir.mkdir(parents=True, exist_ok=True)
    
    all_pdf_tasks = []
    
    direct_pdfs = sorted(input_folder.glob("*.pdf"))
    for pdf in direct_pdfs:
        all_pdf_tasks.append((pdf, pdf.name, False, None))
    
    archives = sorted(input_folder.glob("*.zip")) + sorted(input_folder.glob("*.rar"))
    for archive in archives:
        print(f"📦 正在解压: {archive.name}")
        extracted = extract_archive(archive, archive_temp / archive.stem)
        for pdf in extracted:
            all_pdf_tasks.append((pdf, f"{archive.name}/{pdf.name}", True, archive.name))
    
    if not all_pdf_tasks:
        print(f"⚠️ 未找到 PDF 文件或压缩包: {input_dir}")
        return

    print(f"📁 输入: {input_dir}")
    print(f"📁 输出(存档): {output_dir}")
    print(f"📁 报告: {report_dir}")
    print(f"🆕 新鲜: {fresh_dir}")
    print(f"❓ 疑难杂症: {problem_dir}")
    print(f"📄 直接PDF: {len(direct_pdfs)} 个 | 压缩包: {len(archives)} 个 | 待处理PDF: {len(all_pdf_tasks)} 个\n")

    start_time = time.time()
    
    total_ok = 0
    total_skip = 0
    total_fail = 0
    
    success_files = []
    fail_files = []
    skip_files = []
    
    fresh_files = []
    problem_files = []

    for pdf_path, display_name, is_from_archive, archive_name in all_pdf_tasks:
        page_files = split_pdf(pdf_path, temp_dir)
        if page_files is None:
            total_fail += 1
            fail_files.append(display_name)
            shutil.copy2(pdf_path, problem_dir)
            problem_files.append(Path(pdf_path).name)
            continue
        
        source_label = f"📦 {archive_name} → " if is_from_archive else ""
        if len(page_files) > 1:
            print(f"{source_label}🔍 {display_name}（共 {len(page_files)} 页，已拆分）")
        else:
            print(f"{source_label}🔍 {display_name}")
        
        pdf_has_success = False
        pdf_has_fail = False
        
        for page_file in page_files:
            status, s_ok, s_skip, gen_files = process_one(page_file, output_dir, fresh_dir)
            if status == "success":
                total_ok += s_ok
                total_skip += s_skip
                success_files.extend(gen_files)
                fresh_files.extend(gen_files)
                pdf_has_success = True
            elif status == "skipped":
                total_skip += s_skip
                for f in gen_files:
                    skip_files.append(f)
            else:
                total_fail += 1
                fail_files.append(page_file.name)
                pdf_has_fail = True
        
        if not pdf_has_success and pdf_has_fail:
            problem_dest = problem_dir / Path(pdf_path).name
            if not problem_dest.exists():
                try:
                    shutil.copy2(pdf_path, problem_dir)
                    problem_files.append(Path(pdf_path).name)
                except PermissionError:
                    print(f"    ⚠️ 无法复制到疑难杂症（权限/已占用）: {Path(pdf_path).name}")
            else:
                print(f"    ⏭️ 疑难杂症中已存在，跳过复制: {Path(pdf_path).name}")
        
        if len(page_files) > 1:
            for pf in page_files:
                if pf.exists():
                    try:
                        pf.unlink()
                    except Exception:
                        pass
    
    for td in [temp_dir, archive_temp]:
        if td.exists():
            try:
                shutil.rmtree(str(td), ignore_errors=True)
            except Exception:
                pass

    end_time = time.time()
    duration = end_time - start_time
    hours = int(duration // 3600)
    minutes = int((duration % 3600) // 60)
    seconds = duration % 60

    report_path = generate_report(report_dir, start_time, end_time,
                                  success_files, fail_files, skip_files,
                                  fresh_files, problem_files)

    print(f"\n{'='*50}")
    print(f"📊 处理完成")
    print(f"   ✅ 成功生成: {total_ok} 个文件")
    print(f"   ⏭️ 同名跳过: {total_skip} 个文件")
    print(f"   ⚠️  识别失败: {total_fail} 个文件")
    print(f"   🆕 新鲜！文件: {len(fresh_files)} 个")
    print(f"   ❓ 疑难杂症原文件: {len(problem_files)} 个")
    print(f"   ⏱️  处理耗时: {hours}小时 {minutes}分钟 {seconds:.2f}秒")
    print(f"   📝 报告位置: {report_path}")
    print(f"{'='*50}")


if __name__ == "__main__":
    INPUT_FOLDER = r"D:\OneDrive\质保书改名\质保书（原料）"
    OUTPUT_FOLDER = r"D:\OneDrive\质保书改名\质保书（存档）"
    batch_process(INPUT_FOLDER, OUTPUT_FOLDER)

    #额外注意
    #把模式1正则里的钢卷号匹配从 \d{10,} 改成 [A-Z0-9]{10,}，允许字母
    #这份质保书有 2页，第2页只有钢卷号和非金属夹杂物数据，没有重量/厚度/宽度。代码拆分PDF后第2页会识别失败，但只要第1页能成功识别，整份质保书最终还是会正常归档（不会进"疑难杂症"）。