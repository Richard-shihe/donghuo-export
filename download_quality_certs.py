"""
质保书 PDF 下载主脚本 v2
流程:
1. 登录 IEC
2. 调用 queryCertificate 获取质保书列表（HTML格式）
3. 解析每行数据，获取加密字段 (tcNumTcRsa, factoryOrderNumRsa)
4. 调用 decryptByCertificateInfo 解密（传加密值）
5. 构造 POST 请求到 ecommerce.ibaosteel.com 获取 PDF
"""
import os, sys, json, re, time, urllib.parse
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ibaosteel_client import IEC

BASE = "https://www.ibaosteel.com"

# 加载环境变量
username = os.environ.get('IBAO_USERNAME', '')
password = os.environ.get('IBAO_PASSWORD', '')
iec = IEC(username=username, password=password, retries=30)
iec.login()
s = iec.session
token = iec.token
print(f"✅ 登录成功, token={token[:20]}...")

# 初始化 session
ref = f"{BASE}/iecs/index?token={token}"
s.get(ref, timeout=20)
page_url = f"{BASE}/iecs/freight/productCertificate/productCertificate/initLoads"
s.get(page_url, headers={"Referer": ref}, timeout=20)

# 页面默认参数（从 HTML 表单提取）
PAGE_DEFAULTS = {
    "segNo": "QE000000",
    "system": "ES",
    "userNum": "062122",
    "uuCode": "U41634",
    "userNo": "QE000000",
    "companyNo": "QE000000",
    "saleNetwork": "E",
    "BSUrl": "https://ecommerce.ibaosteel.com/icsc/TLfqmAction/tLfqmActionCasNewRedirectEncrypt?",
    "CBUrl": "https://ecommerce.ibaosteel.com/icsc/channelBoard/channelBoardRedirectEncrypt?",
    "access_token": "U1ZYQno4SDE1L1NmVU02RFJJYkFneEtxZzdGZXJXK1BvNzVlWVFLZlRWST0=",
}

HEADERS = {
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "text/html, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": page_url,
}

# 尝试获取 customerId（从页面的 customerId_select）
# 先访问页面获取
customer_id = ""  # 默认空，需要从页面获取


def query_certificate(page_num=1, page_size=10,
                      delivary_from="202607", delivary_to="202609",
                      contract_num="", factory_order_num="", pack_num="", tc_num=""):
    """调用 queryCertificate API，返回 (总数, 行数据列表)"""
    body = {
        "delivaryDateFrom": delivary_from,
        "delivaryDateTo": delivary_to,
        "segNo": PAGE_DEFAULTS["segNo"],
        "system": PAGE_DEFAULTS["system"],
        "userNum": PAGE_DEFAULTS["userNum"],
        "uuCode": PAGE_DEFAULTS["uuCode"],
        "userNo": PAGE_DEFAULTS["userNo"],
        "companyNo": PAGE_DEFAULTS["companyNo"],
        "saleNetwork": PAGE_DEFAULTS["saleNetwork"],
        "contractNum": contract_num,
        "factoryOrderNum": factory_order_num,
        "packNum": pack_num,
        "tcNumTc": tc_num,
        "inDateFrom": "",
        "inDateTo": "",
        "sortField": "",
        "sortRule": "",
        "pageDomain": {"pageNum": str(page_num), "pageSize": str(page_size)},
    }
    
    api = f"{BASE}/iecs/freight/productCertificate/productCertificate/queryCertificate"
    r = s.post(api, data=json.dumps(body), headers=HEADERS, timeout=30)
    html = r.text
    
    # 检查是否错误
    if 'isError="true"' in html:
        print(f"    ❌ 查询返回错误: {html[:300]}")
        return 0, []
    
    # 提取总条数
    total_match = re.search(r'id="total"\s+value="(\d+)"', html)
    total = int(total_match.group(1)) if total_match else 0
    
    # 解析每行数据
    rows = []
    tr_pattern = re.compile(r'<tr\s+class="actives">(.*?)</tr>', re.DOTALL)
    
    for tr_match in tr_pattern.finditer(html):
        tr_content = tr_match.group(1)
        
        # 提取 checkbox 的所有属性
        ck_pattern = r'<input[^>]*name="ck_box"([^>]*)>'
        ck_match = re.search(ck_pattern, tr_content)
        if not ck_match:
            continue
        
        attrs = ck_match.group(1)
        
        def get_attr(name):
            m = re.search(rf'{name}="([^"]*)"', attrs)
            return m.group(1) if m else ""
        
        row = {
            "tcNumTc": get_attr("tcNumTc"),
            "tcNumTcRsa": get_attr("tcNumTcRsa"),
            "factoryOrderNumRsa": get_attr("factoryOrderNumRsa"),
            "netWeight": get_attr("netWeight"),
            "pieceNum": get_attr("pieceNum"),
            "boardPlank": get_attr("boardPlank"),
        }
        
        # 提取显示文本
        cert_match = re.search(r'<span>([^<]+)</span>', tr_content)
        row["certNo"] = cert_match.group(1) if cert_match else row["tcNumTc"]
        
        # 获取更多显示列
        td_texts = re.findall(r'<td[^>]*>(.*?)</td>', tr_content, re.DOTALL)
        row["factoryOrderNum"] = re.sub(r'<[^>]+>', '', td_texts[2]).strip() if len(td_texts) > 2 else ""
        row["subOrderNo"] = re.sub(r'<[^>]+>', '', td_texts[3]).strip() if len(td_texts) > 3 else ""
        row["spec"] = re.sub(r'<[^>]+>', '', td_texts[4]).strip() if len(td_texts) > 4 else ""
        row["genDate"] = re.sub(r'<[^>]+>', '', td_texts[5]).strip() if len(td_texts) > 5 else ""
        row["deliveryDate"] = re.sub(r'<[^>]+>', '', td_texts[6]).strip() if len(td_texts) > 6 else ""
        row["orderQty"] = re.sub(r'<[^>]+>', '', td_texts[7]).strip() if len(td_texts) > 7 else ""
        
        rows.append(row)
    
    return total, rows


def decrypt_certificate_info(rows, customer_id=""):
    """
    调用 decryptByCertificateInfo 解密证书信息
    关键：inList 中 tcNumTc 字段传的是 tcNumTcRsa（加密值），不是明文！
    """
    in_list = []
    for row in rows:
        tc_rsa = row["tcNumTcRsa"]
        fa_rsa = row["factoryOrderNumRsa"]
        
        # 只有加密值都不为空才加入
        if tc_rsa and fa_rsa:
            in_list.append({
                "tcNumTc": tc_rsa,          # 加密的 tcNumTcRsa
                "boardPlank": row["boardPlank"],
                "factoryOrderNum": fa_rsa,   # 加密的 factoryOrderNumRsa
                "customerId": customer_id,
            })
    
    if not in_list:
        print(f"    ⚠️ inList 为空，跳过解密")
        return None
    
    print(f"    inList 长度: {len(in_list)}")
    print(f"    第一条 tcNumTc(加密): {in_list[0]['tcNumTc'][:50]}...")
    print(f"    第一条 factoryOrderNum(加密): {in_list[0]['factoryOrderNum'][:50]}...")
    
    api = f"{BASE}/iecs/freight/productCertificate/productCertificate/decryptByCertificateInfo"
    req_headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": page_url,
    }
    
    r = s.post(api, data=json.dumps(in_list), headers=req_headers, timeout=30)
    print(f"    decrypt Status={r.status_code}")
    
    try:
        data = r.json()
        code = data.get('code')
        print(f"    decrypt Response: code={code}, msg={data.get('msg', '')}")
        
        if str(code) == '0':
            resp_data = data.get('data', {})
            print(f"    decrypt data keys: {list(resp_data.keys()) if isinstance(resp_data, dict) else 'not dict'}")
            if isinstance(resp_data, dict):
                print(f"    tcNumTcRsa(解密后): {resp_data.get('tcNumTcRsa', '')[:50]}...")
                print(f"    customerId: {resp_data.get('customerId', '')}")
                print(f"    boardPlankY: {resp_data.get('boardPlankY', '')}")
            return resp_data
        else:
            print(f"    ❌ decrypt failed: {data.get('msg', '')}")
            return None
    except Exception as e:
        print(f"    ❌ decrypt parse error: {e}")
        print(f"    Raw response: {r.text[:500]}")
        return None


def download_pdf_via_icsc(decrypt_data, output_dir, cert_no=""):
    """通过 ICSC 系统下载 PDF"""
    if not decrypt_data:
        return None
    
    # 从解密结果获取参数
    tc_num_rsa = decrypt_data.get("tcNumTcRsa", "")
    customer_id = decrypt_data.get("customerId", "")
    has_board = str(decrypt_data.get("boardPlankY", "false")).lower() == "true"
    
    # 构造 POST 参数
    params = {
        "uuCode": PAGE_DEFAULTS["uuCode"],
        "userNum": PAGE_DEFAULTS["userNum"],
        "saleNetWork": PAGE_DEFAULTS["saleNetwork"],
        "companyNo": PAGE_DEFAULTS["companyNo"],
        "userNo": PAGE_DEFAULTS["userNo"],
        "system": PAGE_DEFAULTS["system"],
        "tcNumTc": tc_num_rsa,
        "smartSegNo": "QE000000",
        "orderOwner": customer_id,
    }
    
    # 选择 URL（boardPlank 决定用 BSUrl 还是 CBUrl）
    base_url = PAGE_DEFAULTS["CBUrl"] if has_board else PAGE_DEFAULTS["BSUrl"]
    url = base_url + "access_token=" + PAGE_DEFAULTS["access_token"]
    
    print(f"    POST to: {base_url[:60]}...")
    print(f"    tcNumTc={tc_num_rsa[:50]}...")
    print(f"    orderOwner={customer_id}")
    
    try:
        r = s.post(url, data=params, timeout=60, allow_redirects=True, stream=True)
        content_type = r.headers.get('Content-Type', '')
        content_length = r.headers.get('Content-Length', 'unknown')
        print(f"    Response: status={r.status_code}, type={content_type}, length={content_length}")
        
        # 检查是否是 PDF
        if 'pdf' in content_type or 'octet-stream' in content_type or 'application/pdf' in content_type:
            # 尝试从 Content-Disposition 获取文件名
            cd = r.headers.get('Content-Disposition', '')
            filename_match = re.search(r'filename\s*=\s*"?([^";]+)"?', cd)
            filename = filename_match.group(1) if filename_match else f"{cert_no}.pdf"
            
            # 清理文件名
            filename = re.sub(r'[\\/:*?"<>|]', '_', filename)
            if not filename.endswith('.pdf'):
                filename += '.pdf'
            
            out_path = output_dir / filename
            with open(out_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            size = out_path.stat().st_size
            print(f"    ✅ Downloaded: {out_path.name} ({size} bytes)")
            return out_path
        else:
            # 可能返回 HTML 或其他
            if 'text/html' in content_type or 'html' in content_type:
                content = r.text[:500]
                print(f"    ❌ 返回 HTML 而非 PDF")
                print(f"    Content: {content[:300]}")
                
                # 保存调试文件
                debug_path = output_dir.parent / f"debug_{cert_no}_{int(time.time())}.html"
                with open(debug_path, 'w', encoding='utf-8') as f:
                    f.write(r.text)
                print(f"    Debug saved: {debug_path.name}")
            else:
                print(f"    ❌ 未知 Content-Type: {content_type}")
                # 尝试保存
                out_path = output_dir / f"{cert_no}_unknown.bin"
                with open(out_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f"    Saved as: {out_path.name} ({out_path.stat().st_size} bytes)")
            return None
            
    except Exception as e:
        print(f"    ❌ Download error: {e}")
        import traceback
        traceback.print_exc()
        return None


def download_all_certs():
    """下载所有质保书 PDF"""
    output_dir = Path("_quality_certs") / "pdfs"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 第一步：分页获取所有质保书数据
    print("\n" + "="*60)
    print("[Step 1] 获取所有质保书列表")
    print("="*60)
    
    all_rows = []
    page_size = 100
    page_num = 1
    
    while True:
        total, rows = query_certificate(
            page_num=page_num,
            page_size=page_size,
            delivary_from="202607",
            delivary_to="202609"
        )
        
        print(f"  Page {page_num}: {len(rows)} rows (total={total})")
        all_rows.extend(rows)
        
        if len(rows) < page_size:
            break
        page_num += 1
        if page_num > 100:
            break
    
    print(f"\n  共获取 {len(all_rows)} 条质保书记录")
    
    # 保存数据快照
    data_file = output_dir.parent / "cert_list.json"
    with open(data_file, 'w', encoding='utf-8') as f:
        json.dump(all_rows, f, ensure_ascii=False, indent=2)
    print(f"  数据已保存: {data_file}")
    
    # 第二步：逐个尝试下载 PDF（先测试前 3 条）
    print("\n" + "="*60)
    print("[Step 2] 下载质保书 PDF（测试前 3 条）")
    print("="*60)
    
    success_count = 0
    fail_count = 0
    test_count = min(3, len(all_rows))
    
    for i, row in enumerate(all_rows[:test_count]):
        cert_no = row['certNo']
        print(f"\n  [{i+1}/{test_count}] 质保书号={cert_no}, 钢卷号={row['tcNumTc']}")
        print(f"    tcNumTcRsa(加密)={row['tcNumTcRsa'][:50]}...")
        print(f"    factoryOrderNumRsa(加密)={row['factoryOrderNumRsa'][:50]}...")
        print(f"    boardPlank={row['boardPlank']}")
        
        # 解密
        print(f"    [解密] decryptByCertificateInfo ...")
        decrypt_result = decrypt_certificate_info([row], customer_id)
        
        if decrypt_result:
            # 下载 PDF
            print(f"    [下载] POST 到 ICSC ...")
            result = download_pdf_via_icsc(decrypt_result, output_dir, cert_no)
            if result:
                success_count += 1
            else:
                fail_count += 1
        else:
            fail_count += 1
        
        time.sleep(1)
    
    print(f"\n  测试结果: 成功={success_count}, 失败={fail_count}")
    
    # 第三步：如果测试成功，下载全部
    if success_count > 0:
        print(f"\n" + "="*60)
        print(f"[Step 3] 下载全部 {len(all_rows)} 条质保书 PDF")
        print("="*60)
        
        full_success = 0
        full_fail = 0
        
        for i, row in enumerate(all_rows):
            cert_no = row['certNo']
            
            # 解密
            decrypt_result = decrypt_certificate_info([row], customer_id)
            
            if decrypt_result:
                result = download_pdf_via_icsc(decrypt_result, output_dir, cert_no)
                if result:
                    full_success += 1
                else:
                    full_fail += 1
            else:
                full_fail += 1
            
            if (i+1) % 10 == 0:
                print(f"  进度: {i+1}/{len(all_rows)} (成功={full_success}, 失败={full_fail})")
            
            time.sleep(0.5)
        
        print(f"\n  全部完成: 成功={full_success}, 失败={full_fail}")
    else:
        print("\n  ⚠️ 测试阶段全部失败，请检查问题后再试")


if __name__ == "__main__":
    download_all_certs()
