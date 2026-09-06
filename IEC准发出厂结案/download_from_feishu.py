"""从飞书云盘下载最新的 准发下载_*.xlsx 文件"""
import os
import requests

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
FOLDER_TOKEN = "QIvpfoJnqlg8IIdT1PXctnB3nne"
NO_PROXY = {"http": None, "https": None}


def get_token():
    app_id = os.environ.get("FEISHU_APP_ID", "")
    app_secret = os.environ.get("FEISHU_APP_SECRET", "")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15, proxies=NO_PROXY,
    )
    return r.json()["tenant_access_token"]


def main():
    token = get_token()
    print(f"token OK\n")

    # 列出文件夹下的文件
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files"
    params = {"folder_id": FOLDER_TOKEN, "page_size": 50}
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, params=params, timeout=15, proxies=NO_PROXY)
    data = r.json()
    print(f"[调试] code={data.get('code')} msg={data.get('msg')}")
    print(f"[调试] data keys: {list((data.get('data') or {}).keys())}")
    files = (data.get("data") or {}).get("files") or []
    print(f"[调试] files count: {len(files)}")
    if files:
        print(f"[调试] first file keys: {list(files[0].keys())}")
        print(f"[调试] first file: {files[0]}")
    if data.get("code") != 0:
        print(f"[错误] 列文件: {data}")
        return
    xlsx_files = [f for f in files if f.get("name", "").endswith(".xlsx")
                  and "准发下载" in f.get("name", "")]
    xlsx_files.sort(key=lambda x: x.get("name", ""), reverse=True)

    print(f"文件夹共 {len(files)} 个文件，其中 准发下载 xlsx {len(xlsx_files)} 个:\n")
    for f in xlsx_files[:5]:
        print(f"  {f.get('name')}  token={f.get('token')}  id={f.get('id')}")

    if not xlsx_files:
        print("没找到 xlsx 文件")
        return

    target = xlsx_files[0]
    print(f"\n下载最新: {target['name']}")

    dl_url = f"{FEISHU_OPEN_BASE}/drive/v1/files/{target['token']}/download"
    r = requests.get(dl_url, headers=headers, timeout=30, proxies=NO_PROXY)
    if r.status_code != 200:
        print(f"[错误] 下载失败: {r.status_code} {r.text[:200]}")
        return

    out_path = target["name"]
    with open(out_path, "wb") as fp:
        fp.write(r.content)
    print(f"已下载: {out_path} ({len(r.content)} bytes)")


if __name__ == "__main__":
    main()
