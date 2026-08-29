"""
核对 L6ER000579 的准发量/出厂量在飞书进度表中的值
"""
import os
import requests

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"
PROGRESS_TABLE = "tblvugnoJPS8GrpX"
NO_PROXY = {"http": None, "https": None}
TARGET = "L6ER000579"


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

    # 用 list API 拉取进度表全部记录，找到资源号 == L6ER000579 的记录
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}"
           f"/tables/{PROGRESS_TABLE}/records")
    page_token = ""
    found = []
    page = 0
    total = 0
    while True:
        page += 1
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                         params=params, timeout=30, proxies=NO_PROXY)
        data = r.json()
        if data.get("code") != 0:
            print(f"[错误] 页 {page}: {data}")
            break
        d = data.get("data") or {}
        items = d.get("items") or []
        for it in items:
            fields = it.get("fields") or {}
            # 找所有字段里含有 L6ER000579 的
            for k, v in fields.items():
                txt = str(v)
                if TARGET in txt:
                    found.append((it.get("record_id"), k, v))
            total += len(items)
        has_more = d.get("has_more", False)
        page_token = d.get("page_token") or ""
        if not has_more or not page_token:
            break

    print(f"共扫 {total} 条进度表记录\n")
    print(f"=== 含 '{TARGET}' 的字段 ===")
    if not found:
        print("没找到")
        return

    # 找到资源号 = L6ER000579 的记录，打印全部字段
    for rid, fname, val in found:
        # 重新拿这条记录的全部字段
        r = requests.get(f"{url}/{rid}",
                         headers={"Authorization": f"Bearer {token}"},
                         timeout=15, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") == 0:
            rec = (d.get("data") or {}).get("record") or {}
            fields = rec.get("fields") or {}
            print(f"\nrecord_id: {rid}")
            print(f"命中字段: {fname}")
            print(f"全部字段值:")
            for k, v in fields.items():
                print(f"  {k} = {str(v)[:200]}")


if __name__ == "__main__":
    main()
