"""
通过 GitHub API 新增 DH1_USERNAME / DH1_PASSWORD 两个 Secrets（存旧账号 HONG@shihe.donghuo）
然后修改 export_lindiao.yml 让它读 DH1_xxx 而不是 DH_xxx
"""
import base64, json, os, requests
from nacl import public

# 真实 GH_TOKEN 请从环境变量注入（或 gh auth token），禁止硬编码进仓库
GH_TOKEN = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN") or ""
REPO = os.environ.get("GITHUB_REPOSITORY") or "Richard-shihe/donghuo-export"
HEADERS = {
    "Authorization": f"Bearer {GH_TOKEN}",
    "Accept": "application/vnd.github+json",
}

# 1. 获取仓库公钥
r = requests.get(f"https://api.github.com/repos/{REPO}/actions/secrets/public-key",
                headers=HEADERS, timeout=15)
r.raise_for_status()
key_data = r.json()
key_id = key_data["key_id"]
pub_key_b64 = key_data["key"]
print(f"[1] 仓库公钥: key_id={key_id}, key_len={len(pub_key_b64)}")

pub_key = public.PublicKey(base64.b64decode(pub_key_b64))
sealed_box = public.SealedBox(pub_key)

# 2. 创建 DH1_USERNAME 和 DH1_PASSWORD（真实值从环境变量注入，禁止硬编码进仓库）
secrets_to_set = {
    "DH1_USERNAME": os.environ.get("DH1_USERNAME") or os.environ.get("DH_USERNAME") or "",
    "DH1_PASSWORD": os.environ.get("DH1_PASSWORD") or os.environ.get("DH_PASSWORD") or "",
}
for name, value in secrets_to_set.items():
    encrypted = sealed_box.encrypt(value.encode("utf-8"))
    encrypted_b64 = base64.b64encode(encrypted).decode()
    r = requests.put(
        f"https://api.github.com/repos/{REPO}/actions/secrets/{name}",
        headers=HEADERS,
        json={"encrypted_value": encrypted_b64, "key_id": key_id},
        timeout=15,
    )
    print(f"[2] 创建/更新 Secret {name}: HTTP {r.status_code}")
    if r.status_code not in (201, 204):
        print(f"    响应: {r.text[:200]}")
        raise SystemExit(1)

# 3. 列出当前所有 secrets 确认
r = requests.get(f"https://api.github.com/repos/{REPO}/actions/secrets",
                 headers=HEADERS, timeout=15)
print("\n[3] 当前仓库 Secrets 列表:")
for s in r.json().get("secrets", []):
    print(f"  - {s['name']}  (updated_at={s['updated_at']})")
