import subprocess, os
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")
r = subprocess.run(["git", "diff", "HEAD", "--", "feishu_uploader/export_lindiao.py"],
                   capture_output=True, text=True)
print("stdout len:", len(r.stdout))
print("stderr len:", len(r.stderr))
print("--- stdout head 300 ---")
print(r.stdout[:300])
print("--- stderr ---")
print(r.stderr[:200])
