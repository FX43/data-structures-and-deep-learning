"""重建分类导航和进度表；按题号保留表格中已有的复习记录。"""
from pathlib import Path
import os,shutil,subprocess,sys

def main():
    root=Path(__file__).resolve().parent
    bundle=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
    node=Path(os.environ.get('CODEX_NODE',bundle/'node/bin/node.exe'))
    if not node.is_file():
        found=shutil.which('node')
        if not found: raise SystemExit('需要Node.js及@oai/artifact-tool。可设置CODEX_NODE和CODEX_NODE_MODULES。')
        node=Path(found)
    env=os.environ.copy()
    if (bundle/'node/node_modules').is_dir(): env.setdefault('CODEX_NODE_MODULES',str(bundle/'node/node_modules'))
    env.setdefault('CODEX_PYTHON',sys.executable)
    result=subprocess.run([str(node),str(root/'build_excel.mjs')],cwd=root,env=env)
    raise SystemExit(result.returncode)

if __name__=='__main__': main()
