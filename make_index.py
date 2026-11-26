# -*- coding: utf-8 -*-
"""推送后运行：把真实仓库地址、commit、发布时间写进 index.json

用法：
  python make_index.py --repo https://github.com/你的账号/仓库名
"""
import argparse
import json
import subprocess
from datetime import datetime, timezone

ap = argparse.ArgumentParser()
ap.add_argument('--repo', required=True, help='GitHub 仓库地址，如 https://github.com/me/dsh-index')
ap.add_argument('--path', default='packages/dsh-bid-docs')
ap.add_argument('--index', default='index.json')
args = ap.parse_args()

commit = subprocess.run(['git', 'rev-parse', 'HEAD'], capture_output=True, text=True).stdout.strip()
if len(commit) != 40:
    raise SystemExit('取不到 commit，请先在仓库里 git commit')
ts = int(datetime.now(timezone.utc).timestamp())

doc = json.load(open(args.index, encoding='utf-8'))
for plugin in doc['plugins']:
    if plugin.get('repository'):
        plugin['repository']['url'] = args.repo
        plugin['repository']['directory'] = args.path
    for v in plugin['versions']:
        v['source'] = {'type': 'git', 'url': args.repo, 'commit': commit, 'path': args.path}
        v['publishTimestamp'] = ts
json.dump(doc, open(args.index, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
print('已写入 commit', commit, '→', args.index)
