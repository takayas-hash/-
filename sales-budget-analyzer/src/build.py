#!/usr/bin/env python3
"""ツールのHTMLを組み立てる。

  python3 sales-budget-analyzer/src/build.py [SheetJSのxlsx.full.min.jsのパス]

- dist/app.html            : Claudeのページ（Artifact）用。<!doctype> などは付けない
- index.html               : ブラウザで開く版（SheetJSはCDNから読み込む）
- 予算差異アナライザー.html : SheetJSを埋め込んだ1ファイル版（オフラインで動く）
  SheetJSのパスを渡したときだけ作る（npm i xlsx@0.18.5 の dist/xlsx.full.min.js）
"""
import os, sys
here = os.path.dirname(os.path.abspath(__file__))
root = os.path.dirname(here)
t = open(os.path.join(here, 'template.html'), encoding='utf-8').read()
c = open(os.path.join(here, 'core.js'), encoding='utf-8').read()
page = t.replace('/*CORE*/', c)
os.makedirs(os.path.join(root, 'dist'), exist_ok=True)
open(os.path.join(root, 'dist', 'app.html'), 'w', encoding='utf-8').write(page)
doc = ('<!doctype html>\n<html lang="ja"><head><meta charset="utf-8">'
       '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"></head><body>\n'
       + page + '\n</body></html>\n')
open(os.path.join(root, 'index.html'), 'w', encoding='utf-8').write(doc)
if len(sys.argv) > 1:
    lib = open(sys.argv[1], encoding='utf-8').read()
    tag = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
    assert doc.count(tag) == 1
    one = doc.replace(tag, '<!-- SheetJS Community Edition 0.18.5 (Apache-2.0) を埋め込み。インターネットに接続していなくても動きます -->\n<script>\n' + lib + '\n</script>')
    open(os.path.join(root, '予算差異アナライザー.html'), 'w', encoding='utf-8').write(one)
print('built')
