#!/bin/bash

rm -rf *.spec dist build *.zip

bash -c "pyinstaller half.py --noconfirm"
mkdir dist/half/ddddocr
cp /opt/homebrew/lib/python3.10/site-packages/ddddocr/*.onnx dist/half/ddddocr

cp -f password.txt ./dist
cp -f exec.sh ./dist/执行
cp -f default.xls ./dist

zip -r "$(date +'%Y-%m-%d')-mac-m1.zip" dist/
