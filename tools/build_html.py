# -*- coding: utf-8 -*-
"""screens/app.html(Artifact 원본) → screens/index.html(로컬 브라우저용 단독 파일)"""
import pathlib
root = pathlib.Path(__file__).resolve().parent.parent
body = (root / "screens/app.html").read_text(encoding="utf-8")
head = ('<!doctype html>\n<html lang="ko">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n')
(root / "screens/index.html").write_text(head + body + "\n</body>\n</html>\n", encoding="utf-8")
print("saved screens/index.html")
