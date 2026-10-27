# Ranet Toolkit

한글·PDF 문서 변환·편집 도구 (Windows)

- HWP → HWPX·PDF 일괄 변환
- PDF 분할·병합·압축·편집
- 도면 PDF → EMF·SVG 벡터 변환
- PPT → 이미지 변환

## 내려받기

[최신 Ranet.Toolkit.exe](https://github.com/skulloom/ranet-toolkit/releases/latest/download/Ranet.Toolkit.exe) · 설치 없이 실행, 실행할 때 새 버전을 확인해 스스로 업데이트합니다.
소개·다른 프로그램: <https://jnranet.com/apps/>

## 소스에서 빌드

Windows, Python 3.10.

```
cd src
python -m pip install -r requirements.txt
cd 개발소스
python -m PyInstaller "Ranet ToolKit.spec" --noconfirm --distpath .. --workpath build
```

`src\Ranet Toolkit.exe` 가 만들어집니다. 소스로 바로 실행하려면 `src` 에서 `python "Ranet ToolKit.py"`.
HWP 변환은 한컴오피스 한글, PPT 변환은 PowerPoint 가 설치된 PC 에서 동작합니다.

## 라이선스

Copyright (C) 2026 지역농업네트워크 전남 협동조합 (Region & Agriculture NETwork Cooperative, Jeonnam)

이 프로그램은 [GNU Affero General Public License v3.0](LICENSE) 에 따라 배포합니다.
PDF 처리에 AGPL-3.0 라이선스인 [PyMuPDF](https://github.com/pymupdf/PyMuPDF) 를 쓰기 때문입니다.
보증은 없습니다(LICENSE 15·16조).

같이 쓰는 오픈소스: Qt for Python(PySide6, LGPL-3.0) · pypdf · lxml · Pillow · Requests · pywin32 · PyAutoGUI · pyperclip · python-pptx · ReportLab · PyInstaller ·
글꼴 Pretendard(SIL OFL 1.1, `src/assets/fonts/OFL.txt`). 각 소프트웨어의 저작권은 해당 저작권자에게 있습니다.

`version.txt`·`changelog.txt` 는 프로그램의 자동 업데이트가 읽는 파일입니다(옮기지 마십시오).
