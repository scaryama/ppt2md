# 📊 PPT to Markdown 변환기

PowerPoint 파일(.ppt, .pptx)을 Markdown 형식으로 자동 변환하는 도구입니다. 모든 슬라이드를 하나의 Markdown 파일로 통합합니다. GUI와 커맨드라인 두 가지 방식으로 사용할 수 있습니다.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/PyQt6-6.0+-green.svg)](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## ✨ 주요 기능

- 📄 **PowerPoint 파일 읽기**: .ppt 및 .pptx 형식 파일의 모든 슬라이드 자동 인식
- 📝 **Markdown 변환**: 모든 슬라이드를 하나의 Markdown 파일로 변환
- 🖱️ **드래그 앤 드롭**: GUI에서 파일을 드래그하여 간편하게 변환
- 📊 **실시간 진행 상황**: 변환 진행 상황을 실시간으로 확인
- 🔄 **비동기 처리**: 변환 중에도 GUI가 멈추지 않음
- 📦 **실행 파일 지원**: Python 없이도 실행 가능한 .exe 파일 생성
- 📋 **내용 추출**: 슬라이드의 텍스트, 표, 리스트, 이미지 플레이스홀더 추출

## 🚀 빠른 시작

### 설치

```bash
# 저장소 클론
git clone <repository-url>
cd ppt2md.git

# 가상 환경 생성 및 활성화
python -m venv venv
venv\Scripts\activate  # Windows

# 의존성 설치
pip install python-pptx PyQt6
```

### 사용 방법

#### GUI 버전 (권장)

```bash
python ppt2md_gui.py
```

1. GUI 창이 열리면 PPT/PPTX 파일을 드래그 앤 드롭하거나 "파일 선택" 버튼 클릭
2. 변환 진행 상황 확인
3. 변환 완료 후 원본 파일과 같은 위치에 Markdown 파일 생성

#### 커맨드라인 버전

```bash
# 기본 사용
python ppt2md.py "경로/파일.pptx"

# 출력 디렉토리 지정
python ppt2md.py "경로/파일.pptx" "출력/경로"
```

## 📋 요구사항

- Python 3.8 이상
- Windows 10 이상 (GUI 버전)

### 필수 패키지

```
python-pptx>=0.6.0
PyQt6>=6.0.0
```

## 📁 프로젝트 구조

```
ppt2md.git/
├── ppt2md.py              # 커맨드라인 버전 (핵심 변환 로직)
├── ppt2md_gui.py          # GUI 버전
├── build_exe.bat          # Windows 배치 빌드 스크립트
├── build_exe.ps1          # PowerShell 빌드 스크립트
├── README_BUILD.md        # 빌드 가이드
└── doc/                   # 변환 대상 PowerPoint 파일들
```

## 🔧 빌드 및 배포

### 실행 파일(.exe) 생성

```bash
# 배치 파일 사용
build_exe.bat

# 또는 PowerShell 스크립트 사용
.\build_exe.ps1
```

빌드 완료 후 `dist\PPT2Markdown.exe` 파일이 생성됩니다.

자세한 빌드 가이드는 [README_BUILD.md](README_BUILD.md)를 참고하세요.

## 📖 사용 예시

### 입력 파일
- `presentation.pptx` (여러 슬라이드 포함)

### 출력 파일
모든 슬라이드가 포함된 하나의 Markdown 파일:
- `presentation.md`

### Markdown 파일 형식

```markdown
# 프레젠테이션 이름
*원본 파일: presentation.pptx*
*총 슬라이드 수: 5*
---

## Slide 1

### 제목

내용 텍스트
- 불릿 포인트 1
- 불릿 포인트 2

---

## Slide 2

| 컬럼1 | 컬럼2 | 컬럼3 |
|-------|-------|-------|
| 데이터1 | 데이터2 | 데이터3 |

[Image: Slide 2]

---
```

## 🛠️ 기술 스택

- **Python 3.8+**: 프로그래밍 언어
- **PyQt6**: GUI 프레임워크
- **python-pptx**: PowerPoint 파일 읽기 및 파싱
- **PyInstaller**: 실행 파일 빌드 도구

## 📝 주요 기능 상세

### 내용 추출
- **텍스트**: 텍스트 박스 및 도형에서 텍스트 추출
- **표**: PowerPoint 표를 Markdown 표로 변환
- **리스트**: 불릿 포인트 및 번호 매기기 리스트 보존
- **이미지**: 이미지를 `[Image: Slide N]` 플레이스홀더 텍스트로 대체
- **제목**: 슬라이드 제목 자동 감지 및 포맷팅

### 특수 문자 처리
- 파이프 문자(`|`): 표에서 `\|`로 이스케이프
- 줄바꿈: 텍스트 내용에서 보존
- 빈 슬라이드: "*This slide is empty.*"로 표시

### 에러 처리
- 파일 존재 여부 확인
- 빈 슬라이드 처리
- 개별 슬라이드 처리 실패 시에도 계속 진행
- 상세한 에러 메시지 및 트레이스백 제공

## 🤝 기여하기

버그 리포트, 기능 제안, Pull Request를 환영합니다!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 🙏 감사의 말

이 프로젝트는 PowerPoint 기반 프레젠테이션 문서를 Markdown 형식으로 변환하여 버전 관리 시스템 및 문서 워크플로우와의 통합을 용이하게 하기 위해 개발되었습니다.

## 📚 관련 문서

- [빌드 가이드](README_BUILD.md)
- [영문 문서](README.md)

---

**Made with ❤️ for better documentation workflow**

