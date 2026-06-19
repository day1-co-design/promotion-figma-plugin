# Output Editing Boundary

아웃풋만 수정할 때는 이 폴더만 수정합니다.

## 수정 가능한 파일

- `wireframe-output-policy.json`

## 이 파일에서 바꿀 수 있는 것

- Figma 와이어프레임 기준 폭
- 슬라이드 간격과 상하 여백
- 캔버스/placeholder 색상과 투명도
- 이미지 목업 투명도
- 정확 모드/호환 모드 상태 문구
- 좌표 누락 안내 문구

## 수정하면 안 되는 영역

- Google Slides 링크 처리
- Apps Script 프록시 처리
- PPTX XML 좌표 추출
- `google-apps-script-proxy.gs`
- `slides-proxy-worker.js`

## 반영 방법

```bash
node scripts/sync-output-policy.js
node wireframe-regression-tests.js
```

Figma는 최종 실행 파일로 `ui.html`과 `code.js`를 읽기 때문에, 정책 파일을 바꾼 뒤 반드시 동기화 스크립트를 실행해야 합니다.
