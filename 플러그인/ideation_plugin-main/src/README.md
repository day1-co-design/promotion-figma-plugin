# Source Boundaries

Figma 플러그인은 최종적으로 `ui.html`과 `code.js`를 실행하지만, 작업 충돌을 줄이기 위해 수정 영역을 나눕니다.

## 작업 영역

- `src/output`: Figma 와이어프레임 아웃풋 정책만 수정
- `ui.html`: 입력 UI, 상태 메시지, 소스 선택 흐름
- `code.js`: Figma 노드 생성, PPTX 파싱, 분석안 생성
- `google-apps-script-proxy.gs`: Google Slides pageElements 프록시
- `slides-proxy-worker.js`: 공개 링크/Worker 보조 프록시

## 작업자별 권장 분리

- 아웃풋 작업자: `src/output/*`
- 데이터 추출 작업자: `google-apps-script-proxy.gs`, `slides-proxy-worker.js`, PPTX 추출 함수
- UI 작업자: `ui.html`의 입력/상태 영역
- Figma 렌더링 작업자: `code.js`의 `createWireframe*` 함수

아웃풋만 수정하는 작업자는 `src/output/wireframe-output-policy.json`을 바꾸고 `scripts/sync-output-policy.js`만 실행합니다.
