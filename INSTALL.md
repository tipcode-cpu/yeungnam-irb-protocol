# 설치 및 사용 안내 (Install & Usage)

영남대학교병원(YUMC) 순환기내과 후향연구용 IRB 제출 문서 3종을 `study_info.json` 한 파일에서 생성합니다.

## 1. 설치

### 방법 A — .skill 패키지 (권장)
1. 이 폴더를 zip으로 압축하고 확장자를 `.skill`로 바꿉니다 (예: `yeungnam-irb-protocol.skill`).
   ```bash
   zip -r yeungnam-irb-protocol.skill . -x ".git/*"
   ```
2. Cowork에서 `.skill` 파일을 열고 **Save skill** 버튼으로 설치합니다.

### 방법 B — 수동 등록
Cowork **Settings → Capabilities** 에서 이 폴더(`SKILL.md` 포함)를 스킬로 추가합니다.

### 방법 C — CLI 직접 실행 (설치 없이)
```bash
pip install python-docx
python scripts/generate_irb.py study_info.json --outdir ./out
```

## 2. 사용법

1. `examples/study_info.template.json`을 복사해 본인 연구 내용으로 채웁니다.
   - `_`로 시작하는 키와 `<...>` 자리는 작성 후 삭제/교체.
   - 작성 참고용으로 `examples/study_info.sample.json`(가상 예시)을 보세요.
2. 생성: Cowork에서 JSON 위치를 알려주고 "IRB 문서 3종 만들어줘" 라고 요청하거나, 위 CLI 실행.
3. 생성된 .docx를 검토 후 수정하거나 JSON을 고쳐 재생성. IRB File No.·점검일은 제출 단계에서 기입.

## 3. 필수 입력 필드
`title`, `objective`, `design`, `investigators.pi`, `inclusion_criteria`, `exclusion_criteria`, `primary_endpoint`, `statistics`

## 4. 주의
- 실제 환자/연구자 정보가 담긴 `study_info.json`은 **공개 저장소에 커밋하지 마세요** (`.gitignore`에 기본 제외).
- 동의면제(서식 41) 항목 문구는 기관 최신 서식과 다를 수 있으니 확인 후 사용하세요.
