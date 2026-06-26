# Yeungnam IRB Protocol Writer

A standalone [Cowork](https://www.anthropic.com) skill that generates **three Yeungnam University Hospital (YUMC) IRB submission documents** from a single JSON input — built for retrospective cohort studies in the Division of Cardiology.

> 한국어 안내는 아래 [한국어](#한국어) 섹션을 참고하세요.

## What it generates

From one `study_info.json`, the skill produces:

| Output | Korean form |
|---|---|
| `research_protocol.docx` | 연구계획서 본문 (13 sections) |
| `protocol_summary.docx` | 임상연구 계획서 요약 [서식 1] |
| `consent_waiver_checklist.docx` | 연구대상자 동의면제 점검표 [서식 41] |

## Quick start

```bash
pip install python-docx
cp examples/study_info.template.json study_info.json   # fill in your study
python scripts/generate_irb.py study_info.json --outdir ./out
```

Generate a single document with `--only protocol|summary|waiver`.

## Install as a Cowork skill

Package this folder as a `.skill` (zip) and use **Save skill** in Cowork, or register the folder via **Settings → Capabilities**. See [INSTALL.md](INSTALL.md).

## Design policies

- **Single source of truth** — every document is generated from one `study_info.json`. Edit the JSON once; all three outputs update together.
- **Citation grounding** — a reference is emitted only if it carries a `pmid`, `doi`, or `registry` id. Ungrounded citations are dropped with a warning, preventing hallucinated references.

## Limitations (by design)

- Generates `.docx` only; does **not** auto-submit to any IRB system (e.g. BRIA).
- Templated for YUMC Cardiology retrospective cohorts. Other departments/institutions need template adjustment.
- The consent-waiver checklist (Form 41) uses generic waiver criteria; replace `criteria` in `scripts/generate_irb.py` (`build_waiver()`) with your institution's exact wording. Final waiver eligibility is decided by the IRB.
- No sample-size / power calculation.

## No PHI

All example data in this repository is **fictional**. No real patient or investigator information is included. Do not commit filled `study_info.json` files containing real study data — see `.gitignore`.

## License

MIT — see [LICENSE](LICENSE).

---

## 한국어

영남대학교병원(YUMC) 순환기내과 후향연구용 **IRB 제출 문서 3종**을 `study_info.json` 한 파일에서 자동 생성하는 독립형 Cowork 스킬입니다.

### 생성 문서
1. 연구계획서 본문 (`research_protocol.docx`) — 13개 섹션
2. 임상연구 계획서 요약 [서식 1] (`protocol_summary.docx`)
3. 연구대상자 동의면제 점검표 [서식 41] (`consent_waiver_checklist.docx`)

### 빠른 시작
```bash
pip install python-docx
cp examples/study_info.template.json study_info.json   # 본인 연구로 채우기
python scripts/generate_irb.py study_info.json --outdir ./out
```

### 핵심 정책
- **Single source of truth**: 모든 문서는 study_info.json 하나에서 생성. 수정은 JSON만 고치면 3종 일괄 반영.
- **Citation grounding**: `pmid`/`doi`/`registry` 중 하나가 있는 인용만 출력. 근거 없는 인용은 자동 제외.

### 한계
- .docx 생성까지만 담당, IRB 시스템 자동 제출 없음. 영남대 순환기내과 후향연구 기준 양식.
- 동의면제 점검항목은 일반 기준이므로 기관 실제 서식으로 교체 필요. 면제 자격 최종 판단은 IRB.

### 개인정보
이 저장소의 모든 예시는 **가상 데이터**입니다. 실제 연구 정보가 담긴 `study_info.json`은 커밋하지 마세요(`.gitignore` 참고).
