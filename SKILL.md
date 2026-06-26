---
name: yeungnam-irb-protocol
description: 영남대학교병원(YUMC) IRB 제출 문서 3종을 자동 생성하는 standalone 스킬. study_info.json 단일 입력에서 (1) 연구계획서 본문 (2) 임상연구 계획서 요약[서식1] (3) 연구대상자 동의면제 점검표[서식41]를 .docx로 생성한다. 영남대 순환기내과 후향 코호트 연구 양식 기준. "IRB 제출", "연구계획서", "동의면제 점검표", "임상연구 계획서 요약", "영남대 IRB 양식"을 언급할 때 사용. 외부 하네스(prereg.json 등) 의존성 없음.
license: MIT
---

# Yeungnam IRB Protocol Writer (standalone)

영남대학교병원(YUMC) IRB 제출용 문서 3종을 단일 JSON 입력에서 자동 생성한다.
clinical-research-harness의 `protocol-writer`에서 IRB 생성 부분만 분리해, 분당서울대병원 템플릿 대신 **영남대 순환기내과 실제 양식**으로 동작하도록 재작성한 독립 스킬이다.

## 트리거
- "IRB 제출", "연구계획서 작성", "동의면제 점검표", "임상연구 계획서 요약[서식1]"
- "영남대 IRB 양식으로 후향연구 계획서 만들어줘"

## 입력 (single source of truth)
`study_info.json` — `references/study_info.schema.json` 스키마를 따른다.
모든 산출물은 이 한 파일에서 생성되므로 가설·기준·통계 계획을 두 번 입력하지 않는다.
작성 시 `examples/study_info.template.json`(빈 템플릿) 또는 `examples/study_info.sample.json`(가상 예시)을 복사해 채우는 것을 권장한다.

핵심 필드: `title`, `objective`, `design`, `investigators.pi`, `inclusion_criteria`,
`exclusion_criteria`, `primary_endpoint`, `secondary_endpoints`, `statistics`,
`data_collected`, `consent_waiver`, `references`.

## 산출물
| 파일 | 양식 | 내용 |
|---|---|---|
| `research_protocol.docx` | 연구계획서 본문 | 13개 섹션 (요약·배경·목적·관리구조·위험이익·윤리·설계·수행·데이터관리·분석·보고·보관·참고문헌) |
| `protocol_summary.docx` | 임상연구 계획서 요약 [서식 1] | 11항목 요약 표 |
| `consent_waiver_checklist.docx` | 연구대상자 동의면제 점검표 [서식 41] | 면제사유 점검 표 + YES/NO 판정 |

## 절차
1. 사용자에게 `study_info.json` 위치를 확인. 없으면 `examples/study_info.sample.json`을
   복사해 사용자 연구 내용으로 채우도록 안내 (또는 대화로 받아 직접 채워준다).
2. 생성 실행:
   ```bash
   python scripts/generate_irb.py /path/to/study_info.json --outdir /path/to/out
   ```
   단일 문서만 필요하면 `--only protocol|summary|waiver`.
3. 생성된 3종 .docx를 사용자 IRB 폴더에 저장하고 `present_files`로 제시.
4. 사용자 검토 후 .docx를 직접 수정하거나, JSON을 고쳐 재생성.

## 비타협 정책
- **Single source of truth**: study_info.json 외 정보를 문서에 자유 추가하지 않는다.
- **Citation Grounding**: `references`에서 `pmid` 또는 `doi`가 있는 항목만 참고문헌에
  출력된다. PMID/DOI 미동반 인용은 자동 제외되고 경고가 표시된다. 인용 환각 금지.

## 한계 (의도된)
- .docx 생성까지만 담당. **BRIA 등 IRB 시스템에 자동 제출하지 않음** — 사용자가 직접 제출.
- 양식은 영남대 순환기내과 후향 코호트 기준. 타 과/타 기관은 템플릿 조정 필요.
- 동의면제 자격(서식41)은 기관 IRB가 최종 판단. 본 스킬은 점검표 초안만 작성한다.
- 표본 크기·검정력 산출은 포함하지 않는다 (별도 검정 단계 필요).

## 의존성
- Python `python-docx` (`pip install python-docx`)
- 외부 하네스