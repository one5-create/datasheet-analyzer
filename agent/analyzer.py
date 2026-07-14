"""Google Gemini API를 사용한 Data Sheet 비교 분석 에이전트."""
import os
import re
import json
import time
from google import genai
from google.genai import types
from google.genai import errors as genai_errors
from dotenv import load_dotenv

load_dotenv()

ANALYSIS_PROMPT = """You are an expert hardware/electronics engineer AI that professionally analyzes technical product Data Sheets.
Analyze the {count} Data Sheet(s) below and produce a detailed comparison report.

Output language: {lang}

File roles:
- REFERENCE: the component currently in use in the circuit (the baseline)
- COMPARISON TARGET: component(s) being evaluated as potential drop-in replacements for the reference

Analysis scope:
1. **Spec/Specification Comparison**: key technical specs, package, pinout, ratings, dimensions, materials, etc.
2. **Performance Metrics Comparison**: speed, capacity, efficiency, accuracy, and all numerical data
3. **Content Summary**: core features and strengths of each product
4. **Circuit Interchangeability Analysis**: The PRIMARY goal of this section is to determine whether the
   COMPARISON TARGET(s) can be used as drop-in replacements for the REFERENCE component IN THE SAME
   CIRCUIT DESIGN (same PCB, same BOM position). Evaluate the following in detail:
   - Pin-for-pin compatibility (package footprint, pin assignments)
   - Electrical parameter compatibility (voltage ratings, current ratings, thresholds, timing)
   - Performance impact if substituted (efficiency change, thermal change, switching behavior change)
   - Whether the circuit requires any modification (resistor value change, layout change, firmware change, etc.)
   - Risk level of substitution: None / Low / Medium / High
   - Specific conditions under which substitution IS safe
   - Specific cautions or failure risks if substituted without care

Output MUST follow this exact JSON schema:
{{
  "summary": "Overall comparison summary (2-3 sentences)",
  "products": [
    {{
      "file_name": "filename",
      "product_name": "product name (inferred)",
      "manufacturer": "manufacturer / brand name (inferred from datasheet, e.g. Texas Instruments, STMicroelectronics, Samsung, etc. Write 'Unknown' if not found)",
      "key_specs": {{"spec name": "value"}},
      "performance": {{"metric name": "value"}},
      "highlights": ["key feature 1", "key feature 2"]
    }}
  ],
  "comparison_table": [
    {{
      "category": "comparison item",
      "values": {{"filename": "value"}}
    }}
  ],
  "interchangeability": {{
    "verdict": "Drop-in Compatible / Conditionally Replaceable / Not Replaceable",
    "substitutable": true,
    "risk_level": "None / Low / Medium / High",
    "pin_compatible": true,
    "package_compatible": true,
    "electrical_compatible": true,
    "circuit_modifications_required": ["modification 1 (or 'None required')"],
    "safe_conditions": ["condition under which substitution is safe"],
    "cautions": ["specific caution or failure risk if substituted carelessly"],
    "performance_impact": ["performance change when substituted, e.g. '5% efficiency drop at high frequency'"]
  }},
  "circuit_analysis": {{
    "provided": false,
    "component_locations": ["where the reference component appears in the schematic (e.g. U1, Q3)"],
    "surrounding_components": ["key passive/active components directly connected to the reference part (values, types)"],
    "circuit_constraints": ["circuit-level constraints visible in the schematic that affect substitution"],
    "layout_notes": ["PCB layout or routing observations relevant to substitution"],
    "schematic_based_modifications": ["additional modifications identified FROM the schematic that are NOT obvious from datasheet alone"],
    "overall_circuit_assessment": "One-paragraph assessment of substitution feasibility based purely on the schematic review"
  }},
  "recommendation": "Overall analysis and substitution recommendation"
}}

--- Data Sheet Content ---
{content}
{circuit_section}
"""


class DataSheetAnalyzer:
    def __init__(self):
        # 1) 환경변수 (.env 또는 시스템)
        api_key = os.getenv("GEMINI_API_KEY")
        model   = os.getenv("GEMINI_MODEL", "models/gemini-2.5-flash")

        # 2) Streamlit Cloud secrets (배포 환경)
        if not api_key:
            try:
                import streamlit as st
                api_key = st.secrets.get("GEMINI_API_KEY", "")
                model   = st.secrets.get("GEMINI_MODEL", model)
            except Exception:
                pass

        if not api_key:
            raise ValueError(
                "GEMINI_API_KEY가 설정되지 않았습니다.\n"
                "로컬: .env 파일에 GEMINI_API_KEY=your_key 추가\n"
                "Streamlit Cloud: Secrets 메뉴에 GEMINI_API_KEY 등록"
            )
        self.client = genai.Client(api_key=api_key)
        self.model  = model

    def analyze(self, parsed_files: list[dict], lang: str = "Korean",
                roles: list[str] = None, circuit_images: list[dict] = None) -> dict:
        """파싱된 파일 목록을 받아 Gemini로 비교 분석합니다.
        
        Args:
            roles: 각 파일의 역할 리스트 ("reference" | "comparison").
            circuit_images: 회로도 이미지 리스트 [{"file_name", "mime_type", "image_bytes"}].
        """
        if roles is None:
            roles = ["reference"] + ["comparison"] * (len(parsed_files) - 1)

        content_parts = []
        for i, (file_data, role) in enumerate(zip(parsed_files, roles), 1):
            role_label = (
                "REFERENCE (currently used component / 기존 사용 부품)"
                if role == "reference"
                else "COMPARISON TARGET (replacement candidate / 비교 대상 부품)"
            )
            content_parts.append(
                f"[File {i}: {file_data['file_name']}] [{role_label}]\n{file_data['full_text']}"
            )

        combined_content = "\n\n" + "=" * 60 + "\n\n".join(content_parts)

        # 회로도 이미지 없으면 빈 문자열, 있으면 분석 지시 추가
        if circuit_images:
            names = ", ".join(img["file_name"] for img in circuit_images)
            circuit_section = (
                f"\n\n--- Circuit Diagram(s): {names} ---\n"
                "The circuit diagram image(s) are attached to this message.\n"
                "Populate the \"circuit_analysis\" field in your JSON output with:\n"
                "- \"provided\": true\n"
                "- \"component_locations\": where the REFERENCE component appears in the schematic\n"
                "- \"surrounding_components\": directly connected passives/actives with values\n"
                "- \"circuit_constraints\": constraints visible in schematic affecting substitution\n"
                "- \"layout_notes\": PCB/routing observations\n"
                "- \"schematic_based_modifications\": extra mods identified ONLY from schematic\n"
                "- \"overall_circuit_assessment\": one-paragraph schematic-based feasibility summary\n"
                "Also revise circuit_modifications_required and cautions in interchangeability based on the schematic."
            )
        else:
            circuit_section = "\nNo circuit diagram provided. Set circuit_analysis.provided = false and leave other circuit_analysis fields empty."

        prompt = ANALYSIS_PROMPT.format(
            count=len(parsed_files),
            lang=lang,
            content=combined_content,
            circuit_section=circuit_section,
        )

        print(f"  AI 모델로 분석 중...")
        # 멀티모달 콘텐츠: 텍스트 프롬프트 + 회로도 이미지()
        contents: list = [types.Part.from_text(text=prompt)]
        if circuit_images:
            for img in circuit_images:
                contents.append(
                    types.Part.from_bytes(
                        data=img["image_bytes"],
                        mime_type=img["mime_type"],
                    )
                )
        for attempt in range(3):
            try:
                response = self.client.models.generate_content(
                    model=self.model,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                        response_mime_type="application/json",
                        max_output_tokens=65536,
                    ),
                )
                break
            except genai_errors.ServerError as e:
                if attempt < 2:
                    wait = (attempt + 1) * 15
                    print(f"  서버 과부하, {wait}초 후 재시도... ({attempt + 1}/3)")
                    time.sleep(wait)
                else:
                    raise

        try:
            result = _parse_json_response(response.text)
        except Exception:
            result = {"raw_response": response.text}

        return result


def _parse_json_response(text: str) -> dict:
    """Gemini 응답에서 JSON을 추출합니다 (다양한 포맷 대응)."""
    # BOM 및 선행/후행 공백 제거
    text = text.strip('\ufeff').strip()

    # 1차: 직접 파싱
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2차: 마크다운 코드 블록 (```json ... ```) 제거
    code_block = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', text)
    if code_block:
        try:
            return json.loads(code_block.group(1).strip('\ufeff').strip())
        except json.JSONDecodeError:
            pass

    # 3차: 첫 { } 추출
    brace_match = re.search(r'\{[\s\S]*\}', text)
    if brace_match:
        try:
            return json.loads(brace_match.group())
        except json.JSONDecodeError:
            pass

    # 4차: JSON이 중간에 잘린 경우 — 열린 괄호/대괄호를 닫아 복구 시도
    fragment = brace_match.group() if brace_match else text
    stack, result_chars = [], []
    for ch in fragment:
        result_chars.append(ch)
        if ch in ('{', '['):
            stack.append(ch)
        elif ch == '}' and stack and stack[-1] == '{':
            stack.pop()
        elif ch == ']' and stack and stack[-1] == '[':
            stack.pop()
    # 닫히지 않은 괄호 지운
    for opener in reversed(stack):
        result_chars.append('}' if opener == '{' else ']')
    try:
        return json.loads(''.join(result_chars))
    except json.JSONDecodeError:
        pass

    return {"raw_response": text}
