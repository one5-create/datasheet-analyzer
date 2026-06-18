"""Google Gemini API를 사용한 Data Sheet 비교 분석 에이전트."""
import os
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
  "recommendation": "Overall analysis and substitution recommendation"
}}

--- Data Sheet Content ---
{content}
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

    def analyze(self, parsed_files: list[dict], lang: str = "Korean", roles: list[str] = None) -> dict:
        """파싱된 파일 목록을 받아 Gemini로 비교 분석합니다.
        
        Args:
            roles: 각 파일의 역할 리스트 ("reference" | "comparison").
                   첫 번째 파일이 기준(reference)이고 나머지는 비교 대상.
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
        prompt = ANALYSIS_PROMPT.format(
            count=len(parsed_files),
            lang=lang,
            content=combined_content,
        )

        print(f"  AI 모델로 분석 중...")
        for attempt in range(3):
            try:
                response = self.client.models.generate_content(
                    model=self.model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                        response_mime_type="application/json",
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
            result = json.loads(response.text)
        except json.JSONDecodeError:
            result = {"raw_response": response.text}

        return result
