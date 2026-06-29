# -*- coding: utf-8 -*-
r"""
Single-file VLM checklist evaluator for chart visualization tasks.

This public version keeps only the visual checklist evaluation path:
- read tasks from tasks.json
- find model-generated .xlsx / .png outputs
- export chart objects from .xlsx to PNG when needed
- call a VLM with checklist criteria
- save per-task score + acc and an overall summary

Notes:
- No pivot logic
- No service container / config.py dependency
- No API key stored in code

Examples (PowerShell):
    $env:VLM_API_KEY="YOUR_API_KEY"
    python run_visual_vlm_checklist_eval.py `
      --tasks-json "C:\path\to\dataset.json" `
      --output-dir "C:\path\to\outputs"

    python run_visual_vlm_checklist_eval.py `
      --tasks-json "C:\path\to\dataset.json" `
      --output-dir "C:\path\to\outputs" `
      --api-key "YOUR_API_KEY" `
      --model "glm-4.6v"
"""

import argparse
import base64
import datetime
import json
import os
import re
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from openai import OpenAI

try:
    import openpyxl

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import win32com.client

    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

try:
    from PIL import Image

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


DEFAULT_BASE_URL = "https://open.bigmodel.cn/api/paas/v4"
DEFAULT_MODEL = "glm-4.6v"
DEFAULT_TEMPERATURE = 0.1
DEFAULT_MAX_TOKENS = 4096
DEFAULT_TIMEOUT = 300
DEFAULT_ACC_THRESHOLD = 0.7
DEFAULT_SLEEP_SECONDS = 8.0

_SPREADSHEET_PROG_ID: Optional[str] = None
_SPREADSHEET_PROG_ID_RESOLVED = False


class SafeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime.datetime, datetime.date)):
            return obj.isoformat()
        return super().default(obj)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Evaluate chart outputs with a VLM checklist and compute ACC.",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--tasks-json",
        required=True,
        help="Path to task definitions JSON.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory containing model outputs (.xlsx / .png).",
    )
    parser.add_argument(
        "--report-path",
        default=None,
        help=(
            "Where to save the JSON report.\n"
            "If omitted, defaults to a sibling file named evaluation_report_<output-dir-name>.json"
        ),
    )
    parser.add_argument(
        "--task-ids",
        nargs="*",
        default=None,
        help='Optional explicit task IDs, e.g. --task-ids "Task 1" "Task 2"',
    )
    parser.add_argument(
        "--acc-threshold",
        type=float,
        default=DEFAULT_ACC_THRESHOLD,
        help="ACC threshold. score > threshold => acc=1, else 0. Default: 0.7",
    )
    parser.add_argument(
        "--api-key",
        default=os.getenv("VLM_API_KEY") or os.getenv("OPENAI_API_KEY"),
        help="VLM API key. If omitted, reads VLM_API_KEY or OPENAI_API_KEY from environment.",
    )
    parser.add_argument(
        "--base-url",
        default=os.getenv("VLM_BASE_URL", DEFAULT_BASE_URL),
        help=f"OpenAI-compatible base URL. Default: {DEFAULT_BASE_URL}",
    )
    parser.add_argument(
        "--model",
        default=os.getenv("VLM_MODEL", DEFAULT_MODEL),
        help=f"Vision model name. Default: {DEFAULT_MODEL}",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=DEFAULT_TEMPERATURE,
        help=f"Sampling temperature. Default: {DEFAULT_TEMPERATURE}",
    )
    parser.add_argument(
        "--max-tokens",
        type=int,
        default=DEFAULT_MAX_TOKENS,
        help=f"Max completion tokens. Default: {DEFAULT_MAX_TOKENS}",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"Request timeout in seconds. Default: {DEFAULT_TIMEOUT}",
    )
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=DEFAULT_SLEEP_SECONDS,
        help=f"Sleep between successful VLM tasks. Default: {DEFAULT_SLEEP_SECONDS}",
    )
    parser.add_argument(
        "--export-dir",
        default=None,
        help="Optional directory for exported chart images. Default: <output-dir>/exported_charts_public",
    )

    args = parser.parse_args()
    if not args.api_key:
        parser.error("Provide --api-key or set VLM_API_KEY / OPENAI_API_KEY.")
    return args


def normalize_task(task: Dict[str, Any]) -> Dict[str, Any]:
    normalized = dict(task)
    if not normalized.get("category") and normalized.get("instruction_type"):
        normalized["category"] = normalized["instruction_type"]
    return normalized


def load_tasks(tasks_json_path: str) -> List[Dict[str, Any]]:
    with open(tasks_json_path, "r", encoding="utf-8-sig") as f:
        tasks = json.load(f)

    if not isinstance(tasks, list):
        raise ValueError(f"Expected a list of tasks in {tasks_json_path}")

    return [normalize_task(task) for task in tasks]


def resolve_report_path(report_path: Optional[str], output_dir: str) -> str:
    if report_path:
        return report_path
    output_dir_path = Path(output_dir)
    return str(output_dir_path.parent / f"evaluation_report_{output_dir_path.name}.json")


def compute_acc_from_score(score: float, threshold: float) -> int:
    return 1 if score > threshold else 0


def extract_effective_score(result: Dict[str, Any]) -> float:
    if result.get("status") != "success":
        return 0.0

    if "score" in result:
        try:
            return float(result.get("score", 0.0))
        except (TypeError, ValueError):
            return 0.0

    if "checklist" in result:
        try:
            return float(result.get("checklist", {}).get("score", 0.0))
        except (TypeError, ValueError):
            return 0.0

    if result.get("eval_method") == "combo":
        try:
            return float(result.get("combo_score", 0.0))
        except (TypeError, ValueError):
            return 0.0

    try:
        return float(result.get("vlm", {}).get("score", 0.0))
    except (TypeError, ValueError):
        return 0.0


def normalize_loaded_result(result: Dict[str, Any], acc_threshold: float) -> Dict[str, Any]:
    normalized = dict(result)
    score = extract_effective_score(normalized)
    normalized["score"] = round(score, 4)
    normalized["acc"] = compute_acc_from_score(score, acc_threshold)

    if "checklist" not in normalized and "vlm" in normalized:
        vlm_block = normalized.get("vlm") or {}
        normalized["checklist"] = {
            "score": round(score, 4),
            "pass": vlm_block.get("pass", 0),
            "fail": vlm_block.get("fail", 0),
            "total": vlm_block.get("total", 0),
            "details": vlm_block.get("details", []),
        }
    return normalized


def detect_spreadsheet_app() -> Optional[str]:
    global _SPREADSHEET_PROG_ID, _SPREADSHEET_PROG_ID_RESOLVED

    if _SPREADSHEET_PROG_ID_RESOLVED:
        return _SPREADSHEET_PROG_ID

    _SPREADSHEET_PROG_ID_RESOLVED = True
    if not WIN32COM_AVAILABLE:
        return None

    for prog_id in ("Excel.Application", "Ket.Application"):
        try:
            dispatch = getattr(win32com.client, "DispatchEx", win32com.client.Dispatch)
            app = dispatch(prog_id)
            app.Quit()
            _SPREADSHEET_PROG_ID = prog_id
            print(f"[Info] Detected spreadsheet COM: {prog_id}")
            return _SPREADSHEET_PROG_ID
        except Exception:
            continue

    return None


class ChartImageExporter:
    def __init__(self, output_dir: Optional[str] = None):
        self.output_dir = output_dir
        self._app = None

    def _get_app(self):
        prog_id = detect_spreadsheet_app()
        if not prog_id:
            raise RuntimeError("No spreadsheet COM server found (need MS Excel or WPS Office).")

        if self._app is None:
            dispatch = getattr(win32com.client, "DispatchEx", win32com.client.Dispatch)
            self._app = dispatch(prog_id)
            self._app.Visible = False
            self._app.DisplayAlerts = False
        return self._app

    def _wait_for_stable_file(
        self,
        path: str,
        timeout: float = 5.0,
        stability_checks: int = 3,
        interval: float = 0.2,
    ) -> bool:
        start_time = time.time()
        prev_size = -1
        stable_count = 0

        while time.time() - start_time < timeout:
            if not os.path.exists(path):
                time.sleep(interval)
                continue

            try:
                curr_size = os.path.getsize(path)
            except OSError:
                time.sleep(interval)
                continue

            if curr_size > 0 and curr_size == prev_size:
                stable_count += 1
                if stable_count >= stability_checks:
                    return True
            else:
                stable_count = 0

            prev_size = curr_size
            time.sleep(interval)

        return False

    def _export_chart_with_retry(self, chart_obj, output_path: str, max_retries: int = 3) -> bool:
        for attempt in range(max_retries):
            try:
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except OSError:
                        pass

                chart_obj.Chart.Export(output_path, "PNG")
                if self._wait_for_stable_file(output_path, timeout=3.0):
                    return True

                if attempt < max_retries - 1:
                    wait_time = 0.5 * (attempt + 1)
                    print(
                        f"    [Retry] Export attempt {attempt + 1} failed, "
                        f"waiting {wait_time}s before retry..."
                    )
                    time.sleep(wait_time)
            except Exception as exc:
                if attempt < max_retries - 1:
                    print(f"    [Retry] Export exception: {exc}, retrying...")
                    time.sleep(0.5)
                else:
                    print(f"    [Error] Export failed after {max_retries} attempts: {exc}")

        return False

    def export_charts(self, xlsx_path: str, prefix: str = "chart") -> List[str]:
        if not os.path.exists(xlsx_path):
            print(f"[Warn] File not found: {xlsx_path}")
            return []

        if not detect_spreadsheet_app():
            print("[Warn] No spreadsheet COM server available, skipping chart export.")
            return []

        app = self._get_app()
        workbook = None
        exported_paths: List[str] = []

        try:
            workbook = app.Workbooks.Open(os.path.abspath(xlsx_path))
            out_dir = os.path.abspath(self.output_dir or os.path.dirname(xlsx_path))
            os.makedirs(out_dir, exist_ok=True)

            chart_index = 0
            for sheet_index in range(1, workbook.Sheets.Count + 1):
                sheet = workbook.Sheets(sheet_index)
                try:
                    chart_objects = sheet.ChartObjects()
                    count = chart_objects.Count
                except Exception:
                    continue

                for chart_obj_index in range(1, count + 1):
                    try:
                        chart_obj = chart_objects.Item(chart_obj_index)
                        chart_index += 1
                        output_filename = f"{prefix}_{chart_index}.png"
                        output_path = os.path.join(out_dir, output_filename)
                        if self._export_chart_with_retry(chart_obj, output_path):
                            exported_paths.append(output_path)
                            file_size = os.path.getsize(output_path)
                            print(
                                f"  [Export] ChartObject {chart_index} -> "
                                f"{output_filename} ({file_size} bytes)"
                            )
                        else:
                            print(f"  [Warn] Export failed after retries: {output_filename}")
                    except Exception as exc:
                        print(f"  [Warn] Failed to export ChartObject {chart_obj_index}: {exc}")

            return exported_paths
        except Exception as exc:
            if exported_paths:
                print(f"[Warn] Chart export interrupted after partial success for {xlsx_path}: {exc}")
                return exported_paths
            print(f"[Error] Failed to open {xlsx_path}: {exc}")
            return []
        finally:
            if workbook is not None:
                try:
                    workbook.Close(False)
                except Exception:
                    pass

    def combine_images(self, image_paths: List[str], output_path: str) -> Optional[str]:
        if not PIL_AVAILABLE:
            print("[Warn] Pillow not available, cannot stitch images.")
            return None

        if not image_paths:
            return None

        images = []
        try:
            for image_path in image_paths:
                images.append(Image.open(image_path).convert("RGB"))

            total_width = max(image.width for image in images)
            total_height = sum(image.height for image in images)
            combined_image = Image.new("RGB", (total_width, total_height), (255, 255, 255))

            y_offset = 0
            for image in images:
                x_offset = (total_width - image.width) // 2
                combined_image.paste(image, (x_offset, y_offset))
                y_offset += image.height

            combined_image.save(output_path)
            print(f"  [Stitch] Combined {len(images)} charts -> {os.path.basename(output_path)}")
            return output_path
        except Exception as exc:
            print(f"[Error] Failed to stitch images: {exc}")
            return None
        finally:
            for image in images:
                try:
                    image.close()
                except Exception:
                    pass

    def export_all_charts_stitched(self, xlsx_path: str, output_path: Optional[str] = None) -> Optional[str]:
        prefix = f"{Path(xlsx_path).stem}_temp_chart"
        exported_paths = self.export_charts(xlsx_path, prefix=prefix)
        if not exported_paths:
            return None

        if len(exported_paths) == 1:
            single_path = exported_paths[0]
            if output_path:
                try:
                    os.replace(single_path, output_path)
                except OSError:
                    import shutil

                    shutil.copy2(single_path, output_path)
                    os.remove(single_path)
                return output_path
            return single_path

        if output_path is None:
            out_dir = self.output_dir or os.path.dirname(xlsx_path)
            output_path = os.path.join(out_dir, f"{Path(xlsx_path).stem}_stitched_charts.png")

        stitched_path = self.combine_images(exported_paths, output_path)
        for exported_path in exported_paths:
            try:
                if os.path.exists(exported_path):
                    os.remove(exported_path)
            except OSError:
                pass
        return stitched_path

    def close(self):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def __del__(self):
        self.close()


class VlmJudge:
    def __init__(
        self,
        api_key: str,
        base_url: str,
        model: str,
        temperature: float = DEFAULT_TEMPERATURE,
        max_tokens: int = DEFAULT_MAX_TOKENS,
        timeout: int = DEFAULT_TIMEOUT,
    ):
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        self.timeout = timeout

    def _encode_image(self, image_path: str) -> str:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")

    def _extract_response_text(self, content: Any) -> str:
        if isinstance(content, str):
            return content

        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict):
                    text = item.get("text")
                    if text:
                        parts.append(text)
                else:
                    text = getattr(item, "text", None)
                    if text:
                        parts.append(text)
            return "\n".join(parts)

        return "" if content is None else str(content)

    def evaluate(self, image_input: Union[str, List[str]], checklist: List[Dict[str, Any]]) -> Dict[str, Any]:
        image_paths = [image_input] if isinstance(image_input, str) else list(image_input)
        for path in image_paths:
            if not os.path.exists(path):
                return {"error": f"Image not found: {path}", "score": 0.0}

        checklist_text = ""
        for item in checklist:
            checklist_text += f"{item['id']}. {item['assertion']}\n"

        is_multi_chart = len(image_paths) > 1
        if is_multi_chart:
            scope_line = "You are a strict ChartQA Judge for MULTIPLE CHARTS."
            task_line = "Verify if the chart images meet the CHECKLIST requirements."
            prompt_text = (
                f"Evaluate these {len(image_paths)} charts against the following checklist:\n\n"
                f"{checklist_text}"
            )
        else:
            scope_line = "You are a strict ChartQA Judge."
            task_line = "Verify if the chart image meets the CHECKLIST requirements."
            prompt_text = f"Evaluate this chart against the following checklist:\n\n{checklist_text}"

        system_prompt = f"""{scope_line}
{task_line}

RULES:
- Judge each checklist item as PASS or FAIL based ONLY on visual evidence.
- If a criterion says "For chart 'Title': ...", use only that chart.
- Keep FAIL reason concise (<= 8 words). Do not provide PASS reasons.

OUTPUT FORMAT (JSON only, no markdown):
{{
  "pass_ids": [1, 2, 5],
  "fail": [{{"id": 3, "reason": "legend label missing"}}, {{"id": 4, "reason": "wrong chart type"}}]
}}
Only include IDs from the checklist.
"""

        user_content = [{"type": "text", "text": prompt_text}]
        for path in image_paths:
            base64_image = self._encode_image(path)
            user_content.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{base64_image}",
                    },
                }
            )

        try:
            response = self._call_with_retry(system_prompt, user_content)
            content = self._extract_response_text(response.choices[0].message.content)
            return self._parse_evaluation(content, checklist)
        except Exception as exc:
            print(f"[!] VLM API Error (after retries): {exc}")
            return {"error": str(exc), "score": 0.0}

    def _call_with_retry(
        self,
        system_prompt: str,
        user_content: List[Dict[str, Any]],
        max_retries: int = 5,
        base_delay: float = 10.0,
    ):
        last_error = None
        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_content},
                    ],
                    temperature=self.temperature,
                    max_tokens=self.max_tokens,
                    timeout=self.timeout,
                )

                content = self._extract_response_text(response.choices[0].message.content)
                if not content.strip():
                    raise RuntimeError("Empty VLM response")
                return response
            except Exception as exc:
                last_error = exc
                error_text = str(exc).lower()
                retryable = any(
                    token in error_text
                    for token in (
                        "429",
                        "500",
                        "502",
                        "503",
                        "504",
                        "50507",
                        "rate",
                        "timeout",
                        "timed out",
                        "connection",
                        "empty vlm response",
                        "empty response",
                    )
                )
                if not retryable or attempt == max_retries - 1:
                    raise

                delay = base_delay * (2**attempt)
                print(
                    f"    [Retry {attempt + 1}/{max_retries}] API error, waiting {delay:.0f}s... "
                    f"({str(exc)[:80]})"
                )
                time.sleep(delay)

        raise last_error

    def _parse_evaluation(self, content: str, checklist: List[Dict[str, Any]]) -> Dict[str, Any]:
        if not content or not content.strip():
            return {"error": "Empty VLM response", "score": 0.0}

        try:
            clean_content = content.strip()
            clean_content = clean_content.replace("<|begin_of_box|>", "").replace("<|end_of_box|>", "")

            if "```json" in clean_content:
                clean_content = clean_content.split("```json", 1)[1].split("```", 1)[0]
            elif "```" in clean_content:
                clean_content = clean_content.split("```", 1)[1].split("```", 1)[0]

            clean_content = clean_content.strip()
            start = clean_content.find("{")
            if start == -1:
                return {"error": "No JSON found in response", "score": 0.0}

            end = clean_content.rfind("}")
            json_str = clean_content[start : end + 1] if end > start else clean_content[start:]

            try:
                data = json.loads(json_str)
            except json.JSONDecodeError:
                data = self._recover_truncated_json(json_str)

            if data is None:
                print(f"[!] JSON Parse Error, Content: {content[:500]}")
                return {"error": "Failed to parse VLM response", "score": 0.0}

            parsed_results, total_items, passed_items = self._normalize_results(data, checklist)
            total_expected = len(checklist)
            final_score = passed_items / total_expected if total_expected > 0 else 0.0

            if total_items < total_expected:
                print(f"    [Warn] VLM returned {total_items}/{total_expected} items (truncated?)")

            return {
                "score": final_score,
                "details": parsed_results,
                "total_expected": total_expected,
                "total_returned": total_items,
            }
        except Exception as exc:
            print(f"[!] JSON Parse Error: {exc}, Content: {content[:500]}")
            return {"error": "Failed to parse VLM response", "score": 0.0}

    def _normalize_results(self, data: Dict[str, Any], checklist: List[Dict[str, Any]]):
        assertion_ids = {item["id"] for item in checklist}
        assertion_map = {item["id"]: item.get("assertion", "") for item in checklist}
        parsed_by_id: Dict[Any, Dict[str, Any]] = {}

        results = data.get("results", [])
        if isinstance(results, list):
            for res in results:
                if not isinstance(res, dict):
                    continue
                item_id = res.get("id")
                status = str(res.get("result", "FAIL")).upper()
                if status not in ("PASS", "FAIL"):
                    status = "FAIL"
                reason = res.get("reason", "")
                if item_id in assertion_ids:
                    parsed_by_id[item_id] = {
                        "id": item_id,
                        "assertion": assertion_map.get(item_id, ""),
                        "result": status,
                        "reason": reason,
                    }

        pass_ids = data.get("pass_ids", [])
        if isinstance(pass_ids, list):
            for pid in pass_ids:
                try:
                    item_id = int(pid)
                except Exception:
                    item_id = pid
                if item_id in assertion_ids:
                    parsed_by_id[item_id] = {
                        "id": item_id,
                        "assertion": assertion_map.get(item_id, ""),
                        "result": "PASS",
                        "reason": "",
                    }

        fail_items = data.get("fail", [])
        if isinstance(fail_items, list):
            for fail_item in fail_items:
                item_id = None
                reason = ""
                if isinstance(fail_item, dict):
                    item_id = fail_item.get("id")
                    reason = fail_item.get("reason", "")
                else:
                    try:
                        item_id = int(fail_item)
                    except Exception:
                        item_id = fail_item
                if item_id in assertion_ids:
                    parsed_by_id[item_id] = {
                        "id": item_id,
                        "assertion": assertion_map.get(item_id, ""),
                        "result": "FAIL",
                        "reason": reason,
                    }

        parsed_results = [parsed_by_id[key] for key in sorted(parsed_by_id.keys())]
        total_items = len(parsed_results)
        passed_items = sum(1 for result in parsed_results if result.get("result") == "PASS")
        return parsed_results, total_items, passed_items

    def _recover_truncated_json(self, json_str: str) -> Optional[Dict[str, Any]]:
        legacy_pattern = (
            r'\{"id"\s*:\s*(\d+)\s*,\s*"result"\s*:\s*"(PASS|FAIL)"\s*,\s*"reason"\s*:\s*"([^"]*?)"\s*\}'
        )
        matches = re.findall(legacy_pattern, json_str, re.IGNORECASE)
        if matches:
            results = []
            for item_id, result, reason in matches:
                results.append(
                    {
                        "id": int(item_id),
                        "result": result.upper(),
                        "reason": reason,
                    }
                )
            print(f"    [Recovery] Recovered {len(results)} results from truncated JSON")
            return {"results": results}

        pass_ids: List[int] = []
        fail_items: List[Dict[str, Any]] = []

        pass_match = re.search(r'"pass_ids"\s*:\s*\[([^\]]*)', json_str, re.IGNORECASE | re.DOTALL)
        if pass_match:
            pass_ids = [int(x) for x in re.findall(r"\d+", pass_match.group(1))]

        fail_pattern = r'\{"id"\s*:\s*(\d+)\s*,\s*"reason"\s*:\s*"([^"]*?)"'
        fail_matches = re.findall(fail_pattern, json_str, re.IGNORECASE)
        for item_id, reason in fail_matches:
            fail_items.append({"id": int(item_id), "reason": reason})

        if pass_ids or fail_items:
            print(
                f"    [Recovery] Recovered compact format: "
                f"pass={len(pass_ids)}, fail={len(fail_items)}"
            )
            return {"pass_ids": pass_ids, "fail": fail_items}

        return None


@dataclass
class TaskEvaluationResult:
    task_id: str
    category: str = ""
    status: str = "pending"
    error: Optional[str] = None
    eval_method: str = "vlm_checklist"
    num_criteria: int = 0
    score: float = 0.0
    acc: int = 0
    image_source: str = ""
    num_charts: int = 0
    chart_titles: List[str] = field(default_factory=list)
    output_file: str = ""
    checklist_pass: int = 0
    checklist_fail: int = 0
    checklist_total: int = 0
    checklist_details: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "task_id": self.task_id,
            "category": self.category,
            "status": self.status,
            "error": self.error,
            "eval_method": self.eval_method,
            "num_criteria": self.num_criteria,
            "score": round(self.score, 4),
            "acc": self.acc,
            "image_source": self.image_source,
            "num_charts": self.num_charts,
            "chart_titles": self.chart_titles,
            "output_file": self.output_file,
            "checklist": {
                "score": round(self.score, 4),
                "pass": self.checklist_pass,
                "fail": self.checklist_fail,
                "total": self.checklist_total,
                "details": self.checklist_details,
            },
        }


class VisualChecklistEvaluator:
    def __init__(
        self,
        output_dir: str,
        judge: VlmJudge,
        acc_threshold: float = DEFAULT_ACC_THRESHOLD,
        export_dir: Optional[str] = None,
    ):
        self.output_dir = output_dir
        self.judge = judge
        self.acc_threshold = acc_threshold
        self.export_dir = os.path.abspath(export_dir or os.path.join(output_dir, "exported_charts_public"))
        self._chart_exporter: Optional[ChartImageExporter] = None

    @property
    def chart_exporter(self) -> ChartImageExporter:
        if self._chart_exporter is None:
            self._chart_exporter = ChartImageExporter(self.export_dir)
        return self._chart_exporter

    def close(self):
        if self._chart_exporter is not None:
            self._chart_exporter.close()

    def _find_output_file(self, task_id: str, ext: str) -> Optional[str]:
        task_id_underscored = task_id.replace(" ", "_")
        candidates = [
            os.path.join(self.output_dir, f"1_{task_id}_output{ext}"),
            os.path.join(self.output_dir, f"{task_id}_output{ext}"),
            os.path.join(self.output_dir, f"{task_id_underscored}_output{ext}"),
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        return None

    def _criteria_to_checklist(self, criteria: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        checklist = []
        for criterion in criteria:
            checklist.append(
                {
                    "id": criterion["id"],
                    "assertion": criterion.get("instruction", criterion.get("description", "")),
                    "weight": 1.0,
                    "rubric_category": criterion.get("rubric_category", ""),
                }
            )
        return checklist

    def _extract_chart_titles(self, criteria: List[Dict[str, Any]]) -> List[str]:
        titles = set()
        for criterion in criteria:
            chart = criterion.get("chart")
            if chart:
                titles.add(chart)
        return sorted(titles)

    def _count_charts_in_xlsx(self, xlsx_path: str) -> int:
        if not OPENPYXL_AVAILABLE:
            return -1

        workbook = None
        try:
            workbook = openpyxl.load_workbook(xlsx_path, data_only=True)
            return sum(len(getattr(sheet, "_charts", []) or []) for sheet in workbook.worksheets)
        except Exception:
            return -1
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass

    def _normalize_chart_title(self, title: str) -> str:
        if not title:
            return ""
        return " ".join(str(title).strip().lower().split())

    def _resolve_title_formula(self, formula: str, workbook) -> str:
        if not formula:
            return ""

        formula_text = str(formula).strip()
        if len(formula_text) >= 2 and formula_text[0] == '"' and formula_text[-1] == '"':
            return formula_text[1:-1]
        if len(formula_text) >= 3 and formula_text.startswith('="') and formula_text.endswith('"'):
            return formula_text[2:-1]

        match = re.match(r"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)$", formula_text)
        if not match:
            return ""

        sheet_name = match.group(1) or match.group(2)
        cell_ref = f"{match.group(3)}{match.group(4)}"
        try:
            worksheet = workbook[sheet_name]
            value = worksheet[cell_ref].value
            return str(value).strip() if value is not None else ""
        except Exception:
            return ""

    def _extract_openpyxl_chart_title(self, chart, workbook) -> str:
        try:
            title_obj = getattr(chart, "title", None)
            if title_obj is None:
                return ""

            tx = getattr(title_obj, "tx", None)
            if tx is None:
                return ""

            rich = getattr(tx, "rich", None)
            if rich and getattr(rich, "p", None):
                parts = []
                for paragraph in rich.p:
                    runs = getattr(paragraph, "r", None) or []
                    for run in runs:
                        text = getattr(run, "t", None)
                        if text:
                            parts.append(text)
                if parts:
                    return "".join(parts).strip()

            str_ref = getattr(tx, "strRef", None)
            if str_ref and getattr(str_ref, "f", None):
                return self._resolve_title_formula(str_ref.f, workbook)
        except Exception:
            return ""

        return ""

    def _extract_chart_titles_in_order(self, xlsx_path: str) -> List[str]:
        if not OPENPYXL_AVAILABLE:
            return []

        workbook = None
        titles = []
        try:
            workbook = openpyxl.load_workbook(xlsx_path, data_only=True)
            for worksheet in workbook.worksheets:
                for chart in getattr(worksheet, "_charts", []) or []:
                    titles.append(self._extract_openpyxl_chart_title(chart, workbook))
        except Exception:
            return []
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass
        return titles

    def _build_chart_image_map(self, xlsx_path: str, image_paths: List[str]) -> Dict[str, str]:
        titles = self._extract_chart_titles_in_order(xlsx_path)
        if not titles:
            return {}

        if len(titles) != len(image_paths):
            print(
                f"  [Warn] Chart title count {len(titles)} != exported image count "
                f"{len(image_paths)}, building partial mapping"
            )

        mapping: Dict[str, str] = {}
        for title, image_path in zip(titles, image_paths):
            key = self._normalize_chart_title(title)
            if key and key not in mapping:
                mapping[key] = image_path
        return mapping

    def _evaluate_with_missing_repair(
        self,
        image_input: Union[str, List[str]],
        checklist: List[Dict[str, Any]],
        max_rounds: int = 2,
    ) -> Dict[str, Any]:
        total_expected = len(checklist)
        if total_expected == 0:
            return {"score": 0.0, "details": [], "total_expected": 0, "total_returned": 0}

        assertion_map = {item["id"]: item.get("assertion", "") for item in checklist}
        merged: Dict[Any, Dict[str, Any]] = {}
        remaining = list(checklist)
        last_error = None

        for round_idx in range(max_rounds + 1):
            if not remaining:
                break

            partial = self.judge.evaluate(image_input, remaining)
            if "error" in partial:
                last_error = partial.get("error", "Unknown error")
                if not merged:
                    return partial
                print(f"    [Warn] Repair round {round_idx + 1} failed: {last_error}")
                break

            remaining_ids = {item["id"] for item in remaining}
            for detail in partial.get("details", []):
                item_id = detail.get("id")
                if item_id in remaining_ids:
                    normalized = dict(detail)
                    if not normalized.get("assertion"):
                        normalized["assertion"] = assertion_map.get(item_id, "")
                    merged[item_id] = normalized

            missing_ids = [item["id"] for item in remaining if item["id"] not in merged]
            if not missing_ids:
                break

            if round_idx < max_rounds:
                print(f"    [Repair] Missing {len(missing_ids)} items, retrying missing IDs only...")
                remaining = [item for item in remaining if item["id"] in missing_ids]

        ordered_ids = [item["id"] for item in checklist]
        details = [merged[item_id] for item_id in ordered_ids if item_id in merged]
        passed = sum(1 for item_id in ordered_ids if merged.get(item_id, {}).get("result") == "PASS")
        total_returned = len(details)

        if total_returned < total_expected:
            print(f"    [Warn] After repair, returned {total_returned}/{total_expected} items.")
            if last_error:
                print(f"    [Warn] Last repair error: {last_error}")

        return {
            "score": passed / total_expected if total_expected > 0 else 0.0,
            "details": details,
            "total_expected": total_expected,
            "total_returned": total_returned,
        }

    def _evaluate_multi_chart_with_routing(
        self,
        image_paths: List[str],
        checklist: List[Dict[str, Any]],
        criteria: List[Dict[str, Any]],
        chart_image_map: Dict[str, str],
    ) -> Dict[str, Any]:
        criteria_by_id = {criterion.get("id"): criterion for criterion in criteria}
        routed_groups: Dict[str, Dict[str, Any]] = {}
        global_items = []

        for item in checklist:
            criterion = criteria_by_id.get(item.get("id"), {})
            chart_title = criterion.get("chart")
            chart_key = self._normalize_chart_title(chart_title)

            if chart_key and chart_key in chart_image_map:
                if chart_key not in routed_groups:
                    routed_groups[chart_key] = {
                        "title": chart_title,
                        "image": chart_image_map[chart_key],
                        "items": [],
                    }
                routed_groups[chart_key]["items"].append(item)
            else:
                global_items.append(item)

        merged: Dict[Any, Dict[str, Any]] = {}
        ordered_ids = [item["id"] for item in checklist]

        for group in routed_groups.values():
            print(f"  [Route] Chart '{group['title']}': {len(group['items'])} criteria -> single image")
            partial = self._evaluate_with_missing_repair(group["image"], group["items"])
            if "error" in partial:
                return partial
            for detail in partial.get("details", []):
                merged[detail["id"]] = detail

        if global_items:
            print(
                f"  [Route] Global/Unmatched: {len(global_items)} criteria -> "
                f"all {len(image_paths)} images"
            )
            partial = self._evaluate_with_missing_repair(image_paths, global_items)
            if "error" in partial:
                return partial
            for detail in partial.get("details", []):
                merged[detail["id"]] = detail

        details = [merged[item_id] for item_id in ordered_ids if item_id in merged]
        passed = sum(1 for item_id in ordered_ids if merged.get(item_id, {}).get("result") == "PASS")
        total_expected = len(checklist)
        total_returned = len(details)

        if total_returned < total_expected:
            print(f"  [Warn] Routed evaluation returned {total_returned}/{total_expected} items.")

        return {
            "score": passed / total_expected if total_expected > 0 else 0.0,
            "details": details,
            "total_expected": total_expected,
            "total_returned": total_returned,
        }

    def evaluate_task(self, task: Dict[str, Any]) -> TaskEvaluationResult:
        task_id = task["id"]
        category = task.get("category", "")
        criteria = task.get("criteria", []) or []

        result = TaskEvaluationResult(
            task_id=task_id,
            category=category,
            num_criteria=len(criteria),
        )

        if not criteria:
            result.status = "skipped"
            result.error = "No checklist criteria provided for VLM evaluation"
            return result

        xlsx_path = self._find_output_file(task_id, ".xlsx")
        png_path = self._find_output_file(task_id, ".png")

        if not xlsx_path and not png_path:
            result.status = "error"
            result.error = f"No output file found for {task_id}"
            return result

        image_input: Optional[Union[str, List[str]]] = None
        chart_image_map: Dict[str, str] = {}

        if png_path:
            image_input = png_path
            result.output_file = os.path.basename(png_path)
            result.image_source = "png_output"
            result.num_charts = 1
            print(f"  [Info] Using PNG output: {os.path.basename(png_path)}")
        elif xlsx_path:
            result.output_file = os.path.basename(xlsx_path)
            print(f"  [Info] Exporting charts from XLSX: {os.path.basename(xlsx_path)}")
            try:
                os.makedirs(self.export_dir, exist_ok=True)
                expected_chart_count = self._count_charts_in_xlsx(xlsx_path)
                exported_paths = self.chart_exporter.export_charts(
                    xlsx_path,
                    prefix=f"1_{task_id}_exported_chart",
                )
                exported_paths = [
                    path
                    for path in exported_paths
                    if os.path.exists(path) and os.path.getsize(path) > 0
                ]

                if not exported_paths:
                    result.status = "error"
                    result.error = "XLSX has no chart objects to export"
                    return result

                exported_count = len(exported_paths)
                result.num_charts = expected_chart_count if expected_chart_count > 0 else exported_count

                if exported_count == 1:
                    image_input = exported_paths[0]
                    result.image_source = "xlsx_export"
                    print(f"  [Info] Exported chart image: {os.path.basename(exported_paths[0])}")
                else:
                    image_input = exported_paths
                    result.image_source = "xlsx_export_multi"
                    chart_image_map = self._build_chart_image_map(xlsx_path, exported_paths)
                    print(
                        f"  [Info] Exported {exported_count} chart images "
                        "for multi-image VLM evaluation"
                    )

                if expected_chart_count > 0 and exported_count < expected_chart_count:
                    print(
                        f"  [Warn] Exported {exported_count}/{expected_chart_count} charts. "
                        "Trying stitched fallback..."
                    )
                    stitched_path = self.chart_exporter.export_all_charts_stitched(
                        xlsx_path,
                        os.path.join(self.export_dir, f"1_{task_id}_exported_chart.png"),
                    )
                    if (
                        stitched_path
                        and os.path.exists(stitched_path)
                        and os.path.getsize(stitched_path) > 0
                    ):
                        image_input = stitched_path
                        result.image_source = "stitched_fallback"
                        chart_image_map = {}
                        print(
                            f"  [Info] Using stitched fallback image: "
                            f"{os.path.basename(stitched_path)}"
                        )
            except Exception as exc:
                result.status = "error"
                result.error = f"Chart export failed: {exc}"
                return result

        if image_input is None:
            result.status = "error"
            result.error = "Could not obtain chart image"
            return result

        checklist = self._criteria_to_checklist(criteria)
        result.chart_titles = self._extract_chart_titles(criteria)
        result.checklist_total = len(checklist)
        print(f"  [Info] Checklist: {len(checklist)} items")
        print(f"  [VLM] Evaluating with {len(checklist)} criteria...")

        try:
            if isinstance(image_input, list) and len(image_input) > 1:
                vlm_result = self._evaluate_multi_chart_with_routing(
                    image_paths=image_input,
                    checklist=checklist,
                    criteria=criteria,
                    chart_image_map=chart_image_map,
                )
            else:
                vlm_result = self._evaluate_with_missing_repair(image_input, checklist)
        except Exception as exc:
            result.status = "error"
            result.error = f"VLM evaluation failed: {exc}"
            return result

        if "error" in vlm_result:
            result.status = "error"
            result.error = f"VLM error: {vlm_result['error']}"
            return result

        score = float(vlm_result.get("score", 0.0))
        details = vlm_result.get("details", []) or []
        passed = sum(1 for detail in details if detail.get("result") == "PASS")

        result.status = "success"
        result.score = score
        result.acc = compute_acc_from_score(score, self.acc_threshold)
        result.checklist_pass = passed
        result.checklist_fail = result.checklist_total - passed
        result.checklist_details = details
        print(
            f"  [Result] Score: {result.score:.2%} | "
            f"ACC: {result.acc} | PASS: {passed}/{result.checklist_total}"
        )
        return result


def build_summary(
    tasks_to_eval: List[Dict[str, Any]],
    completed_results: Dict[str, Dict[str, Any]],
    acc_threshold: float,
) -> Dict[str, Any]:
    ordered_results = [
        completed_results[task["id"]]
        for task in tasks_to_eval
        if task["id"] in completed_results
    ]

    total_tasks = len(tasks_to_eval)
    completed_count = len(ordered_results)
    pending = total_tasks - completed_count
    evaluated = sum(1 for result in ordered_results if result.get("status") == "success")
    skipped = sum(1 for result in ordered_results if result.get("status") == "skipped")
    errors = sum(1 for result in ordered_results if result.get("status") == "error")
    tasks_with_criteria = sum(1 for task in tasks_to_eval if task.get("criteria"))

    success_scores = [extract_effective_score(result) for result in ordered_results if result.get("status") == "success"]
    total_score_all = sum(extract_effective_score(result) for result in ordered_results)
    acc_count = sum(compute_acc_from_score(extract_effective_score(result), acc_threshold) for result in ordered_results)

    return {
        "total_tasks": total_tasks,
        "completed": completed_count,
        "pending": pending,
        "evaluated": evaluated,
        "skipped": skipped,
        "errors": errors,
        "tasks_with_checklist": tasks_with_criteria,
        "tasks_without_checklist": total_tasks - tasks_with_criteria,
        "success_only_avg_score": round(sum(success_scores) / len(success_scores), 4) if success_scores else 0.0,
        "all_task_avg_score": round(total_score_all / total_tasks, 4) if total_tasks > 0 else 0.0,
        "acc": round(acc_count / total_tasks, 4) if total_tasks > 0 else 0.0,
        "acc_tasks": acc_count,
        "acc_total": total_tasks,
        "acc_threshold": acc_threshold,
    }


def save_report(
    tasks_to_eval: List[Dict[str, Any]],
    completed_results: Dict[str, Dict[str, Any]],
    report_path: str,
    acc_threshold: float,
    model: str,
    base_url: str,
):
    ordered_results = [
        completed_results[task["id"]]
        for task in tasks_to_eval
        if task["id"] in completed_results
    ]
    report = {
        "meta": {
            "generated_at": datetime.datetime.now(),
            "eval_method": "vlm_checklist_only",
            "model": model,
            "base_url": base_url,
            "note": "Missing outputs, skipped tasks, and unfinished tasks are counted as score=0 and acc=0 in summary metrics.",
        },
        "summary": build_summary(tasks_to_eval, completed_results, acc_threshold),
        "results": ordered_results,
    }

    os.makedirs(os.path.dirname(report_path) or ".", exist_ok=True)
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False, cls=SafeEncoder)


def print_summary(summary: Dict[str, Any]):
    print("\n" + "=" * 70)
    print("EVALUATION SUMMARY")
    print("=" * 70)
    print(f"  Total tasks:             {summary['total_tasks']}")
    print(f"  Completed:               {summary['completed']}")
    print(f"  Pending:                 {summary['pending']}")
    print(f"  Evaluated:               {summary['evaluated']}")
    print(f"  Skipped:                 {summary['skipped']}")
    print(f"  Errors:                  {summary['errors']}")
    print(f"  With checklist:          {summary['tasks_with_checklist']}")
    print(f"  Without checklist:       {summary['tasks_without_checklist']}")
    print(f"  Success-only avg score:  {summary['success_only_avg_score']:.2%}")
    print(f"  All-task avg score:      {summary['all_task_avg_score']:.2%}")
    print(
        f"  ACC (score > {summary['acc_threshold']}): "
        f"{summary['acc_tasks']}/{summary['acc_total']} = {summary['acc']:.2%}"
    )
    print("=" * 70)


def main():
    args = parse_args()
    report_path = resolve_report_path(args.report_path, args.output_dir)
    all_tasks = load_tasks(args.tasks_json)

    if args.task_ids:
        requested_ids = set(args.task_ids)
        selected_tasks = [task for task in all_tasks if task["id"] in requested_ids]
        found_ids = {task["id"] for task in selected_tasks}
        missing_ids = [task_id for task_id in args.task_ids if task_id not in found_ids]
        if missing_ids:
            print(f"[Warn] Task IDs not found in tasks JSON: {missing_ids}")
    else:
        selected_tasks = all_tasks

    if not selected_tasks:
        print("No tasks selected. Nothing to evaluate.")
        return 1

    all_ids = [task["id"] for task in selected_tasks]
    selected_id_set = set(all_ids)
    task_map = {task["id"]: task for task in selected_tasks}

    print("Evaluation config:")
    print(f"  tasks_json:     {args.tasks_json}")
    print(f"  output_dir:     {args.output_dir}")
    print(f"  report_path:    {report_path}")
    print(f"  task_count:     {len(all_ids)}")
    print(f"  acc_threshold:  {args.acc_threshold}")
    print(f"  model:          {args.model}")
    print(f"  base_url:       {args.base_url}")
    print(f"  export_dir:     {args.export_dir or os.path.join(args.output_dir, 'exported_charts_public')}")

    completed_results: Dict[str, Dict[str, Any]] = {}
    if os.path.exists(report_path):
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
            for result in old_data.get("results", []):
                task_id = result.get("task_id")
                if task_id in selected_id_set:
                    completed_results[task_id] = normalize_loaded_result(result, args.acc_threshold)
            if completed_results:
                print(f"Resuming: {len(completed_results)} tasks already completed")
        except Exception as exc:
            print(f"[Warn] Failed to load existing report, starting fresh: {exc}")

    remaining_ids = [task_id for task_id in all_ids if task_id not in completed_results]
    print(f"Remaining: {len(remaining_ids)} tasks to evaluate")
    print(f"Report: {report_path}\n")

    if not remaining_ids:
        save_report(
            selected_tasks,
            completed_results,
            report_path,
            args.acc_threshold,
            args.model,
            args.base_url,
        )
        print_summary(build_summary(selected_tasks, completed_results, args.acc_threshold))
        print("All selected tasks already evaluated.")
        return 0

    judge = VlmJudge(
        api_key=args.api_key,
        base_url=args.base_url,
        model=args.model,
        temperature=args.temperature,
        max_tokens=args.max_tokens,
        timeout=args.timeout,
    )
    evaluator = VisualChecklistEvaluator(
        output_dir=args.output_dir,
        judge=judge,
        acc_threshold=args.acc_threshold,
        export_dir=args.export_dir,
    )

    start_time = time.time()
    consecutive_exceptions = 0

    try:
        for task_id in remaining_ids:
            task = task_map.get(task_id)
            if task is None:
                print(f"[Warn] Task {task_id} not found, skipping")
                continue

            print("\n" + "=" * 70)
            print(
                f"[{len(completed_results) + 1}/{len(all_ids)}] "
                f"Evaluating {task_id} ({task.get('category', 'N/A')})"
            )
            print("=" * 70)

            try:
                result = evaluator.evaluate_task(task)
                completed_results[task_id] = result.to_dict()
                consecutive_exceptions = 0

                if result.status == "success":
                    print(f"  >> success ({result.score:.2%}, acc={result.acc})")
                elif result.status == "skipped":
                    print(f"  >> skipped ({result.error})")
                else:
                    print(f"  >> error ({result.error})")
            except Exception as exc:
                consecutive_exceptions += 1
                print(f"  [FATAL ERROR] {str(exc)[:200]}")
                traceback.print_exc()
                completed_results[task_id] = {
                    "task_id": task_id,
                    "category": task.get("category", ""),
                    "status": "error",
                    "error": str(exc)[:500],
                    "eval_method": "vlm_checklist",
                    "num_criteria": len(task.get("criteria", []) or []),
                    "score": 0.0,
                    "acc": 0,
                    "image_source": "",
                    "num_charts": 0,
                    "chart_titles": [],
                    "output_file": "",
                    "checklist": {
                        "score": 0.0,
                        "pass": 0,
                        "fail": 0,
                        "total": len(task.get("criteria", []) or []),
                        "details": [],
                    },
                }

                if consecutive_exceptions >= 5:
                    print("\n[ABORT] 5 consecutive exceptions. Progress saved, re-run to resume.")
                    save_report(
                        selected_tasks,
                        completed_results,
                        report_path,
                        args.acc_threshold,
                        args.model,
                        args.base_url,
                    )
                    return 1

                wait_seconds = min(60 * consecutive_exceptions, 300)
                print(f"  Waiting {wait_seconds}s before next task...")
                time.sleep(wait_seconds)

            save_report(
                selected_tasks,
                completed_results,
                report_path,
                args.acc_threshold,
                args.model,
                args.base_url,
            )

            if (
                completed_results[task_id].get("status") == "success"
                and args.sleep_seconds > 0
                and task_id != remaining_ids[-1]
            ):
                time.sleep(args.sleep_seconds)
    finally:
        evaluator.close()

    elapsed = time.time() - start_time
    summary = build_summary(selected_tasks, completed_results, args.acc_threshold)

    print("\n" + "=" * 70)
    print(f"DONE! Total time: {elapsed:.1f}s ({elapsed / 60:.1f} min)")
    print(f"Completed results: {summary['completed']}/{summary['total_tasks']}")
    print(f"Report: {report_path}")
    print_summary(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
