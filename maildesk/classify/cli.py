"""Maildesk AI 분류기 — Claude Code CLI 방식 (claude.ai 구독 사용, API 결제 불필요)

llm.py(API 방식)와 '같은 프롬프트, 같은 JSON 형식, 같은 검증'을 쓰고 호출 경로만 다르다.
나중에 API 결제가 가능해지면 --classifier llm 으로 바꾸기만 하면 된다.

공정성·안전 장치:
- 메일 1건마다 새 세션으로 호출하고 저장하지 않는다(--no-session-persistence).
- 빈 임시 폴더에서 실행하고 모든 도구를 끈다(--tools ""). → 정답표 파일 등을 읽을 수 없다.
- 시스템 프롬프트를 통째로 바꾼다(--system-prompt). → 이 프로젝트의 메모리/설정 맥락이 섞이지 않는다.
- 메일 본문은 표준입력으로 넘긴다(명령줄 따옴표 문제 방지).
"""
import json
import os
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path

from llm import SCHEMA, build_system, build_user, failed, postprocess

CLASSIFIER_VERSION = "cli-v1"
TASK = "표준입력으로 받은 메일 1건을 시스템 지시에 따라 분류하고 JSON 스키마에 맞춰 답하세요."


def find_claude():
    exe = shutil.which("claude")
    if exe:
        return exe
    native = Path(os.environ.get("USERPROFILE", "")) / ".local" / "bin" / "claude.exe"
    if native.exists():
        return str(native)
    raise RuntimeError("claude CLI를 찾지 못했습니다. 설치 후 새 터미널에서 'claude --version'을 확인하세요.")


def call_claude(exe, system, task, stdin_text, schema, workdir, effort="medium", model=None, timeout=300):
    """claude -p 를 도구 없이 1회 호출하고 JSON 결과(dict)를 돌려준다. 실패 시 {'_error': ...}."""
    cmd = [exe, "-p", task,
           "--output-format", "json",
           "--json-schema", json.dumps(schema, ensure_ascii=False),
           "--system-prompt", system,
           "--tools", "",
           "--disallowedTools", "mcp__*",
           "--strict-mcp-config",
           "--disable-slash-commands",
           "--no-session-persistence",
           "--max-turns", "3",
           "--effort", effort]
    if model:
        cmd += ["--model", model]
    try:
        proc = subprocess.run(cmd, input=stdin_text, capture_output=True, text=True,
                              encoding="utf-8", cwd=workdir, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {"_error": "CLI 시간 초과"}
    try:
        res = json.loads(proc.stdout)
    except json.JSONDecodeError:
        return {"_error": f"CLI 출력 해석 실패 (종료코드 {proc.returncode}): {(proc.stderr or proc.stdout)[:200]}"}
    if res.get("is_error"):
        res["_error"] = f"CLI 오류: {str(res.get('result'))[:200]}"
    elif res.get("structured_output") is None:
        try:
            res["structured_output"] = json.loads(res.get("result") or "")
        except json.JSONDecodeError:
            res["_error"] = "구조화 출력 없음"
    return res


class CLIClassifier:
    def __init__(self, label_rules, reference_time, effort="medium", model=None, timeout=300):
        self.exe = find_claude()
        self.system = build_system(label_rules, reference_time)
        self.effort = effort
        self.model = model
        self.timeout = timeout
        self.workdir = tempfile.mkdtemp(prefix="maildesk-cli-")  # 빈 폴더
        self.usage = Counter()
        self.served_models = Counter()
        self.est_cost = 0.0

    def classify(self, mail):
        res = call_claude(self.exe, self.system, TASK, build_user(mail), SCHEMA, self.workdir,
                          self.effort, self.model, self.timeout)
        if "result" not in res and "_error" in res:
            return failed(res["_error"])

        self.est_cost += res.get("total_cost_usd") or 0
        u = res.get("usage") or {}
        self.usage["input"] += u.get("input_tokens", 0)
        self.usage["output"] += u.get("output_tokens", 0)
        self.usage["cache_read"] += u.get("cache_read_input_tokens", 0)
        self.usage["cache_write"] += u.get("cache_creation_input_tokens", 0)
        models = list((res.get("modelUsage") or {}).keys())
        for m in models:
            self.served_models[m] += 1

        if "_error" in res:
            return failed(res["_error"])
        return postprocess(res["structured_output"], mail, ",".join(models))

    def cost_usd(self):
        return self.est_cost
