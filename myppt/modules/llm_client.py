# ai_gen.py
# -*- coding: utf-8 -*-
"""
统一封装 LLM 调用
-----------------------------------------------------------------
对外函数
    gen_outline(template_md, prompts)  -> Markdown
    gen_content(outline_md, prompts)   -> Markdown
抛出的唯一异常
    LLMError
"""
from __future__ import annotations
import os, re, sys, time, traceback
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import config
from openai import OpenAI

try:                                        # openai ≥ 1.0
    from openai import APIError, RateLimitError, Timeout
except ImportError:                         # openai 0.27.x ↓
    from openai.error import APIError, RateLimitError, Timeout  # type: ignore
if "Timeout" not in globals():              # noqa: WPS421
    class Timeout(Exception):               # type: ignore
        pass

from .md_utils import fence_md

__all__ = ["gen_outline", "gen_content", "LLMError"]

# ───────── 正则 ─────────
_len_pat  = re.compile(r"<!--\s*len:(\d+)\s*-->")
_lock_pat = re.compile(r"<!--LOCK-->(.*?)<!--/LOCK-->", re.S)

# ───────── 错误封装 ─────────
@dataclass
class LLMError(Exception):
    message: str
    type: str = "LLMError"
    callstack: Optional[str] = None
    def __str__(self): return f"{self.type}: {self.message}"

# ───────── Key / URL ─────────
def _ensure_key(key_name: str = "llm_key") -> str:
    val = os.getenv(key_name.upper()) or config.get(key_name)
    if not val:
        raise LLMError(f"缺少 {key_name}，请在配置中填写", "MissingKey")
    return val
def _ensure_base_url() -> str:
    return os.getenv("LLM_BASE_URL", "https://api.deepseek.com/v1").rstrip("/")

# ───────── 工具函数 ─────────
def _strip_md_fence(text: str) -> str:
    if text.startswith("```"):
        parts = text.split("```", 2)
        if len(parts) >= 3:
            return parts[1].lstrip("markdown").lstrip().lstrip("\n").rstrip()
    return text.strip()
def _split_lines_keep_eol(txt: str) -> List[str]:
    return re.findall(r".*?(?:\n|$)", txt)
def _find_zero_len_lines(md: str) -> List[Tuple[int, str]]:
    res = []
    for idx, ln in enumerate(_split_lines_keep_eol(md)):
        m = _len_pat.search(ln)
        if m and m.group(1) == "0":
            res.append((idx, ln.rstrip("\n")))
    return res

# ───────── ChatCompletion ─────────
def _chat_completion(sys_prompt: str, usr_prompt: str,
                     *, retries: int = 3,
                     model: str = "deepseek-chat",
                     temperature: float = 0.25) -> str:
    client = OpenAI(api_key=_ensure_key("llm_key"),
                    base_url=_ensure_base_url())
    for attempt in range(1, retries + 1):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user",   "content": usr_prompt}],
                temperature=temperature,
                stream=False,
            )
            return _strip_md_fence(resp.choices[0].message.content or "")
        except (RateLimitError, Timeout) as e:
            if attempt >= retries:
                raise LLMError("接口限流/超时，达到最大重试次数",
                               "RateLimit/Timeout",
                               traceback.format_exc()) from e
            backoff = 2 ** attempt
            print(f"[LLM] 第 {attempt}/{retries} 次失败，{backoff}s 后重试 …",
                  file=sys.stderr)
            time.sleep(backoff)
        except APIError as e:
            raise LLMError(str(e), "APIError", traceback.format_exc()) from e
        except Exception as e:
            raise LLMError(str(e), "UnexpectedError", traceback.format_exc()) from e
    raise LLMError("达到最大重试次数仍失败", "MaxRetryError")

# ───────── 内部处理 ─────────
def _lock_zero_len_lines(md: str) -> Tuple[str, List[Tuple[int, str]]]:
    zero = _find_zero_len_lines(md)
    lines = _split_lines_keep_eol(md)
    for idx, origin in zero:
        lines[idx] = f"<!--LOCK-->{origin}<!--/LOCK-->\n"
    return "".join(lines), zero

def _unlock_and_dedup(md: str,
                      zero_lines: List[Tuple[int, str]],
                      *, strip_len_tag: bool = False) -> str:
    # 1. 解锁
    md = _lock_pat.sub(lambda m: m.group(1), md)

    # 2. len=0 占位符仅保留第一次出现
    seen: set[str] = set()
    out_lines: List[str] = []
    for ln in _split_lines_keep_eol(md):
        m = _len_pat.search(ln)
        if m and m.group(1) == "0":            # len=0
            key = _len_pat.sub("", ln).strip()
            if key in seen:
                continue
            seen.add(key)
        out_lines.append(ln)

    # 3. 是否去掉 len 注释
    if strip_len_tag:
        tmp = []
        for ln in out_lines:
            if _len_pat.search(ln):
                m = _len_pat.search(ln)
                # len=0 占位符 → 保留整行（不带注释）
                if m.group(1) == "0":
                    ln = _len_pat.sub("", ln)
                else:
                    ln = _len_pat.sub("", ln)
            tmp.append(ln)
        out_lines = tmp

    return "".join(out_lines).rstrip()

# ───────── outline ─────────
def gen_outline(template_md: str, prompts: Dict) -> str:
    locked, zero = _lock_zero_len_lines(template_md)
    sys_prompt = (
        f"你是一名{prompts['ai_role']}，擅长撰写{prompts['report_type']}的大纲。\n"
        "所有被 <!--LOCK--> 包裹的行无需填充，必须原样保留；"
        "其他占位符请替换为小标题，保持 Markdown 层级与顺序一致。")
    usr_prompt = (f"主题：{prompts['topic']}\n\n"
                  f"{fence_md(locked)}")
    raw = _chat_completion(sys_prompt, usr_prompt)
    outline = _unlock_and_dedup(raw, zero, strip_len_tag=False)

    print("\n================= 生成的大纲 =================\n")
    print(outline, "\n=============================================\n")
    return outline

# ───────── content ─────────
def gen_content(outline_md: str, prompts: Dict) -> str:
    locked, zero = _lock_zero_len_lines(outline_md)
    sys_prompt = (
        "你是一名资深演示稿撰写助手。\n"
        "规则：\n"
        "1. LOCK 包裹行保持不变;\n"
        "2. 其它行尾 <!--len:x--> 为长度提示, 生成 x±20% 字数正文;\n"
        "3. 输出不得含 LOCK 或 <!--len:x--> 注释，但必须保留原占位符行。")
    usr_prompt = (f"主题：{prompts['topic']}\n\n"
                  f"{fence_md(locked)}")
    raw = _chat_completion(sys_prompt, usr_prompt)

    # 最后 strip_len_tag=True → 清理注释，但占位符行仍在
    full = _unlock_and_dedup(raw, zero, strip_len_tag=True)

    print("\n================= 生成的正文 =================\n")
    print(full, "\n=============================================\n")
    return full