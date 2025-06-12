"""
保存/读取 LLM 与 Unsplash 的 API-Key
配置文件写在当前脚本所在目录：api_keys.txt
"""
import json
from pathlib import Path
from typing import Optional

# ① 保存位置改为“同文件夹”
_CFG_PATH = Path(__file__).resolve().parent / "api_keys.txt"
_DEFAULT  = {"llm_key": "", "unsplash_key": ""}


def _load() -> dict:
    if not _CFG_PATH.exists():
        return _DEFAULT.copy()
    try:
        return json.loads(_CFG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return _DEFAULT.copy()


def _dump(cfg: dict):
    _CFG_PATH.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def get(key_name: str) -> Optional[str]:
    cfg = _load()
    return cfg.get(key_name) or None


def set_(key_name: str, value: str):
    cfg = _load()
    cfg[key_name] = value.strip()
    _dump(cfg)