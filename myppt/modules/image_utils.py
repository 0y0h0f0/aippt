# image_utils.py
# -*- coding: utf-8 -*-
"""
Unsplash 图片搜索工具
-----------------------------------------------------------------
unsplash_search("cat", per_page=5) -> List[bytes]
"""

from __future__ import annotations

import logging
import os
import random
from typing import List, Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry

# ──────────────────── Key 获取：多来源兼容 ────────────────────
def _ensure_key(key_name: str = "unsplash_key") -> str:
    """
    依次尝试：
        1. 环境变量         UNSPLASH_KEY
        2. config.get()     来自 api_keys.txt / config.py
        3. llm_client._ensure_key（若该模块可用）
    均失败则抛 KeyError
    """
    # 1. 环境变量
    env_val = os.getenv(key_name.upper())
    if env_val:
        return env_val

    # 2. config.py
    try:
        import config  # 项目内统一配置模块
        cfg_val = config.get(key_name)  # type: ignore[attr-defined]
        if cfg_val:
            return cfg_val
    except Exception:
        pass

    # 3. llm_client
    try:
        from .ai_gen import _ensure_key as llm_ensure_key  # 与 ai_gen.py 同级
        return llm_ensure_key(key_name)
    except Exception:
        pass

    raise KeyError(f"Unsplash Access-Key 未设置（环境变量 {key_name.upper()} 或 config 中均不存在）")


__all__ = ["unsplash_search", "UnsplashError"]

LOG = logging.getLogger(__name__)


class UnsplashError(RuntimeError):
    """封装对外抛出的统一异常"""
    pass


# ──────────────────── 内部工具：带重试的 session ────────────────────
def _get_session(max_retries: int = 3, backoff: float = 0.5) -> requests.Session:
    session = requests.Session()
    retry_cfg = Retry(
        total=max_retries,
        backoff_factor=backoff,
        status_forcelist=(500, 502, 503, 504),
        allowed_methods=("GET",),
        raise_on_status=False,
    )
    session.mount("https://", HTTPAdapter(max_retries=retry_cfg))
    return session


_SESSION = _get_session()


# ──────────────────── 对外主函数 ────────────────────
def unsplash_search(
    query: str,
    per_page: int = 3,
    *,
    limit: Optional[int] = None,            # 兼容旧参数
    orientation: str = "landscape",         # landscape / portrait / squarish
    size: str = "small",                    # small / regular / full
    order_by: str = "relevant",             # relevant / latest
    timeout: int = 8,
) -> List[bytes]:
    """
    在 Unsplash 搜索图片并返回字节数组列表；若未配置 Key 则返回空列表。
    """
    # ---------- 参数处理 ----------
    if limit is not None:
        per_page = limit
    if per_page <= 0:
        return []

    orientation = orientation.lower()
    size        = size.lower()
    order_by    = order_by.lower()

    if orientation not in {"landscape", "portrait", "squarish"}:
        raise UnsplashError("orientation 必须为 landscape / portrait / squarish")
    if size not in {"small", "regular", "full"}:
        raise UnsplashError("size 必须为 small / regular / full")
    if order_by not in {"latest", "relevant"}:
        raise UnsplashError("order_by 必须为 latest / relevant")

    # ---------- 读取 Key ----------
    try:
        client_id = _ensure_key("unsplash_key")
    except KeyError as exc:
        LOG.warning("Unsplash Key 未配置：%s；跳过搜索直接返回空列表", exc)
        return []

    # ---------- 调用 Unsplash ----------
    url = "https://api.unsplash.com/search/photos"
    params = {
        "query": query,
        "orientation": orientation,
        "per_page": per_page,
        "order_by": order_by,
    }
    headers = {"Authorization": f"Client-ID {client_id}"}

    try:
        resp = _SESSION.get(url, params=params, headers=headers, timeout=timeout)
        resp.raise_for_status()
        results = resp.json().get("results", [])
    except requests.RequestException as exc:
        LOG.error("Unsplash 搜索失败: %s", exc)
        raise UnsplashError(f"Unsplash 搜索失败: {exc}") from exc

    if not results:
        LOG.info("Unsplash '%s' 无搜索结果", query)
        return []

    # ---------- 下载图片 ----------
    blobs: List[bytes] = []
    for item in results:
        img_url = item["urls"][size]
        try:
            img_r = _SESSION.get(img_url, timeout=timeout)
            img_r.raise_for_status()
            blobs.append(img_r.content)
        except requests.RequestException as exc:
            LOG.warning("下载图片失败(%s): %s", img_url, exc)

    random.shuffle(blobs)
    return blobs