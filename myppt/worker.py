# worker.py
# -*- coding: utf-8 -*-
"""
后台线程：
1. ExtractWorker  —— 读取模板、抽取文本 → Markdown
2. OutlineWorker  —— 调 LLM 根据 Prompt 生成大纲
3. PPTWorker      —— 先生成候选图片 → 等 UI 回传选择 → 渲染并输出 PPT
"""
from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Callable

from pptx import Presentation
from PyQt5.QtCore import QEventLoop, QObject, QThread, pyqtSignal, Qt

# 业务逻辑
from modules.extractor   import pptx_to_markdown
from modules.llm_client  import gen_outline, gen_content
from modules.ppt_builder import prepare_image_candidates, render_ppt

# ----------------------------------------------------------------------
class ErrorDetails:
    """简易错误包装，便于跨线程传递"""
    def __init__(self, message: str, err_type: str = "GeneralError"):
        self.message = message
        self.type    = err_type
    def __str__(self): return f"{self.type}: {self.message}"


# ─────────────────────── 公共：进度回调 ───────────────────────
def _touch_ui_event_loop() -> None:
    """若已加载 PyQt5，调用一次 processEvents() 防卡死"""
    try:
        from PyQt5.QtWidgets import QApplication  # type: ignore
        app = QApplication.instance()
        if app is not None:
            app.processEvents()
    except ImportError:
        pass


def _make_progress_prefix(name: str) -> Callable[[float, str], None]:
    """
    返回一个 progress(pct, msg) 回调：
    1. 在终端打印统一格式 '[Worker] name  37.5%  msg'
    2. 每次调用后触发一次 UI 事件循环
    """
    def _cb(pct: float, txt: str):
        sys.stdout.write(
            f"\r[Worker] {name:<10} {pct*100:6.1f}%  {txt:<40}"
        )
        sys.stdout.flush()
        if pct >= 1.0:
            sys.stdout.write("\n")
        _touch_ui_event_loop()
    return _cb


# ============================== 1. 提取 ==============================
class ExtractWorker(QThread):
    finished = pyqtSignal(str, object)        # (markdown, error)

    def __init__(self, ppt_path: Path):
        super().__init__()
        self.ppt_path = ppt_path

    # 主工作 ---------------------------------------------------------
    def run(self):
        try:
            print("[Worker] Extract   开始抽取 Markdown ...")
            markdown, _ = pptx_to_markdown(self.ppt_path)
            print("[Worker] Extract   完成")
            self.finished.emit(markdown, None)
        except Exception as e:
            traceback.print_exc()
            self.finished.emit("", ErrorDetails(str(e), "MarkdownExtractionError"))


# ============================== 2. 大纲 ==============================
class OutlineWorker(QThread):
    finished = pyqtSignal(str, object)        # (outline, error)

    def __init__(self, md_text: str, prompts: dict):
        super().__init__()
        self.md      = md_text
        self.prompts = prompts

    def run(self):
        try:
            print("[Worker] Outline   调用大模型生成大纲 ...")
            outline = gen_outline(self.md, self.prompts)
            print("[Worker] Outline   生成完成")
            self.finished.emit(outline, None)
        except Exception as e:
            traceback.print_exc()
            self.finished.emit("", ErrorDetails(str(e), "OutlineGenerationError"))


# ============================== 3. 生成 PPT ==========================
class PPTWorker(QThread):
    """
    工作流程：
    (线程内)
        1. prepare_image_candidates() → emit images_ready(mapping)
    (UI 线程)
        2. 用户逐张选择 → emit choices_provided(mapping)
    (线程内)
        3. render_ppt() → emit finished(ppt_path, error)
    """
    images_ready     = pyqtSignal(dict)                 # {(slide_idx,rId): {...}}
    choices_provided = pyqtSignal(dict)                 # 用户回传选择
    finished         = pyqtSignal(object, object)       # (ppt_path, error)

    def __init__(self, template_path: Path,
                 outline_md: str,
                 prompts: dict):
        super().__init__()
        self.template_path = template_path
        self.outline_md    = outline_md
        self.prompts       = prompts
        self._user_choices: Optional[Dict[Tuple[int, str], bytes | None]] = None

        # 让 UI 把用户选择通过 choices_provided 信号发回来
        self.choices_provided.connect(self._recv_choices, Qt.QueuedConnection)

        # 本 worker 独享的进度回调
        self._progress = _make_progress_prefix("PPT")

    # ------------------------- 子线程入口 -------------------------
    def run(self):
        try:
            if not (self.prompts.get("topic") and self.outline_md):
                raise ValueError("Topic 或 Outline 数据缺失")

            topic_kw = self.prompts["topic"]

            # Step-1 生成完整 Markdown 文本 ---------------------------------
            print("[Worker] PPT       调用大模型生成正文 ...")
            full_md = gen_content(self.outline_md, self.prompts)
            print("[Worker] PPT       正文生成完毕")

            # Step-2 为占位图获取候选图片 -----------------------------------
            prs = Presentation(str(self.template_path))
            mapping  = prepare_image_candidates(
                prs, topic_kw, progress=self._progress)

            # 把 mapping 发送给 UI 线程
            self.images_ready.emit(mapping)

            # Step-3 等待 UI 选择 -----------------------------------------
            loop = QEventLoop()
            self._wait_loop = loop
            print("[Worker] PPT       等待用户选择图片 ...")
            loop.exec_()              # 在收到 choices 后会 quit()
            print("[Worker] PPT       已收到用户选择")

            user_choices = self._user_choices or {k: None for k in mapping}

            # Step-4 渲染最终 PPT ----------------------------------------
            ppt_path = render_ppt(
                self.template_path, markdown=full_md, topic=topic_kw,
                user_choices=user_choices, progress=self._progress
            )
            self.finished.emit(ppt_path, None)

        except Exception as e:
            traceback.print_exc()
            self.finished.emit(None, ErrorDetails(str(e), "PPTGenerationError"))

    # ------------------------- 接收用户选择 ------------------------
    def _recv_choices(self, choices: Dict[Tuple[int, str], bytes | None]):
        """由 UI 线程通过 choices_provided 信号调用"""
        self._user_choices = choices
        # 退出本地事件循环，继续 run() 之后的渲染
        if hasattr(self, "_wait_loop") and isinstance(self._wait_loop, QEventLoop):
            self._wait_loop.quit()