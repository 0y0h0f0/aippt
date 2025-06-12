# ui_main.py
# -*- coding: utf-8 -*-
"""
基于 PyQt5 的 AI-PPT 生成器
Author : Your Name / 2025-05
"""
from __future__ import annotations

# ——— 关闭冗余 ICC 日志（必须在首次 import PyQt5 之前） ———
import os
os.environ["QT_LOGGING_RULES"] = "qt.gui.icc=false"

import sys
import traceback
from enum import IntEnum
from pathlib import Path
from typing import Dict, Optional, Tuple

import config
from worker import ExtractWorker, OutlineWorker, PPTWorker

# ─── PyQt5 ───
from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui  import QPixmap, QIcon, QColor, QPainter
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QMessageBox, QProgressDialog, QPushButton,
    QStackedWidget, QVBoxLayout, QWidget, QLineEdit, QTextEdit,
    QDialog, QDialogButtonBox, QInputDialog, QGraphicsDropShadowEffect,
)

# ─── 资源路径 ───
BASE_DIR     = Path(__file__).resolve().parent
RES_DIR      = BASE_DIR / "resources"
TEMPLATE_DIR = BASE_DIR / "templates"

BG_IMAGE    = RES_DIR / "bg_main.jpg"
ICON_UPLOAD = RES_DIR / "icon_upload.png"
ICON_NEXT   = RES_DIR / "icon_next.png"
ICON_BACK   = RES_DIR / "icon_back.png"

# ─── 全局 QSS ───
GLOBAL_QSS = r"""
* { font-family: "微软雅黑"; color: #2c3e50; }
QLabel#titleLabel { font-size: 22px; font-weight: 700; }
/* 按钮 */
QPushButton[cls="primary"] {
    min-height: 42px; border-radius: 8px;
    background-color:qlineargradient(x1:0,y1:0,x2:0,y2:1,
                                     stop:0 #4e9af1, stop:1 #2574f5);
    color:#fff;
}
QPushButton[cls="primary"]:hover {
    background-color:qlineargradient(x1:0,y1:0,x2:0,y2:1,
                                     stop:0 #62a7ff, stop:1 #3b84ff);
}
QPushButton[cls="ghost"] {
    min-height:36px; border:2px solid #4e9af1;
    border-radius:6px; color:#4e9af1;
}
QPushButton[cls="ghost"]:hover { background:rgba(78,154,241,.07); }
/* 输入框 */
QLineEdit,QTextEdit{
    border:2px solid #dfe6e9; border-radius:6px;
    padding:8px 10px; font-size:15px;
}
QLineEdit:focus,QTextEdit:focus{ border:2px solid #4e9af1; }
/* 进度框 */
QProgressDialog{ background:#fff; border:2px solid #4e9af1; border-radius:10px; }
QProgressBar{ border:1px solid #dfe6e9; border-radius:5px; text-align:center; }
QProgressBar::chunk{ background:#4e9af1; }
QListWidget{ background:rgba(255,255,255,.85); border:none; }
"""

# ─── 通用小部件 ───
class StyledButton(QPushButton):
    """带悬浮阴影的按钮"""
    def __init__(self, text="", btn_type="primary",
                 icon: str | Path | None = None, parent=None):
        super().__init__(text, parent)
        self.setProperty("cls", btn_type)
        if icon: self.setIcon(QIcon(str(icon)))

    def enterEvent(self, e):
        self.setGraphicsEffect(QGraphicsDropShadowEffect(
            self, blurRadius=15, xOffset=0, yOffset=4,
            color=QColor(0, 0, 0, 80)))
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.setGraphicsEffect(None)
        super().leaveEvent(e)


class CardItem(QListWidgetItem):
    def __init__(self, text: str):
        super().__init__(text)
        self.setSizeHint(QSize(260, 50))


class Page(IntEnum):
    TEMPLATE = 0
    PROMPT   = 1
    OUTLINE  = 2


# ─── Step-1 模板选择页 ───
class TemplatePage(QWidget):
    template_ready = pyqtSignal(Path)

    def __init__(self, builtin_dir: Path):
        super().__init__()
        self.builtin_dir = builtin_dir
        self.selected: Optional[Path] = None
        self._bg_pix = QPixmap(str(BG_IMAGE))
        self._build()

    def paintEvent(self, e):
        if not self._bg_pix.isNull():
            QPainter(self).drawPixmap(self.rect(), self._bg_pix)
        super().paintEvent(e)

    def _build(self):
        self.list_widget = QListWidget()
        self.list_widget.setSpacing(8); self.list_widget.setFixedWidth(300)
        self._load_builtin()
        self.list_widget.itemClicked.connect(
            lambda it: setattr(self, "selected", self.builtin_dir / it.text())
        )

        upload_btn = StyledButton("上传自定义模板", "ghost", ICON_UPLOAD)
        upload_btn.clicked.connect(self._on_upload)
        next_btn   = StyledButton("下一步", "primary", ICON_NEXT)
        next_btn.clicked.connect(self._on_next)

        lay = QVBoxLayout(self); lay.setContentsMargins(80, 50, 80, 50)
        lay.addWidget(QLabel("请选择 PPT 模板", objectName="titleLabel"))
        lay.addSpacing(25); lay.addWidget(self.list_widget, 1)
        lay.addSpacing(15); lay.addWidget(upload_btn)
        lay.addSpacing(10); lay.addWidget(next_btn, 0, Qt.AlignRight)

    def _load_builtin(self):
        self.list_widget.clear()
        if self.builtin_dir.exists():
            for fp in self.builtin_dir.glob("*.pptx"):
                self.list_widget.addItem(CardItem(fp.name))

    def _on_upload(self):
        fp, _ = QFileDialog.getOpenFileName(
            self, "选择模板", "", "PowerPoint 模板 (*.pptx)")
        if fp:
            self.selected = Path(fp)
            self.template_ready.emit(self.selected)

    def _on_next(self):
        if not self.selected:
            QMessageBox.warning(self, "提示", "请先选择或上传模板"); return
        self.template_ready.emit(self.selected)


# ─── Step-2 Prompt 页 ───
class PromptPage(QWidget):
    prompt_ready   = pyqtSignal(dict)
    back_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._build()

    def _build(self):
        self.setStyleSheet("background:#f5f6f7;")
        # 不能用关键字 placeholderText，需手动调用
        self.report_type = QLineEdit(); self.report_type.setPlaceholderText("如：季度总结 / 项目路演 …")
        self.ai_role     = QLineEdit(); self.ai_role.setPlaceholderText("让 AI 扮演的专家角色")
        self.topic       = QLineEdit(); self.topic.setPlaceholderText("汇报主题")

        next_btn = StyledButton("生成大纲", "primary", ICON_NEXT)
        back_btn = StyledButton("返回",      "ghost",   ICON_BACK)
        next_btn.clicked.connect(self._on_next)
        back_btn.clicked.connect(self.back_requested)

        form = QVBoxLayout()
        for lab, w in [("报告类型", self.report_type),
                       ("AI 扮演角色", self.ai_role),
                       ("主题", self.topic)]:
            lbl = QLabel(f"{lab}："); lbl.setStyleSheet("font-weight:600;")
            form.addWidget(lbl); form.addWidget(w); form.addSpacing(15)

        lay = QVBoxLayout(self); lay.setContentsMargins(80, 50, 80, 50)
        lay.addWidget(QLabel("填写关键信息", objectName="titleLabel"))
        lay.addSpacing(15); lay.addLayout(form); lay.addStretch()
        row = QHBoxLayout(); row.addWidget(back_btn); row.addStretch(); row.addWidget(next_btn)
        lay.addLayout(row)

    def _on_next(self):
        rpt, role, topic = (w.text().strip() for w in
                            (self.report_type, self.ai_role, self.topic))
        if not (rpt and role and topic):
            QMessageBox.warning(self, "提示", "请完整填写三项内容"); return
        self.prompt_ready.emit({"report_type": rpt, "ai_role": role, "topic": topic})


# ─── Step-3 大纲 & 图片选择 ───
class OutlinePage(QWidget):
    outline_confirmed = pyqtSignal(str)
    back_requested    = pyqtSignal()

    class ImageChoiceDialog(QDialog):
        def __init__(self, origin: bytes, candidates: list[bytes], parent=None):
            super().__init__(parent)
            self.setWindowTitle("选择替换图片"); self.setModal(True)
            self._chosen: Optional[bytes] = None

            lst = QListWidget(); lst.setViewMode(QListWidget.IconMode)
            lst.setIconSize(QSize(220, 140)); lst.setSpacing(8)
            def add_img(blob: bytes, tip: str):
                pix = QPixmap(); pix.loadFromData(blob)
                itm = QListWidgetItem(tip); itm.setData(Qt.UserRole, blob)
                itm.setIcon(QIcon(pix)); lst.addItem(itm)
            add_img(origin, "原图")
            for i, b in enumerate(candidates, 1): add_img(b, f"候选 {i}")

            lst.itemClicked.connect(
                lambda it: setattr(self, "_chosen", it.data(Qt.UserRole)))
            btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            btn_box.accepted.connect(self.accept); btn_box.rejected.connect(self.reject)

            lay = QVBoxLayout(self)
            lay.addWidget(QLabel("点击选择想要的图片（或取消保持原图）："))
            lay.addWidget(lst, 1); lay.addWidget(btn_box)

        @property
        def chosen(self) -> Optional[bytes]:
            return self._chosen

    def __init__(self):
        super().__init__()
        self._build()

    def _build(self):
        self.setStyleSheet("background:#f8f9fa;")
        self.text_edit = QTextEdit()

        gen_btn  = StyledButton("生成完整 PPT", "primary", ICON_NEXT)
        back_btn = StyledButton("返回",       "ghost",   ICON_BACK)
        gen_btn.clicked.connect(self._emit_outline)
        back_btn.clicked.connect(self.back_requested)

        lay = QVBoxLayout(self); lay.setContentsMargins(60, 40, 60, 40)
        lay.addWidget(QLabel("编辑大纲", objectName="titleLabel"))
        lay.addSpacing(15); lay.addWidget(self.text_edit, 1); lay.addSpacing(10)
        row = QHBoxLayout(); row.addWidget(back_btn); row.addStretch(); row.addWidget(gen_btn)
        lay.addLayout(row)

    def load_outline(self, txt: str):
        self.text_edit.setPlainText(txt)

    def _emit_outline(self):
        txt = self.text_edit.toPlainText().strip()
        if not txt:
            QMessageBox.warning(self, "提示", "大纲为空"); return
        self.outline_confirmed.emit(txt)


# ─── 主窗口 ───
class MainWindow(QWidget):
    template_path: Optional[Path] = None
    prompts:       Optional[dict] = None
    outline_md:    Optional[str]  = None

    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI PPT 生成器"); self.resize(960, 680)

        QApplication.instance().setStyleSheet(GLOBAL_QSS)
        self._check_key_once()

        # 页面栈
        self.stack = QStackedWidget(self)
        self.page_template = TemplatePage(TEMPLATE_DIR)
        self.page_prompt   = PromptPage()
        self.page_outline  = OutlinePage()
        for p in (self.page_template, self.page_prompt, self.page_outline):
            self.stack.addWidget(p)
        QVBoxLayout(self).addWidget(self.stack)

        # 进度对话框
        self.progress = QProgressDialog("请稍候…", None, 0, 0, self,
                                        flags=Qt.FramelessWindowHint)
        self.progress.setWindowModality(Qt.ApplicationModal)
        self.progress.setCancelButton(None); self.progress.close()

        # 信号
        self.page_template.template_ready.connect(self._on_template_ready)
        self.page_prompt.prompt_ready.connect(self._on_prompt_ready)
        self.page_prompt.back_requested.connect(lambda: self._goto(Page.TEMPLATE))
        self.page_outline.outline_confirmed.connect(self._on_outline_confirmed)
        self.page_outline.back_requested.connect(lambda: self._goto(Page.PROMPT))

        # 线程句柄
        self.extract_worker: Optional[ExtractWorker] = None
        self.outline_worker: Optional[OutlineWorker] = None
        self.ppt_worker:     Optional[PPTWorker]     = None

    # ---------- 资源释放 ----------
    def closeEvent(self, e):
        for w in (self.extract_worker, self.outline_worker, self.ppt_worker):
            if w and w.isRunning(): w.quit(); w.wait(2000)
        super().closeEvent(e)

    # ---------- Key 初次检查 ----------
    def _check_key_once(self):
        for k, title in (("llm_key", "大模型 API-Key"),
                         ("unsplash_key", "Unsplash API-Key")):
            if config.get(k): continue
            key, ok = QInputDialog.getText(self, "首次使用", f"请输入 {title}：")
            if not ok or not key.strip():
                QMessageBox.critical(self, "错误",
                                     f"缺少 {title}，无法继续"); sys.exit(1)
            config.set_(k, key.strip())

    def _goto(self, p: Page):
        self.stack.setCurrentIndex(p)

    # ---------- Step-1 ----------
    def _on_template_ready(self, path: Path):
        self.template_path = path; self.prompts = None; self.outline_md = None
        self._goto(Page.PROMPT)

    # ---------- Step-2 ----------
    def _on_prompt_ready(self, prompts: dict):
        self.prompts = prompts
        assert self.template_path is not None
        self._show_progress("解析模板…")
        self.extract_worker = ExtractWorker(self.template_path)
        self.extract_worker.finished.connect(
            self._on_extract_done, Qt.QueuedConnection)
        self.extract_worker.start()

    def _on_extract_done(self, markdown: str, err):
        if err: self._critical("解析模板失败", err); return
        self._show_progress("AI 正在生成大纲…")
        assert self.prompts is not None
        self.outline_worker = OutlineWorker(markdown, self.prompts)
        self.outline_worker.finished.connect(
            self._on_outline_done, Qt.QueuedConnection)
        self.outline_worker.start()

    def _on_outline_done(self, outline: str, err):
        self.progress.close()
        if err or not outline:
            self._critical("生成大纲失败", err or "未知原因"); return
        self.outline_md = outline
        self.page_outline.load_outline(outline)
        self._goto(Page.OUTLINE)

    # ---------- Step-3 ----------
    def _on_outline_confirmed(self, outline_text: str):
        self.outline_md = outline_text
        self._show_progress("AI 正在生成 PPT…")

        assert self.template_path and self.prompts
        self.ppt_worker = PPTWorker(self.template_path,
                                    outline_text, self.prompts)
        self.ppt_worker.images_ready.connect(
            self._on_images_ready, Qt.QueuedConnection)
        self.ppt_worker.finished.connect(
            self._on_ppt_done, Qt.QueuedConnection)
        self.ppt_worker.start()

    def _on_images_ready(self, mapping):
        # 弹窗选择
        choices: Dict[Tuple[int, str], Optional[bytes]] = {}
        for key, val in mapping.items():
            if val is None: choices[key] = None; continue
            if isinstance(val, dict):
                origin, candidates = val.get("origin"), val.get("candidates", [])
            else:
                origin, *candidates = val
            dlg = OutlinePage.ImageChoiceDialog(origin, candidates, self)
            choices[key] = dlg.chosen if dlg.exec_() == QDialog.Accepted else None

        # 把用户选择回传给子线程 **关键改动在这里**
        if self.ppt_worker is not None:
            self.ppt_worker.choices_provided.emit(choices)  # ← 修正

    def _on_ppt_done(self, ppt_path: Optional[Path], err):
        self.progress.close()
        if err or not ppt_path:
            self._critical("PPT 生成失败", err or "未知原因"); return
        save_to, _ = QFileDialog.getSaveFileName(
            self, "保存 PPT", "generated.pptx",
            "PowerPoint 文件 (*.pptx)")
        if not save_to: return
        try:
            os.replace(ppt_path, save_to)
            QMessageBox.information(self, "完成", "已成功保存！")
        except Exception as e:
            self._critical("保存失败", e)

    # ---------- 工具 ----------
    def _show_progress(self, txt: str):
        self.progress.setLabelText(txt)
        if not self.progress.isVisible(): self.progress.show()
        QApplication.processEvents()

    def _critical(self, msg: str, err):
        self.progress.close()
        print("[Error]", msg, err, file=sys.stderr)
        traceback.print_exc()
        QMessageBox.critical(self, "错误", f"{msg}：{err}")


# ─── 入口 ───
def main():
    from PyQt5.QtCore import Qt
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv); app.setStyle("Fusion")
    win = MainWindow(); win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()