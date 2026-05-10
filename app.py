#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Fusion Pro v2.0
───────────────────
Gelişmiş PDF Araç Seti: Birleştirme | PDF→Word | Word→PDF
Windows 10/11 Uyumlu • Hızlı • Modern UI

Gerekli Kütüphaneler:
    pip install PyQt6 pypdf pdf2docx docx2pdf
"""

import sys
import os
import time
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QFileDialog, QMessageBox, QLabel, QProgressBar,
    QFrame, QStackedWidget, QListWidgetItem, QAbstractItemView,
    QSizePolicy, QCheckBox, QGraphicsDropShadowEffect, QScrollArea
)
from PyQt6.QtCore import (
    Qt, QSize, pyqtSignal, QThread, QPropertyAnimation,
    QEasingCurve, QTimer, QMimeData, QPoint, QRect
)
from PyQt6.QtGui import (
    QDragEnterEvent, QDropEvent, QIcon, QFont, QColor,
    QPainter, QPainterPath, QLinearGradient, QPixmap,
    QPalette, QPen, QBrush, QFontDatabase
)


# ═══════════════════════════════════════════════════════════════════════
# RENK PALETİ
# ═══════════════════════════════════════════════════════════════════════

class Colors:
    # Sidebar
    SIDEBAR_BG = "#0F172A"
    SIDEBAR_HOVER = "#1E293B"
    SIDEBAR_ACTIVE = "#3B82F6"
    SIDEBAR_TEXT = "#94A3B8"
    SIDEBAR_TEXT_ACTIVE = "#FFFFFF"
    SIDEBAR_BORDER = "#1E293B"

    # Genel
    BG = "#F1F5F9"
    CARD = "#FFFFFF"
    CARD_BORDER = "#E2E8F0"

    # Metin
    TEXT = "#0F172A"
    TEXT_SECONDARY = "#64748B"
    TEXT_MUTED = "#94A3B8"

    # Aksanlar
    PRIMARY = "#3B82F6"
    PRIMARY_HOVER = "#2563EB"
    PRIMARY_LIGHT = "#EFF6FF"

    SUCCESS = "#10B981"
    SUCCESS_HOVER = "#059669"
    SUCCESS_LIGHT = "#ECFDF5"

    DANGER = "#EF4444"
    DANGER_HOVER = "#DC2626"
    DANGER_LIGHT = "#FEF2F2"

    WARNING = "#F59E0B"
    WARNING_LIGHT = "#FFFBEB"
    WARNING_BORDER = "#FDE68A"

    # Giriş
    INPUT_BG = "#F8FAFC"
    INPUT_BORDER = "#CBD5E1"
    INPUT_FOCUS = "#3B82F6"

    # Progress
    PROGRESS_BG = "#E2E8F0"
    PROGRESS_CHUNK = "#3B82F6"


# ═══════════════════════════════════════════════════════════════════════
# ARKA PLAN İŞÇİ THREAD'LERİ
# ═══════════════════════════════════════════════════════════════════════

class MergeWorker(QThread):
    """PDF birleştirme işçisi - pypdf kullanır (hızlı)"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, files, output_path):
        super().__init__()
        self.files = files
        self.output_path = output_path

    def run(self):
        try:
            try:
                from pypdf import PdfWriter
            except ImportError:
                from PyPDF2 import PdfWriter

            writer = PdfWriter()
            total = len(self.files)

            for i, pdf_path in enumerate(self.files):
                self.progress.emit(
                    int((i / total) * 85),
                    f"İşleniyor ({i+1}/{total}): {os.path.basename(pdf_path)}"
                )
                writer.append(pdf_path)

            self.progress.emit(88, "PDF dosyası yazılıyor...")
            with open(self.output_path, "wb") as f:
                writer.write(f)
            writer.close()

            self.progress.emit(100, "Birleştirme tamamlandı!")
            self.finished.emit(True, self.output_path)

        except ImportError:
            self.finished.emit(False, "pypdf kütüphanesi bulunamadı!\n\npip install pypdf")
        except Exception as e:
            self.finished.emit(False, str(e))


class PdfToWordWorker(QThread):
    """PDF → Word dönüştürücü - pdf2docx kullanır"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, input_path, output_path):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path

    def run(self):
        try:
            from pdf2docx import Converter

            self.progress.emit(5, "PDF dosyası okunuyor...")
            cv = Converter(self.input_path)

            self.progress.emit(15, "Sayfalar analiz ediliyor...")
            page_count = len(cv.fitz_doc)  # type: ignore

            self.progress.emit(25, f"Dönüştürülüyor ({page_count} sayfa)...")
            cv.convert(self.output_path)
            cv.close()

            self.progress.emit(100, "Dönüştürme tamamlandı!")
            self.finished.emit(True, self.output_path)

        except ImportError:
            self.finished.emit(
                False,
                "pdf2docx kütüphanesi bulunamadı!\n\npip install pdf2docx"
            )
        except Exception as e:
            self.finished.emit(False, str(e))


class WordToPdfWorker(QThread):
    """Word → PDF dönüştürücü - docx2pdf kullanır (MS Word gerektirir)"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, input_path, output_path):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path

    def run(self):
        try:
            from docx2pdf import convert

            self.progress.emit(15, "Word dosyası okunuyor...")
            self.progress.emit(40, "PDF'ye dönüştürülüyor...")
            convert(self.input_path, self.output_path)

            self.progress.emit(100, "Dönüştürme tamamlandı!")
            self.finished.emit(True, self.output_path)

        except ImportError:
            self.finished.emit(
                False,
                "docx2pdf kütüphanesi bulunamadı!\n\n"
                "pip install docx2pdf\n\n"
                "Not: Bu özellik bilgisayarınızda Microsoft Word'ün\n"
                "yüklü olmasını gerektirir."
            )
        except Exception as e:
            self.finished.emit(False, str(e))


# ═══════════════════════════════════════════════════════════════════════
# ÖZEL WIDGET'LAR
# ═══════════════════════════════════════════════════════════════════════

class SidebarButton(QPushButton):
    """Kenar çubuğu navigasyon butonu (özel çizim)"""

    def __init__(self, icon_char: str, text: str, parent=None):
        super().__init__(parent)
        self.icon_char = icon_char
        self.label_text = text
        self.is_active = False
        self.setFixedHeight(48)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setCheckable(True)
        self._anim_opacity = 0.0

    def paintEvent(self, event):  # noqa
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()

        # Arka plan
        if self.is_active or self.isChecked():
            # Aktif: mavi arka plan, sol kenarda beyaz çubuk
            p.setBrush(QColor(Colors.SIDEBAR_ACTIVE))
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(8, 4, w - 16, h - 8, 10, 10)
            text_color = QColor(Colors.SIDEBAR_TEXT_ACTIVE)

            # Sol kenar çubuk
            p.setBrush(QColor("#FFFFFF"))
            p.drawRoundedRect(0, 12, 3, h - 24, 2, 2)

        elif self.underMouse():
            p.setBrush(QColor(Colors.SIDEBAR_HOVER))
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(8, 4, w - 16, h - 8, 10, 10)
            text_color = QColor(Colors.SIDEBAR_TEXT_ACTIVE)
        else:
            text_color = QColor(Colors.SIDEBAR_TEXT)

        # İkon
        p.setPen(text_color)
        icon_font = QFont("Segoe UI Emoji", 15)
        p.setFont(icon_font)
        p.drawText(QRect(18, 0, 36, h), Qt.AlignmentFlag.AlignCenter, self.icon_char)

        # Metin
        text_font = QFont("Segoe UI", 11)
        text_font.setWeight(QFont.Weight.Medium)
        p.setFont(text_font)
        p.drawText(QRect(58, 0, w - 70, h), Qt.AlignmentFlag.AlignVCenter, self.label_text)

        p.end()


class FileListWidget(QListWidget):
    """Gelişmiş dosya listesi - sürükle-bırak ve dış dosya desteği"""
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setAlternatingRowColors(True)
        self.setSpacing(2)
        self.setStyleSheet(f"""
            QListWidget {{
                background-color: {Colors.CARD};
                border: 1px solid {Colors.CARD_BORDER};
                border-radius: 10px;
                padding: 6px;
                font-family: 'Segoe UI';
                font-size: 13px;
                outline: none;
            }}
            QListWidget::item {{
                padding: 10px 14px;
                border-radius: 8px;
                margin: 1px 0;
                border: none;
            }}
            QListWidget::item:selected {{
                background-color: {Colors.PRIMARY_LIGHT};
                color: {Colors.PRIMARY};
            }}
            QListWidget::item:hover:!selected {{
                background-color: #F8FAFC;
            }}
            QListWidget::item:alternate {{
                background-color: #FAFBFD;
            }}
        """)

    # Dış dosya sürükle-bırak
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            files = []
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(".pdf"):
                    files.append(path)
            if files:
                self.files_dropped.emit(files)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class DropZone(QFrame):
    """Dosya sürükle-bırak alanı"""
    files_dropped = pyqtSignal(list)

    def __init__(self, accepted_exts: list, parent=None):
        super().__init__(parent)
        self.accepted_exts = accepted_exts
        self.setAcceptDrops(True)
        self.setMinimumHeight(130)
        self._hovering = False
        self._apply_style(False)

    def _apply_style(self, hover: bool):
        if hover:
            self.setStyleSheet(f"""
                DropZone {{
                    background-color: {Colors.PRIMARY_LIGHT};
                    border: 2px dashed {Colors.PRIMARY};
                    border-radius: 14px;
                }}
            """)
        else:
            self.setStyleSheet(f"""
                DropZone {{
                    background-color: {Colors.INPUT_BG};
                    border: 2px dashed {Colors.INPUT_BORDER};
                    border-radius: 14px;
                }}
            """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._apply_style(True)

    def dragLeaveEvent(self, event):
        self._apply_style(False)

    def dropEvent(self, event):
        self._apply_style(False)
        files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            ext = os.path.splitext(path)[1].lower()
            if ext in self.accepted_exts:
                files.append(path)
        if files:
            self.files_dropped.emit(files)
        event.acceptProposedAction()


class AnimatedProgressBar(QProgressBar):
    """Animasyonlu ilerleme çubuğu"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setRange(0, 100)
        self.setValue(0)
        self.setTextVisible(False)
        self.setFixedHeight(6)
        self.setStyleSheet(f"""
            QProgressBar {{
                background-color: {Colors.PROGRESS_BG};
                border: none;
                border-radius: 3px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 {Colors.PRIMARY},
                    stop:1 #60A5FA
                );
                border-radius: 3px;
            }}
        """)


# ═══════════════════════════════════════════════════════════════════════
# YARDIMCI: KART OLUŞTURUCU
# ═══════════════════════════════════════════════════════════════════════

def make_card(title="", subtitle=""):
    """Gölgeli kart widget'ı döndürür: (card_frame, layout)"""
    card = QFrame()
    card.setObjectName("card")
    card.setStyleSheet(f"""
        QFrame#card {{
            background-color: {Colors.CARD};
            border: 1px solid {Colors.CARD_BORDER};
            border-radius: 14px;
        }}
    """)
    # Gölge efekti
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(20)
    shadow.setXOffset(0)
    shadow.setYOffset(4)
    shadow.setColor(QColor(0, 0, 0, 25))
    card.setGraphicsEffect(shadow)

    layout = QVBoxLayout(card)
    layout.setContentsMargins(22, 20, 22, 20)
    layout.setSpacing(14)

    if title:
        lbl = QLabel(title)
        lbl.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {Colors.TEXT}; border:none;")
        layout.addWidget(lbl)
    if subtitle:
        lbl = QLabel(subtitle)
        lbl.setStyleSheet(f"font-size: 13px; color: {Colors.TEXT_SECONDARY}; border:none;")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

    return card, layout


def make_button(text, color="primary", icon_text="", min_w=0):
    """Stil verilmiş buton döndürür."""
    label = f"{icon_text}  {text}" if icon_text else text
    btn = QPushButton(label)

    color_map = {
        "primary":  (Colors.PRIMARY, Colors.PRIMARY_HOVER),
        "success":  (Colors.SUCCESS, Colors.SUCCESS_HOVER),
        "danger":   (Colors.DANGER, Colors.DANGER_HOVER),
    }
    bg, hover = color_map.get(color, (Colors.PRIMARY, Colors.PRIMARY_HOVER))

    btn.setStyleSheet(f"""
        QPushButton {{
            background-color: {bg};
            color: #FFFFFF;
            border: none;
            border-radius: 10px;
            padding: 10px 22px;
            font-family: 'Segoe UI';
            font-size: 13px;
            font-weight: 600;
        }}
        QPushButton:hover {{
            background-color: {hover};
        }}
        QPushButton:pressed {{
            background-color: {hover};
            padding-top: 11px;
            padding-bottom: 9px;
        }}
        QPushButton:disabled {{
            background-color: #CBD5E1;
            color: #94A3B8;
        }}
    """)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setMinimumHeight(42)
    if min_w:
        btn.setMinimumWidth(min_w)
    return btn


def make_outline_button(text, icon_text=""):
    """İkincil (outline) buton"""
    label = f"{icon_text}  {text}" if icon_text else text
    btn = QPushButton(label)
    btn.setStyleSheet(f"""
        QPushButton {{
            background-color: transparent;
            color: {Colors.TEXT};
            border: 1.5px solid {Colors.CARD_BORDER};
            border-radius: 10px;
            padding: 8px 18px;
            font-family: 'Segoe UI';
            font-size: 13px;
            font-weight: 500;
        }}
        QPushButton:hover {{
            background-color: {Colors.BG};
            border-color: #94A3B8;
        }}
        QPushButton:disabled {{
            color: #CBD5E1;
            border-color: #E2E8F0;
        }}
    """)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setMinimumHeight(38)
    return btn


# ═══════════════════════════════════════════════════════════════════════
# ANA UYGULAMA PENCERESİ
# ═══════════════════════════════════════════════════════════════════════

class PDFFusionPro(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Fusion Pro")
        self.setMinimumSize(1000, 660)
        self.resize(1100, 720)
        self.setAcceptDrops(True)

        # Pencere ikonu (programatik)
        self._set_app_icon()

        # Durum değişkenleri
        self.pdf_files: list[str] = []
        self.current_worker: QThread | None = None
        self.p2w_selected_file: str | None = None
        self.w2p_selected_file: str | None = None

        self._build_ui()
        self._connect_signals()
        self._switch_page(0)

    # ─── Pencere İkonu ─────────────────────────────────────────────
    def _set_app_icon(self):
        px = QPixmap(256, 256)
        px.fill(Qt.GlobalColor.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Arka plan — yuvarlatılmış kare, mavi degrade
        grad = QLinearGradient(0, 0, 256, 256)
        grad.setColorAt(0, QColor(Colors.PRIMARY))
        grad.setColorAt(1, QColor("#60A5FA"))
        p.setBrush(QBrush(grad))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(0, 0, 256, 256, 52, 52)

        # "P" harfi
        p.setPen(QColor("#FFFFFF"))
        f = QFont("Segoe UI", 130, QFont.Weight.Bold)
        p.setFont(f)
        p.drawText(QRect(0, 0, 256, 256), Qt.AlignmentFlag.AlignCenter, "P")
        p.end()

        icon = QIcon(px)
        self.setWindowIcon(icon)

        # Masaüstü kısayolu için icon.ico dosyasını oluştur (ilk çalıştırmada)
        self._save_icon_file(px)

    @staticmethod
    def _save_icon_file(px: QPixmap):
        """icon.ico dosyasını uygulama dizinine kaydeder (kısayol ikonu için)."""
        ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        if not os.path.exists(ico_path):
            px.save(ico_path, "ICO")

    # ─── ANA UI YAPISI ─────────────────────────────────────────────
    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── SIDEBAR ──
        sidebar = QFrame()
        sidebar.setFixedWidth(240)
        sidebar.setStyleSheet(f"""
            QFrame {{
                background-color: {Colors.SIDEBAR_BG};
                border-right: 1px solid {Colors.SIDEBAR_BORDER};
            }}
        """)
        sb_layout = QVBoxLayout(sidebar)
        sb_layout.setContentsMargins(12, 20, 12, 20)
        sb_layout.setSpacing(4)

        # Logo
        logo = QLabel("  PDF Fusion Pro")
        logo.setStyleSheet(f"""
            color: #FFFFFF;
            font-size: 19px;
            font-weight: 800;
            padding: 6px 0 18px 0;
            letter-spacing: 0.5px;
        """)
        sb_layout.addWidget(logo)

        # Üst ayırıcı
        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background-color: {Colors.SIDEBAR_HOVER};")
        sb_layout.addWidget(sep)
        sb_layout.addSpacing(14)

        # Bölüm başlığı
        section_lbl = QLabel("  ARAÇLAR")
        section_lbl.setStyleSheet(f"""
            color: {Colors.SIDEBAR_TEXT};
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 1.5px;
            padding-bottom: 6px;
        """)
        sb_layout.addWidget(section_lbl)

        # Navigasyon butonları
        self.nav_buttons: list[SidebarButton] = []
        nav_items = [
            ("📋", "PDF Birleştir"),
            ("📝", "PDF → Word"),
            ("📑", "Word → PDF"),
        ]
        for icon_ch, text in nav_items:
            btn = SidebarButton(icon_ch, text)
            idx = len(self.nav_buttons)
            btn.clicked.connect(lambda checked, i=idx: self._switch_page(i))
            sb_layout.addWidget(btn)
            self.nav_buttons.append(btn)

        sb_layout.addStretch()

        # Alt bilgi
        sep2 = QFrame()
        sep2.setFixedHeight(1)
        sep2.setStyleSheet(f"background-color: {Colors.SIDEBAR_HOVER};")
        sb_layout.addWidget(sep2)
        sb_layout.addSpacing(8)

        info = QLabel("  PDF Fusion Pro v2.0\n  Windows 10/11")
        info.setStyleSheet(f"color: {Colors.SIDEBAR_TEXT}; font-size: 11px; line-height: 18px;")
        sb_layout.addWidget(info)

        root.addWidget(sidebar)

        # ── İÇERİK ALANI ──
        content = QFrame()
        content.setStyleSheet(f"background-color: {Colors.BG};")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._page_merge())
        self.stack.addWidget(self._page_pdf_to_word())
        self.stack.addWidget(self._page_word_to_pdf())
        content_layout.addWidget(self.stack)

        root.addWidget(content, 1)

    # ═══════════════════════════════════════════════════════════════
    # SAYFA 1 — PDF BİRLEŞTİR
    # ═══════════════════════════════════════════════════════════════
    def _page_merge(self):
        page = QWidget()
        page.setStyleSheet(f"background-color: {Colors.BG};")

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("QScrollArea{border:none; background:transparent;}")

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(30, 28, 30, 28)
        layout.setSpacing(18)

        # Başlık
        header = QLabel("PDF Birleştir")
        header.setStyleSheet(f"font-size: 26px; font-weight: 800; color: {Colors.TEXT};")
        layout.addWidget(header)

        desc = QLabel("Birden fazla PDF dosyasını sürükle-bırak veya dosya seçerek listeye ekleyin, ardından tek bir PDF olarak birleştirin.")
        desc.setStyleSheet(f"font-size: 13px; color: {Colors.TEXT_SECONDARY}; margin-bottom: 4px;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # ── Dosya listesi kartı ──
        card, clayout = make_card()

        # Araç çubuğu
        tb = QHBoxLayout()
        self.m_add_btn = make_button("Dosya Ekle", "primary", "➕")
        self.m_remove_btn = make_outline_button("Kaldır", "✕")
        self.m_clear_btn = make_outline_button("Temizle", "🗑")
        self.m_up_btn = make_outline_button("▲")
        self.m_down_btn = make_outline_button("▼")
        self.m_up_btn.setFixedWidth(48)
        self.m_down_btn.setFixedWidth(48)
        tb.addWidget(self.m_add_btn)
        tb.addWidget(self.m_remove_btn)
        tb.addWidget(self.m_clear_btn)
        tb.addStretch()
        tb.addWidget(self.m_up_btn)
        tb.addWidget(self.m_down_btn)
        clayout.addLayout(tb)

        # Dosya listesi
        self.m_list = FileListWidget()
        self.m_list.setMinimumHeight(220)
        clayout.addWidget(self.m_list)

        # Dosya bilgi etiketi
        self.m_info = QLabel("Henüz dosya eklenmedi")
        self.m_info.setStyleSheet(f"font-size: 12px; color: {Colors.TEXT_MUTED}; border:none;")
        clayout.addWidget(self.m_info)

        layout.addWidget(card)

        # ── Alt bölüm ──
        bottom = QHBoxLayout()
        bottom.setSpacing(16)

        # Seçenekler kartı
        opt_card = QFrame()
        opt_card.setObjectName("optcard")
        opt_card.setStyleSheet(f"""
            QFrame#optcard {{
                background-color: {Colors.CARD};
                border: 1px solid {Colors.CARD_BORDER};
                border-radius: 14px;
            }}
        """)
        opt_shadow = QGraphicsDropShadowEffect()
        opt_shadow.setBlurRadius(16)
        opt_shadow.setXOffset(0)
        opt_shadow.setYOffset(3)
        opt_shadow.setColor(QColor(0, 0, 0, 18))
        opt_card.setGraphicsEffect(opt_shadow)
        opt_layout = QVBoxLayout(opt_card)
        opt_layout.setContentsMargins(18, 16, 18, 16)
        opt_layout.setSpacing(10)

        opt_title = QLabel("⚙️  Seçenekler")
        opt_title.setStyleSheet(f"font-size: 14px; font-weight: 700; color: {Colors.TEXT};")
        opt_layout.addWidget(opt_title)

        self.m_open_after = QCheckBox("Birleştirme sonrası dosyayı aç")
        self.m_open_after.setChecked(True)
        self.m_open_after.setStyleSheet(f"""
            QCheckBox {{
                font-size: 13px; color: {Colors.TEXT}; spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 20px; height: 20px;
                border-radius: 5px;
                border: 2px solid {Colors.INPUT_BORDER};
                background: {Colors.INPUT_BG};
            }}
            QCheckBox::indicator:checked {{
                background-color: {Colors.PRIMARY};
                border-color: {Colors.PRIMARY};
                image: none;
            }}
        """)
        opt_layout.addWidget(self.m_open_after)
        opt_layout.addStretch()

        bottom.addWidget(opt_card, 1)

        # İlerleme + Buton kartı
        action_card = QFrame()
        action_card.setObjectName("actcard")
        action_card.setStyleSheet(f"""
            QFrame#actcard {{
                background-color: {Colors.CARD};
                border: 1px solid {Colors.CARD_BORDER};
                border-radius: 14px;
            }}
        """)
        act_shadow = QGraphicsDropShadowEffect()
        act_shadow.setBlurRadius(16)
        act_shadow.setXOffset(0)
        act_shadow.setYOffset(3)
        act_shadow.setColor(QColor(0, 0, 0, 18))
        action_card.setGraphicsEffect(act_shadow)
        act_layout = QVBoxLayout(action_card)
        act_layout.setContentsMargins(18, 16, 18, 16)
        act_layout.setSpacing(10)

        self.m_progress = AnimatedProgressBar()
        act_layout.addWidget(self.m_progress)

        self.m_status = QLabel("Dosya ekleyerek başlayın")
        self.m_status.setStyleSheet(f"font-size: 12px; color: {Colors.TEXT_SECONDARY};")
        self.m_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        act_layout.addWidget(self.m_status)

        act_layout.addStretch()

        self.m_merge_btn = make_button("PDF'leri Birleştir", "success", "🔗", 220)
        self.m_merge_btn.setEnabled(False)
        act_layout.addWidget(self.m_merge_btn)

        bottom.addWidget(action_card, 1)

        layout.addLayout(bottom)
        layout.addStretch()

        scroll.setWidget(inner)
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.addWidget(scroll)
        return page

    # ═══════════════════════════════════════════════════════════════
    # SAYFA 2 — PDF → WORD
    # ═══════════════════════════════════════════════════════════════
    def _page_pdf_to_word(self):
        page = QWidget()
        page.setStyleSheet(f"background-color: {Colors.BG};")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 28, 30, 28)
        layout.setSpacing(18)

        header = QLabel("PDF → Word Dönüştürücü")
        header.setStyleSheet(f"font-size: 26px; font-weight: 800; color: {Colors.TEXT};")
        layout.addWidget(header)

        desc = QLabel("PDF dosyalarını düzenlenebilir Word (.docx) formatına dönüştürün. Metin, tablolar ve temel biçimlendirme korunur.")
        desc.setStyleSheet(f"font-size: 13px; color: {Colors.TEXT_SECONDARY};")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Ana kart
        card, clayout = make_card()

        # Dosya seçimi
        file_row = QHBoxLayout()
        self.p2w_label = QLabel("📂  Bir PDF dosyası seçin veya aşağıya sürükleyin")
        self.p2w_label.setStyleSheet(f"""
            background-color: {Colors.INPUT_BG};
            border: 1.5px solid {Colors.INPUT_BORDER};
            border-radius: 10px;
            padding: 14px 16px;
            font-size: 13px;
            color: {Colors.TEXT_SECONDARY};
        """)
        self.p2w_label.setMinimumHeight(50)

        self.p2w_select_btn = make_button("PDF Seç", "primary", "📄")
        self.p2w_select_btn.setFixedWidth(160)
        file_row.addWidget(self.p2w_label, 1)
        file_row.addWidget(self.p2w_select_btn)
        clayout.addLayout(file_row)

        # Sürükle-bırak alanı
        self.p2w_drop = DropZone([".pdf"])
        drop_layout = QVBoxLayout(self.p2w_drop)
        drop_icon = QLabel("📥")
        drop_icon.setStyleSheet("font-size: 32px; border:none; background:transparent;")
        drop_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_text = QLabel("PDF dosyasını buraya sürükleyip bırakın")
        drop_text.setStyleSheet(f"color: {Colors.TEXT_SECONDARY}; font-size: 14px; border:none; background:transparent;")
        drop_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addStretch()
        drop_layout.addWidget(drop_icon)
        drop_layout.addWidget(drop_text)
        drop_layout.addStretch()
        clayout.addWidget(self.p2w_drop)

        # İlerleme
        self.p2w_progress = AnimatedProgressBar()
        clayout.addWidget(self.p2w_progress)

        self.p2w_status = QLabel("")
        self.p2w_status.setStyleSheet(f"font-size: 12px; color: {Colors.TEXT_SECONDARY}; border:none;")
        self.p2w_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        clayout.addWidget(self.p2w_status)

        # Dönüştür butonu
        self.p2w_convert_btn = make_button("Word'e Dönüştür (.docx)", "success", "🔄", 250)
        self.p2w_convert_btn.setEnabled(False)
        clayout.addWidget(self.p2w_convert_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(card)
        layout.addStretch()
        return page

    # ═══════════════════════════════════════════════════════════════
    # SAYFA 3 — WORD → PDF
    # ═══════════════════════════════════════════════════════════════
    def _page_word_to_pdf(self):
        page = QWidget()
        page.setStyleSheet(f"background-color: {Colors.BG};")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 28, 30, 28)
        layout.setSpacing(18)

        header = QLabel("Word → PDF Dönüştürücü")
        header.setStyleSheet(f"font-size: 26px; font-weight: 800; color: {Colors.TEXT};")
        layout.addWidget(header)

        desc = QLabel("Word (.docx / .doc) dosyalarını profesyonel PDF formatına dönüştürün.")
        desc.setStyleSheet(f"font-size: 13px; color: {Colors.TEXT_SECONDARY};")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Uyarı
        warn = QFrame()
        warn.setStyleSheet(f"""
            background-color: {Colors.WARNING_LIGHT};
            border: 1px solid {Colors.WARNING_BORDER};
            border-radius: 10px;
        """)
        warn_layout = QHBoxLayout(warn)
        warn_layout.setContentsMargins(14, 10, 14, 10)
        warn_icon = QLabel("⚠️")
        warn_icon.setStyleSheet("font-size: 18px; border:none; background:transparent;")
        warn_text = QLabel("Bu özellik bilgisayarınızda <b>Microsoft Word</b>'ün yüklü olmasını gerektirir.")
        warn_text.setStyleSheet(f"font-size: 12px; color: #92400E; border:none; background:transparent;")
        warn_text.setWordWrap(True)
        warn_layout.addWidget(warn_icon)
        warn_layout.addWidget(warn_text, 1)
        layout.addWidget(warn)

        # Ana kart
        card, clayout = make_card()

        # Dosya seçimi
        file_row = QHBoxLayout()
        self.w2p_label = QLabel("📂  Bir Word dosyası seçin veya aşağıya sürükleyin")
        self.w2p_label.setStyleSheet(f"""
            background-color: {Colors.INPUT_BG};
            border: 1.5px solid {Colors.INPUT_BORDER};
            border-radius: 10px;
            padding: 14px 16px;
            font-size: 13px;
            color: {Colors.TEXT_SECONDARY};
        """)
        self.w2p_label.setMinimumHeight(50)

        self.w2p_select_btn = make_button("Word Seç", "primary", "📄")
        self.w2p_select_btn.setFixedWidth(160)
        file_row.addWidget(self.w2p_label, 1)
        file_row.addWidget(self.w2p_select_btn)
        clayout.addLayout(file_row)

        # Sürükle-bırak alanı
        self.w2p_drop = DropZone([".docx", ".doc"])
        drop_layout = QVBoxLayout(self.w2p_drop)
        drop_icon = QLabel("📥")
        drop_icon.setStyleSheet("font-size: 32px; border:none; background:transparent;")
        drop_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_text = QLabel("Word dosyasını buraya sürükleyip bırakın")
        drop_text.setStyleSheet(f"color: {Colors.TEXT_SECONDARY}; font-size: 14px; border:none; background:transparent;")
        drop_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addStretch()
        drop_layout.addWidget(drop_icon)
        drop_layout.addWidget(drop_text)
        drop_layout.addStretch()
        clayout.addWidget(self.w2p_drop)

        # İlerleme
        self.w2p_progress = AnimatedProgressBar()
        clayout.addWidget(self.w2p_progress)

        self.w2p_status = QLabel("")
        self.w2p_status.setStyleSheet(f"font-size: 12px; color: {Colors.TEXT_SECONDARY}; border:none;")
        self.w2p_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        clayout.addWidget(self.w2p_status)

        # Dönüştür butonu
        self.w2p_convert_btn = make_button("PDF'ye Dönüştür (.pdf)", "success", "🔄", 250)
        self.w2p_convert_btn.setEnabled(False)
        clayout.addWidget(self.w2p_convert_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(card)
        layout.addStretch()
        return page

    # ═══════════════════════════════════════════════════════════════
    # SİNYAL BAĞLANTILARI
    # ═══════════════════════════════════════════════════════════════
    def _connect_signals(self):
        # Merge
        self.m_add_btn.clicked.connect(self._m_add_files)
        self.m_remove_btn.clicked.connect(self._m_remove)
        self.m_clear_btn.clicked.connect(self._m_clear)
        self.m_up_btn.clicked.connect(self._m_move_up)
        self.m_down_btn.clicked.connect(self._m_move_down)
        self.m_merge_btn.clicked.connect(self._m_execute)
        self.m_list.files_dropped.connect(self._m_add_to_list)
        self.m_list.itemSelectionChanged.connect(self._m_update_state)

        # PDF → Word
        self.p2w_select_btn.clicked.connect(self._p2w_select)
        self.p2w_convert_btn.clicked.connect(self._p2w_execute)
        self.p2w_drop.files_dropped.connect(lambda f: self._p2w_set(f[0]))

        # Word → PDF
        self.w2p_select_btn.clicked.connect(self._w2p_select)
        self.w2p_convert_btn.clicked.connect(self._w2p_execute)
        self.w2p_drop.files_dropped.connect(lambda f: self._w2p_set(f[0]))

    # ═══════════════════════════════════════════════════════════════
    # NAVİGASYON
    # ═══════════════════════════════════════════════════════════════
    def _switch_page(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, btn in enumerate(self.nav_buttons):
            btn.is_active = (i == idx)
            btn.setChecked(i == idx)
            btn.update()

    # ─── Ana pencere sürükle-bırak (merge sayfası) ────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".pdf"):
                files.append(path)
        if files and self.stack.currentIndex() == 0:
            self._m_add_to_list(files)
        event.acceptProposedAction()

    # ═══════════════════════════════════════════════════════════════
    # MERGE İŞLEMLERİ
    # ═══════════════════════════════════════════════════════════════
    def _m_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "PDF Dosyaları Seçin", "",
            "PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)"
        )
        if files:
            self._m_add_to_list(files)

    def _m_add_to_list(self, files: list[str]):
        added = 0
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            if f in self.pdf_files:
                continue
            try:
                try:
                    from pypdf import PdfReader
                except ImportError:
                    from PyPDF2 import PdfReader
                reader = PdfReader(f)
                pages = len(reader.pages)
            except Exception as e:
                QMessageBox.warning(
                    self, "Hatalı Dosya",
                    f"Bu dosya okunamadı:\n{os.path.basename(f)}\n\n{e}"
                )
                continue

            self.pdf_files.append(f)
            name = os.path.basename(f)
            size_mb = os.path.getsize(f) / (1024 * 1024)
            display = f"📄  {name}   •   {pages} sayfa   •   {size_mb:.1f} MB"

            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole, f)
            item.setToolTip(f)
            self.m_list.addItem(item)
            added += 1

        self._m_update_info()
        self._m_update_state()

    def _m_remove(self):
        for item in self.m_list.selectedItems():
            path = item.data(Qt.ItemDataRole.UserRole)
            if path in self.pdf_files:
                self.pdf_files.remove(path)
            self.m_list.takeItem(self.m_list.row(item))
        self._m_update_info()
        self._m_update_state()

    def _m_clear(self):
        self.m_list.clear()
        self.pdf_files.clear()
        self._m_update_info()
        self._m_update_state()
        self.m_progress.setValue(0)
        self.m_status.setText("Dosya ekleyerek başlayın")

    def _m_move_up(self):
        row = self.m_list.currentRow()
        if row > 0:
            item = self.m_list.takeItem(row)
            self.m_list.insertItem(row - 1, item)
            self.m_list.setCurrentRow(row - 1)

    def _m_move_down(self):
        row = self.m_list.currentRow()
        if 0 <= row < self.m_list.count() - 1:
            item = self.m_list.takeItem(row)
            self.m_list.insertItem(row + 1, item)
            self.m_list.setCurrentRow(row + 1)

    def _m_update_info(self):
        n = self.m_list.count()
        if n > 0:
            paths = [self.m_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(n)]
            total_mb = sum(os.path.getsize(p) for p in paths) / (1024 * 1024)
            self.m_info.setText(f"📊  Toplam {n} dosya  •  {total_mb:.2f} MB")
        else:
            self.m_info.setText("Henüz dosya eklenmedi")

    def _m_update_state(self):
        has = self.m_list.count() > 0
        self.m_merge_btn.setEnabled(has and self.current_worker is None)
        if has:
            self.m_status.setText(f"✅ {self.m_list.count()} PDF birleştirmeye hazır")
        elif self.current_worker is None:
            self.m_status.setText("Dosya ekleyerek başlayın")

    def _m_execute(self):
        if self.m_list.count() < 1:
            return

        ordered = [
            self.m_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.m_list.count())
        ]

        output, _ = QFileDialog.getSaveFileName(
            self, "Birleştirilmiş PDF'i Kaydet",
            "birlesmis_dosya.pdf",
            "PDF Dosyaları (*.pdf)"
        )
        if not output:
            return
        if not output.lower().endswith(".pdf"):
            output += ".pdf"

        self.m_merge_btn.setEnabled(False)
        self.m_progress.setValue(0)
        self.m_status.setText("⏳ Birleştirme başlatılıyor...")

        self.current_worker = MergeWorker(ordered, output)
        self.current_worker.progress.connect(self._m_on_progress)
        self.current_worker.finished.connect(self._m_on_finished)
        self.current_worker.start()

    def _m_on_progress(self, val, text):
        self.m_progress.setValue(val)
        self.m_status.setText(text)

    def _m_on_finished(self, ok, result):
        self.current_worker = None
        self.m_merge_btn.setEnabled(True)

        if ok:
            self.m_progress.setValue(100)
            self.m_status.setText("✅ Birleştirme başarıyla tamamlandı!")
            QMessageBox.information(
                self, "Başarılı",
                f"PDF'ler başarıyla birleştirildi!\n\n📄 {os.path.basename(result)}"
            )
            if self.m_open_after.isChecked():
                self._open_file(result)
            self._m_clear()
        else:
            self.m_status.setText(f"❌ Hata oluştu")
            QMessageBox.critical(self, "Hata", f"Birleştirme hatası:\n\n{result}")

    # ═══════════════════════════════════════════════════════════════
    # PDF → WORD İŞLEMLERİ
    # ═══════════════════════════════════════════════════════════════
    def _p2w_select(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "PDF Dosyası Seçin", "", "PDF Dosyaları (*.pdf)"
        )
        if f:
            self._p2w_set(f)

    def _p2w_set(self, path):
        self.p2w_selected_file = path
        name = os.path.basename(path)
        size = os.path.getsize(path) / (1024 * 1024)
        self.p2w_label.setText(f"📄  {name}   •   {size:.2f} MB")
        self.p2w_label.setStyleSheet(f"""
            background-color: {Colors.PRIMARY_LIGHT};
            border: 1.5px solid {Colors.PRIMARY};
            border-radius: 10px;
            padding: 14px 16px;
            font-size: 13px;
            color: {Colors.TEXT};
        """)
        self.p2w_convert_btn.setEnabled(True)
        self.p2w_status.setText("✅ Dönüştürmeye hazır")
        self.p2w_progress.setValue(0)

    def _p2w_execute(self):
        if not self.p2w_selected_file:
            return

        base = os.path.splitext(os.path.basename(self.p2w_selected_file))[0]
        output, _ = QFileDialog.getSaveFileName(
            self, "Word Dosyasını Kaydet",
            f"{base}.docx",
            "Word Dosyaları (*.docx)"
        )
        if not output:
            return
        if not output.lower().endswith(".docx"):
            output += ".docx"

        self.p2w_convert_btn.setEnabled(False)
        self.p2w_progress.setValue(0)
        self.p2w_status.setText("⏳ Dönüştürme başlatılıyor...")

        self.current_worker = PdfToWordWorker(self.p2w_selected_file, output)
        self.current_worker.progress.connect(self._p2w_on_progress)
        self.current_worker.finished.connect(self._p2w_on_finished)
        self.current_worker.start()

    def _p2w_on_progress(self, val, text):
        self.p2w_progress.setValue(val)
        self.p2w_status.setText(text)

    def _p2w_on_finished(self, ok, result):
        self.current_worker = None
        self.p2w_convert_btn.setEnabled(True)

        if ok:
            self.p2w_progress.setValue(100)
            self.p2w_status.setText("✅ Dönüştürme tamamlandı!")
            reply = QMessageBox.question(
                self, "Başarılı",
                f"PDF başarıyla Word'e dönüştürüldü!\n\n📝 {os.path.basename(result)}\n\nDosyayı açmak ister misiniz?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._open_file(result)
        else:
            self.p2w_status.setText("❌ Hata oluştu")
            QMessageBox.critical(self, "Hata", f"Dönüştürme hatası:\n\n{result}")

    # ═══════════════════════════════════════════════════════════════
    # WORD → PDF İŞLEMLERİ
    # ═══════════════════════════════════════════════════════════════
    def _w2p_select(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Word Dosyası Seçin", "",
            "Word Dosyaları (*.docx *.doc);;Tüm Dosyalar (*)"
        )
        if f:
            self._w2p_set(f)

    def _w2p_set(self, path):
        self.w2p_selected_file = path
        name = os.path.basename(path)
        size = os.path.getsize(path) / (1024 * 1024)
        self.w2p_label.setText(f"📄  {name}   •   {size:.2f} MB")
        self.w2p_label.setStyleSheet(f"""
            background-color: {Colors.PRIMARY_LIGHT};
            border: 1.5px solid {Colors.PRIMARY};
            border-radius: 10px;
            padding: 14px 16px;
            font-size: 13px;
            color: {Colors.TEXT};
        """)
        self.w2p_convert_btn.setEnabled(True)
        self.w2p_status.setText("✅ Dönüştürmeye hazır")
        self.w2p_progress.setValue(0)

    def _w2p_execute(self):
        if not self.w2p_selected_file:
            return

        base = os.path.splitext(os.path.basename(self.w2p_selected_file))[0]
        output, _ = QFileDialog.getSaveFileName(
            self, "PDF Dosyasını Kaydet",
            f"{base}.pdf",
            "PDF Dosyaları (*.pdf)"
        )
        if not output:
            return
        if not output.lower().endswith(".pdf"):
            output += ".pdf"

        self.w2p_convert_btn.setEnabled(False)
        self.w2p_progress.setValue(0)
        self.w2p_status.setText("⏳ Dönüştürme başlatılıyor...")

        self.current_worker = WordToPdfWorker(self.w2p_selected_file, output)
        self.current_worker.progress.connect(self._w2p_on_progress)
        self.current_worker.finished.connect(self._w2p_on_finished)
        self.current_worker.start()

    def _w2p_on_progress(self, val, text):
        self.w2p_progress.setValue(val)
        self.w2p_status.setText(text)

    def _w2p_on_finished(self, ok, result):
        self.current_worker = None
        self.w2p_convert_btn.setEnabled(True)

        if ok:
            self.w2p_progress.setValue(100)
            self.w2p_status.setText("✅ Dönüştürme tamamlandı!")
            reply = QMessageBox.question(
                self, "Başarılı",
                f"Word başarıyla PDF'ye dönüştürüldü!\n\n📄 {os.path.basename(result)}\n\nDosyayı açmak ister misiniz?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._open_file(result)
        else:
            self.w2p_status.setText("❌ Hata oluştu")
            QMessageBox.critical(self, "Hata", f"Dönüştürme hatası:\n\n{result}")

    # ═══════════════════════════════════════════════════════════════
    # YARDIMCI
    # ═══════════════════════════════════════════════════════════════
    @staticmethod
    def _open_file(path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════
# UYGULAMA GİRİŞ NOKTASI
# ═══════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Uygulama fontu
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Uygulama paleti
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(Colors.BG))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(Colors.TEXT))
    palette.setColor(QPalette.ColorRole.Base, QColor(Colors.CARD))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(Colors.BG))
    palette.setColor(QPalette.ColorRole.Text, QColor(Colors.TEXT))
    palette.setColor(QPalette.ColorRole.Button, QColor(Colors.CARD))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(Colors.TEXT))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(Colors.PRIMARY))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    window = PDFFusionPro()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
