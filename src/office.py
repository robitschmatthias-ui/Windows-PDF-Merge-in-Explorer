"""Office documents to PDF: Microsoft Office via COM, LibreOffice as fallback.

All methods must run in the thread that called pythoncom.CoInitialize()
(COM objects are bound to their thread).
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
from pathlib import Path

import pywintypes
import win32com.client

log = logging.getLogger(__name__)

WORD_EXT = {".doc", ".docx", ".odt", ".rtf", ".txt"}
EXCEL_EXT = {".xlsx", ".ods"}
POWERPOINT_EXT = {".pptx", ".odp"}
OFFICE_EXT = WORD_EXT | EXCEL_EXT | POWERPOINT_EXT

# Office constants
WD_ALERTS_NONE = 0
WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # never run macros when opening

# Passed as password so protected documents fail with an error instead of
# waiting invisibly for a password prompt. Office ignores it for documents
# without a password.
DUMMY_PASSWORD = "#pdf-merge-no-password#"


class OfficeError(Exception):
    """Error with a message meant for the user."""


def _com_message(exc: Exception) -> str:
    if isinstance(exc, pywintypes.com_error):
        excepinfo = exc.args[2] if len(exc.args) > 2 else None
        if excepinfo and excepinfo[2]:
            return str(excepinfo[2]).strip()
        return str(exc.args[1]) if len(exc.args) > 1 else str(exc)
    return str(exc)


def find_soffice() -> Path | None:
    found = shutil.which("soffice")
    if found:
        return Path(found)
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            candidate = Path(base) / "LibreOffice" / "program" / "soffice.exe"
            if candidate.exists():
                return candidate
    return None


class OfficeConverter:
    """Keeps the Office applications open across files and quits them only
    in close() - Word can take a long time to quit."""

    def __init__(self, workdir: Path):
        self.workdir = workdir
        self._apps: dict[str, object] = {}
        self._missing: set[str] = set()

    # ── Microsoft Office ────────────────────────────────────────────────────

    def _app(self, prog_id: str):
        if prog_id in self._missing:
            return None
        if prog_id not in self._apps:
            try:
                # DispatchEx starts a separate instance so the user's open
                # documents stay untouched (PowerPoint is the exception,
                # see close()).
                app = win32com.client.DispatchEx(prog_id)
            except pywintypes.com_error:
                log.info("%s not available", prog_id)
                self._missing.add(prog_id)
                return None
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                log.warning("%s: cannot set AutomationSecurity", prog_id)
            if prog_id == "Word.Application":
                app.Visible = False
                app.DisplayAlerts = WD_ALERTS_NONE
            elif prog_id == "Excel.Application":
                app.Visible = False
                app.DisplayAlerts = False
                app.ScreenUpdating = False
                app.EnableEvents = False
                app.AskToUpdateLinks = False
            self._apps[prog_id] = app
        return self._apps[prog_id]

    def _word(self, src: Path, dst: Path) -> bool:
        app = self._app("Word.Application")
        if app is None:
            return False
        doc = app.Documents.Open(
            FileName=str(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            PasswordDocument=DUMMY_PASSWORD,
            Visible=False,
            NoEncodingDialog=True,
        )
        try:
            doc.ExportAsFixedFormat(OutputFileName=str(dst), ExportFormat=WD_EXPORT_FORMAT_PDF)
        finally:
            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        return True

    def _excel(self, src: Path, dst: Path) -> bool:
        app = self._app("Excel.Application")
        if app is None:
            return False
        wb = app.Workbooks.Open(
            Filename=str(src),
            UpdateLinks=0,
            ReadOnly=True,
            Password=DUMMY_PASSWORD,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        try:
            wb.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=str(dst))
        finally:
            wb.Close(SaveChanges=False)
        return True

    def _powerpoint(self, src: Path, dst: Path) -> bool:
        app = self._app("PowerPoint.Application")
        if app is None:
            return False
        pres = app.Presentations.Open(
            FileName=str(src), ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            pres.SaveAs(str(dst), PP_SAVE_AS_PDF)
        finally:
            pres.Close()
        return True

    def _ms_office(self, src: Path, dst: Path) -> bool:
        ext = src.suffix.lower()
        if ext in WORD_EXT:
            return self._word(src, dst)
        if ext in EXCEL_EXT:
            return self._excel(src, dst)
        if ext in POWERPOINT_EXT:
            return self._powerpoint(src, dst)
        return False

    # ── LibreOffice ─────────────────────────────────────────────────────────

    def _libreoffice(self, src: Path, dst: Path) -> bool:
        soffice = find_soffice()
        if soffice is None:
            return False
        outdir = self.workdir / "lo_out"
        outdir.mkdir(exist_ok=True)
        # Separate profile: if LibreOffice is already open, the headless call
        # would otherwise attach to the running instance and produce nothing.
        profile = (self.workdir / "lo_profile").as_uri()
        subprocess.run(
            [str(soffice), f"-env:UserInstallation={profile}", "--headless",
             "--norestore", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
            check=False,
        )
        produced = outdir / f"{src.stem}.pdf"
        if not produced.exists():
            return False
        os.replace(produced, dst)
        return True

    # ── Public ──────────────────────────────────────────────────────────────

    def convert(self, src: Path, dst: Path) -> None:
        """Converts src to dst (PDF). Raises OfficeError on failure."""
        ms_error = None
        try:
            if self._ms_office(src, dst) and dst.exists():
                return
        except Exception as exc:
            ms_error = _com_message(exc)
            log.warning("MS Office could not convert %s: %s", src.name, ms_error)

        try:
            if self._libreoffice(src, dst):
                return
        except Exception:
            log.exception("LibreOffice conversion of %s failed", src.name)

        if ms_error:
            raise OfficeError(ms_error)
        if not self._apps and find_soffice() is None:
            raise OfficeError("neither Microsoft Office nor LibreOffice found")
        raise OfficeError("conversion failed")

    def close(self) -> None:
        """Quits the started Office applications - but only if none of the
        user's documents are open in them (PowerPoint shares its instance)."""
        collections = {
            "Word.Application": "Documents",
            "Excel.Application": "Workbooks",
            "PowerPoint.Application": "Presentations",
        }
        for prog_id, app in self._apps.items():
            try:
                if getattr(app, collections[prog_id]).Count == 0:
                    app.Quit()
                else:
                    log.info("%s left running: user documents are open", prog_id)
            except Exception:
                log.warning("%s did not quit cleanly", prog_id)
        self._apps.clear()
