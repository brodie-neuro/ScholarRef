#!/usr/bin/env python3
"""ScholarRef desktop GUI."""

from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document

import scholarref
import verify_reference_integrity


MODE_LABEL_TO_KEY = {
    "APA 7 -> Vancouver": "a2v",
    "Harvard -> Vancouver": "h2v",
    "Vancouver -> APA 7": "v2a",
    "Vancouver -> Harvard": "v2h",
    "APA 7 -> Harvard": "a2h",
    "Harvard -> APA 7": "h2a",
}


def run_conversion_job(
    in_path: Path,
    out_path: Path,
    mode_key: str,
    *,
    keep_uncited: bool = True,
    sort_apa: bool = True,
    run_verify: bool = True,
    verify_profile: str = "references-only",
    ref_header_n: int = 1,
    allow_field_codes: bool = False,
    allow_unsupported_parts: bool = False,
    logger=None,
) -> dict:
    """Run one conversion+save(+optional verify) job used by GUI and tests."""
    if logger is None:
        logger = lambda _msg: None

    doc = Document(str(in_path))
    preflight = scholarref.preflight_docx(
        doc,
        ref_header_n=ref_header_n,
        allow_field_codes=allow_field_codes,
        allow_unsupported_parts=allow_unsupported_parts,
    )
    for w in preflight["warnings"]:
        logger(f"[WARN] {w}")
    if preflight["failures"]:
        raise RuntimeError("Preflight failed: " + " | ".join(preflight["failures"]))

    if mode_key in {"a2v", "h2v"}:
        stats = scholarref.convert_author_date_to_vancouver(
            doc,
            keep_uncited=keep_uncited,
            ref_header_n=ref_header_n,
        )
        logger("Conversion stats:")
        logger(f"- unique citations in body: {stats['body_unique_citations']}")
        logger(f"- parenthetical replacements: {stats['paren_replacements']}")
        logger(f"- narrative replacements: {stats['narr_replacements']}")
        logger(f"- header/footer paragraphs scanned: {stats['header_footer_paragraphs']}")
        logger(f"- references written: {stats['reference_count']}")
    elif mode_key in {"v2a", "v2h"}:
        target_style = "apa7" if mode_key == "v2a" else "harvard"
        stats = scholarref.convert_vancouver_to_author_date(
            doc,
            target_style=target_style,
            sort_references=sort_apa,
            ref_header_n=ref_header_n,
        )
        logger("Conversion stats:")
        logger(f"- target style: {target_style}")
        logger(f"- citation replacements: {stats['citation_replacements']}")
        logger(f"- header/footer paragraphs scanned: {stats['header_footer_paragraphs']}")
        logger(f"- references written: {stats['reference_count']}")
    elif mode_key in {"a2h", "h2a"}:
        target_style = "harvard" if mode_key == "a2h" else "apa7"
        stats = scholarref.convert_author_date_to_author_date(
            doc,
            target_style=target_style,
            ref_header_n=ref_header_n,
        )
        logger("Conversion stats:")
        logger(f"- target style: {target_style}")
        logger(f"- citation restyles: {stats['citation_restyles']}")
        logger(f"- header/footer paragraphs scanned: {stats['header_footer_paragraphs']}")
        logger(f"- references written: {stats['reference_count']}")
    else:
        raise ValueError(f"Unsupported mode_key: {mode_key}")

    doc.save(str(out_path))
    logger(f"Saved: {out_path}")

    verify_rc = None
    if run_verify and mode_key in {"a2v", "h2v"}:
        logger(f"Running verification (profile={verify_profile})...")
        verify_rc = verify_reference_integrity.verify(
            source_path=str(in_path),
            output_path=str(out_path),
            profile=verify_profile,
        )
        logger("Verification: PASS" if verify_rc == 0 else "Verification: FAIL")
    elif run_verify:
        logger("Verification skipped: verifier is defined for Vancouver numeric outputs.")

    return {
        "mode": mode_key,
        "stats": stats,
        "output_path": str(out_path),
        "verify_rc": verify_rc,
    }


import customtkinter as ctk
from PIL import Image
from CTkToolTip import CTkToolTip

class ScholarRefApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        
        # Configure app window
        self.title("ScholarRef")
        self.geometry("980x720")
        self.minsize(880, 680)
        
        # 2026 modern dark aesthetic
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green") # we override with teal anyway
        
        try:
            logo_path = Path(__file__).parent / "logo" / "logo.png"
            if tk.TkVersion >= 8.6:
                img = tk.PhotoImage(file=str(logo_path))
                self.iconphoto(False, img)
            else:
                self.iconbitmap(str(logo_path))
        except Exception:
            pass

        self._init_state()
        self._build_ui()
        self._on_mode_changed()

    def _init_state(self) -> None:
        self.mode_var = ctk.StringVar(value="APA 7 -> Vancouver")
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar()
        self.keep_uncited_var = ctk.BooleanVar(value=True)
        self.sort_apa_var = ctk.BooleanVar(value=True)
        self.run_verify_var = ctk.BooleanVar(value=True)
        self.ref_header_n_var = ctk.StringVar(value="1")
        self.allow_field_codes_var = ctk.BooleanVar(value=False)
        self.allow_unsupported_parts_var = ctk.BooleanVar(value=False)
        self.status_var = ctk.StringVar(value="Ready.")
        self._busy = False

    def _build_ui(self) -> None:
        # Layout weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Header Banner (Glass feel) ──
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        
        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(pady=20, padx=40, fill="x")

        title_box = ctk.CTkFrame(header_content, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(title_box, text="ScholarRef", text_color="#ffffff", font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_box, text="Format. Convert. Publish.", text_color="#0ea5a4", font=ctk.CTkFont(family="Segoe UI", size=14)).pack(anchor="w", pady=(0, 0))

        # ── Main Content Area ──
        root_container = ctk.CTkFrame(self, fg_color="transparent")
        root_container.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        root_container.grid_columnconfigure(0, weight=1)
        root_container.grid_rowconfigure(1, weight=1)

        # Config Card
        card = ctk.CTkFrame(root_container, fg_color="#1e293b", corner_radius=12)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        for idx in range(5):
            card.grid_columnconfigure(idx, weight=1 if idx in (1, 2, 3) else 0)

        # Controls spacing configuration
        ctrl_padx = 20
        ctrl_pady = 12

        # Row 0: Mode
        mode_lbl = ctk.CTkLabel(card, text="Conversion Mode", text_color="#cbd5e1", font=ctk.CTkFont(weight="bold"))
        mode_lbl.grid(row=0, column=0, sticky="w", padx=ctrl_padx, pady=(20, ctrl_pady))
        CTkToolTip(mode_lbl, message="Select which citation style to convert from and to.")
        
        mode_combo = ctk.CTkOptionMenu(
            card,
            variable=self.mode_var,
            values=list(MODE_LABEL_TO_KEY.keys()),
            fg_color="#0f172a",
            button_color="#0ea5a4",
            button_hover_color="#0b8f8e",
            font=ctk.CTkFont(size=13),
            dropdown_font=ctk.CTkFont(size=13),
            width=250,
            command=lambda _e: self._on_mode_changed()
        )
        mode_combo.grid(row=0, column=1, sticky="w", pady=(20, ctrl_pady))
        CTkToolTip(mode_combo, message="Change the conversion direction (e.g. APA 7 to Vancouver).")

        # Row 1: Input DOCX
        in_lbl = ctk.CTkLabel(card, text="Input DOCX", text_color="#cbd5e1", font=ctk.CTkFont(weight="bold"))
        in_lbl.grid(row=1, column=0, sticky="w", padx=ctrl_padx, pady=ctrl_pady)
        CTkToolTip(in_lbl, message="The original Microsoft Word document you want to convert.")
        
        input_entry = ctk.CTkEntry(card, textvariable=self.input_var, fg_color="#0f172a", border_color="#334155", text_color="white", height=36)
        input_entry.grid(row=1, column=1, columnspan=3, sticky="ew", pady=ctrl_pady, padx=(0, 15))
        CTkToolTip(input_entry, message="Path to your input .docx file.")
        
        in_btn = ctk.CTkButton(card, text="Browse", width=100, height=36, fg_color="#334155", hover_color="#475569", command=self._browse_input)
        in_btn.grid(row=1, column=4, padx=ctrl_padx, pady=ctrl_pady, sticky="e")
        CTkToolTip(in_btn, message="Open a file dialog to select your input document.")

        # Row 2: Output DOCX (Fully editable)
        out_lbl = ctk.CTkLabel(card, text="Output Directory/Name", text_color="#cbd5e1", font=ctk.CTkFont(weight="bold"))
        out_lbl.grid(row=2, column=0, sticky="w", padx=ctrl_padx, pady=ctrl_pady)
        CTkToolTip(out_lbl, message="Where the successfully converted document will be saved.")
        
        output_entry = ctk.CTkEntry(card, textvariable=self.output_var, fg_color="#0f172a", border_color="#334155", text_color="white", height=36)
        output_entry.grid(row=2, column=1, columnspan=3, sticky="ew", pady=ctrl_pady, padx=(0, 15))
        CTkToolTip(output_entry, message="Path where the new document will be written.")
        
        out_btn = ctk.CTkButton(card, text="Browse", width=100, height=36, fg_color="#334155", hover_color="#475569", command=self._browse_output)
        out_btn.grid(row=2, column=4, padx=ctrl_padx, pady=ctrl_pady, sticky="e")
        CTkToolTip(out_btn, message="Open a file dialog to choose where to save the output.")

        # Row 3: Options Frame
        options = ctk.CTkFrame(card, fg_color="transparent")
        options.grid(row=3, column=0, columnspan=5, sticky="ew", padx=ctrl_padx, pady=(15, 5))
        options.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.keep_uncited_chk = ctk.CTkCheckBox(options, text="Keep uncited references", variable=self.keep_uncited_var, fg_color="#0ea5a4", text_color="#cbd5e1")
        self.keep_uncited_chk.grid(row=0, column=0, sticky="w")
        CTkToolTip(self.keep_uncited_chk, message="If checked, references that aren't cleanly cited in the text will still be kept in the bibliography.")
        
        self.sort_apa_chk = ctk.CTkCheckBox(options, text="Sort alphabetically", variable=self.sort_apa_var, fg_color="#0ea5a4", text_color="#cbd5e1")
        self.sort_apa_chk.grid(row=0, column=1, sticky="w", padx=(10, 0))
        CTkToolTip(self.sort_apa_chk, message="Automatically alphabetize the bibliography (crucial for Author-Date formats like APA/Harvard).")
        
        # User requested to remove the references-only combo, leaving just the verify checkbox.
        self.verify_chk = ctk.CTkCheckBox(options, text="Verify integrity", variable=self.run_verify_var, fg_color="#0ea5a4", text_color="#cbd5e1")
        self.verify_chk.grid(row=0, column=2, sticky="w", padx=(10, 0))
        CTkToolTip(self.verify_chk, message="Run an aggressive pass over the document at the end to guarantee every parenthetical citation was found and replaced.")

        ref_lbl = ctk.CTkLabel(options, text="Ref header #", text_color="#cbd5e1")
        ref_lbl.grid(row=1, column=0, sticky="w", pady=(10, 0))
        CTkToolTip(ref_lbl, message="Which numeric header defines your bibliography Section? Usually 1 (for simply 'References').")
        
        ref_entry = ctk.CTkEntry(
            options,
            textvariable=self.ref_header_n_var,
            width=70,
            fg_color="#0f172a",
            border_color="#334155",
            text_color="white",
        )
        ref_entry.grid(row=1, column=0, sticky="e", pady=(10, 0), padx=(0, 40))
        CTkToolTip(ref_entry, message="Header order index for your bibliography list.")

        self.allow_fields_chk = ctk.CTkCheckBox(
            options,
            text="Allow field codes",
            variable=self.allow_field_codes_var,
            fg_color="#0ea5a4",
            text_color="#cbd5e1",
        )
        self.allow_fields_chk.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        CTkToolTip(self.allow_fields_chk, message="Permit the script to run even if Word field codes (like original EndNote/Mendeley links) are present. Often requires saving as flat text first.")

        self.allow_unsupported_chk = ctk.CTkCheckBox(
            options,
            text="Allow unsupported parts",
            variable=self.allow_unsupported_parts_var,
            fg_color="#0ea5a4",
            text_color="#cbd5e1",
        )
        self.allow_unsupported_chk.grid(row=1, column=2, sticky="w", padx=(10, 0), pady=(10, 0))
        CTkToolTip(self.allow_unsupported_chk, message="Force the logic to run even if the document contains weird, unparseable elements like embedded OLE objects.")

        # Row 4: Action Buttons
        button_row = ctk.CTkFrame(card, fg_color="transparent")
        button_row.grid(row=4, column=0, columnspan=5, sticky="ew", padx=ctrl_padx, pady=(20, 20))
        button_row.grid_columnconfigure(0, weight=1)
        
        sug_btn = ctk.CTkButton(button_row, text="Suggest Output Name", width=160, height=40, fg_color="#334155", hover_color="#475569", font=ctk.CTkFont(weight="bold"), command=self._suggest_output)
        sug_btn.grid(row=0, column=1, sticky="e", padx=(0, 15))
        CTkToolTip(sug_btn, message="Automatically generate an output filename based on your input (e.g., Input_vancouver.docx).")
        
        self.convert_btn = ctk.CTkButton(button_row, text="Run Conversion", width=160, height=40, fg_color="#0ea5a4", hover_color="#0b8f8e", font=ctk.CTkFont(weight="bold"), command=self._start_conversion)
        self.convert_btn.grid(row=0, column=2, sticky="e")
        CTkToolTip(self.convert_btn, message="Start the sub-second rewrite of your manuscript based on your chosen settings!")

        # ── Log Area ──
        log_card = ctk.CTkFrame(root_container, fg_color="#1e293b", corner_radius=12)
        log_card.grid(row=1, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        
        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))
        ctk.CTkLabel(log_header, text="Activity Log", text_color="#cbd5e1", font=ctk.CTkFont(weight="bold")).pack(side="left")

        self.log_text = ctk.CTkTextbox(
            log_card,
            fg_color="#0f172a",
            text_color="#e2e8f0",
            corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=13),
            wrap="word",
            border_spacing=10
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))

        # ── Footer ──
        footer = ctk.CTkFrame(root_container, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        footer.grid_columnconfigure(0, weight=1)
        
        self.progress = ctk.CTkProgressBar(footer, progress_color="#0ea5a4", fg_color="#334155", height=8)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 20))
        self.progress.set(0) # start at 0
        
        ctk.CTkLabel(footer, textvariable=self.status_var, text_color="#94a3b8", font=ctk.CTkFont(size=12)).grid(row=0, column=1, sticky="e")

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input DOCX",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)
        if not self.output_var.get():
            self._suggest_output()

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save output DOCX as",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.output_var.set(path)

    def _suggest_output(self) -> None:
        raw_input = self.input_var.get().strip()
        if not raw_input:
            return
        in_path = Path(raw_input)
        mode_key = MODE_LABEL_TO_KEY[self.mode_var.get()]
        if mode_key in {"a2v", "h2v"}:
            suffix = "_vancouver.docx"
        elif mode_key in {"v2a", "h2a"}:
            suffix = "_apa7.docx"
        else:
            suffix = "_harvard.docx"
        self.output_var.set(str(in_path.with_name(f"{in_path.stem}{suffix}")))

    def _on_mode_changed(self) -> None:
        mode_key = MODE_LABEL_TO_KEY[self.mode_var.get()]
        
        if mode_key in {"a2v", "h2v"}:
            self.keep_uncited_chk.configure(state="normal")
        else:
            self.keep_uncited_chk.deselect()
            self.keep_uncited_chk.configure(state="disabled")
            
        if mode_key in {"v2a", "v2h"}:
            self.sort_apa_chk.configure(state="normal")
        else:
            self.sort_apa_chk.deselect()
            self.sort_apa_chk.configure(state="disabled")

        verify_allowed = mode_key in {"a2v", "h2v"}
        if verify_allowed:
            self.verify_chk.configure(state="normal")
        else:
            self.verify_chk.deselect()
            self.verify_chk.configure(state="disabled")

        if not self.output_var.get().strip():
            self._suggest_output()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            self.convert_btn.configure(state="disabled")
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        else:
            self.convert_btn.configure(state="normal")
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(0)

    def _log(self, msg: str) -> None:
        if threading.current_thread() is not threading.main_thread():
            self.after(0, self._log, msg)
            return
        
        # CTkTextbox is read-only by default sometimes, we need to toggle state
        self.log_text.insert("end", msg.rstrip() + "\n")
        self.log_text.see("end")

    def _set_status(self, msg: str) -> None:
        if threading.current_thread() is not threading.main_thread():
            self.after(0, self._set_status, msg)
            return
        self.status_var.set(msg)

    def _start_conversion(self) -> None:
        if self._busy:
            return
        in_path = Path(self.input_var.get().strip())
        out_path = Path(self.output_var.get().strip())
        if not in_path.exists() or not in_path.is_file():
            messagebox.showerror("ScholarRef", "Input file not found.")
            return
        if not out_path.suffix.lower() == ".docx":
            messagebox.showerror("ScholarRef", "Output must be a .docx file.")
            return
        self._set_busy(True)
        self._set_status("Running conversion...")
        
        # Clear log
        self.log_text.delete("1.0", "end")
        
        self._log("=" * 72)
        self._log(f"Input:  {in_path}")
        self._log(f"Output: {out_path}")
        self._log(f"Mode:   {self.mode_var.get()}")
        self._log(f"Ref header #: {self.ref_header_n_var.get().strip() or '1'}")
        thread = threading.Thread(target=self._run_conversion_worker, args=(in_path, out_path), daemon=True)
        thread.start()

    def _run_conversion_worker(self, in_path: Path, out_path: Path) -> None:
        try:
            mode_key = MODE_LABEL_TO_KEY[self.mode_var.get()]
            profile = "references-only" # Hardcoded since box was removed
            run_verify = self.run_verify_var.get() and mode_key in {"a2v", "h2v"}
            ref_header_n_raw = (self.ref_header_n_var.get() or "1").strip()
            ref_header_n = int(ref_header_n_raw)
            if ref_header_n < 1:
                raise ValueError("Ref header # must be 1 or greater.")
            run_conversion_job(
                in_path=in_path,
                out_path=out_path,
                mode_key=mode_key,
                keep_uncited=self.keep_uncited_var.get(),
                sort_apa=self.sort_apa_var.get(),
                run_verify=run_verify,
                verify_profile=profile,
                ref_header_n=ref_header_n,
                allow_field_codes=self.allow_field_codes_var.get(),
                allow_unsupported_parts=self.allow_unsupported_parts_var.get(),
                logger=self._log,
            )

            self._set_status("Complete.")
            self.after(0, lambda: messagebox.showinfo("ScholarRef", "Conversion completed successfully."))
        except Exception as exc:  # pragma: no cover
            self._log("ERROR: conversion failed")
            self._log(str(exc))
            self._log(traceback.format_exc())
            self._set_status("Failed.")
            self.after(0, lambda: messagebox.showerror("ScholarRef", f"Conversion failed:\n{exc}"))
        finally:
            self.after(0, self._set_busy, False)


def main() -> int:
    app = ScholarRefApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
