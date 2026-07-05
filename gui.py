# -*- coding: utf-8 -*-
"""
GUI (giao dien) cho cong cu Nhap/Sua hang loat bai dang - DHB Reup Facebook.
Thiet ke: Flat design, teal chu dao + cam CTA (theo ui-ux-pro-max design system).
Dung customtkinter (nut bo goc, hover muot). Tai su dung logic trong import_bai_dang.py.
"""
import os, sys, subprocess, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# --- dam bao cac thu vien can thiet (tu cai neu thieu) ---
def _ensure(mod, pip_name=None):
    try:
        __import__(mod)
        return True
    except Exception:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", pip_name or mod, "--quiet"],
                           check=True)
            __import__(mod)
            return True
        except Exception:
            return False

_ensure("openpyxl")   # dung de doc/ghi Excel
if not _ensure("customtkinter"):
    import tkinter.messagebox as mb
    mb.showerror("Thiếu thư viện",
                 "Cần cài 'customtkinter'. Mở CMD gõ:\n\n"
                 "    pip install customtkinter\n\nrồi chạy lại.")
    raise SystemExit
import customtkinter as ctk

from tkinter import messagebox, StringVar
import import_bai_dang as core

# ================= BANG MAU (design system: Flat, teal + orange) =================
TEAL    = "#0D9488"
TEAL_D  = "#0B7A70"
TEAL_XL = "#E6FBF7"
ORANGE  = "#F97316"
ORANGE_D= "#E4620E"
MINT    = "#F0FDFA"   # nen app
CARD    = "#FFFFFF"
INK     = "#134E4A"   # chu chinh (xanh dam)
MUTED   = "#45635D"   # chu phu
BORDER  = "#D6E7E3"
GREEN_BG, GREEN_TX = "#DCFCE7", "#15803D"
AMBER_BG, AMBER_TX = "#FEF3C7", "#B45309"
FONT = "Segoe UI"


class App:
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("light")
        root.title("DHB — Nhập & Sửa hàng loạt bài đăng")
        root.geometry("900x720")
        root.minsize(820, 620)
        root.configure(fg_color=MINT)

        # fonts (tao sau khi co root)
        self.f_title = ctk.CTkFont(FONT, 20, "bold")
        self.f_h     = ctk.CTkFont(FONT, 15, "bold")
        self.f_body  = ctk.CTkFont(FONT, 13)
        self.f_btn   = ctk.CTkFont(FONT, 13, "bold")
        self.f_muted = ctk.CTkFont(FONT, 12)
        self.f_badge = ctk.CTkFont(FONT, 13, "bold")
        self.f_log   = ctk.CTkFont("Consolas", 11)

        self.proj_var = StringVar()
        self.sua_var  = StringVar()

        self._build_header()
        self._build_log()      # ghim day truoc
        self._build_body()     # khu the cuon o giua

        sys.stdout = self
        sys.stderr = self

        self.pages = self.profiles = {}
        self.groups = {}
        self.ready = False
        self.load_core()

    # ---------- stdout -> log ----------
    def write(self, text):
        try:
            self.log.configure(state="normal")
            self.log.insert("end", text)
            self.log.see("end")
            self.log.configure(state="disabled")
            self.root.update_idletasks()
        except Exception:
            pass
    def flush(self): pass
    def logln(self, text=""): self.write(text + "\n")

    # ---------- widget helpers ----------
    def _btn(self, parent, text, cmd, kind="teal"):
        common = dict(text=text, command=cmd, font=self.f_btn, height=42,
                      corner_radius=9, cursor="hand2")
        if kind == "primary":     # cam - hanh dong chinh
            return ctk.CTkButton(parent, fg_color=ORANGE, hover_color=ORANGE_D,
                                 text_color="white", **common)
        if kind == "teal":        # teal dac
            return ctk.CTkButton(parent, fg_color=TEAL, hover_color=TEAL_D,
                                 text_color="white", **common)
        # outline - tien ich
        return ctk.CTkButton(parent, fg_color="transparent", hover_color=TEAL_XL,
                             text_color=TEAL, border_color=TEAL, border_width=2, **common)

    def _action(self, parent, text, cmd, kind, desc, width=240):
        """1 nut + dong mo ta ngan ben duoi."""
        b = self._btn(parent, text, cmd, kind); b.configure(width=width)
        b.pack(anchor="w", pady=(2, 0))
        ctk.CTkLabel(parent, text=desc, font=self.f_muted, text_color=MUTED,
                     justify="left", wraplength=760, anchor="w").pack(anchor="w", padx=2, pady=(4, 12))

    def _desc(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=self.f_muted, text_color=MUTED,
                     justify="left", wraplength=760, anchor="w").pack(anchor="w", padx=2, pady=(5, 12))

    def _card(self, num, title):
        card = ctk.CTkFrame(self.body, fg_color=CARD, corner_radius=14,
                            border_width=1, border_color=BORDER)
        card.pack(fill="x", pady=(0, 14))
        head = ctk.CTkFrame(card, fg_color=CARD)
        head.pack(fill="x", padx=18, pady=(16, 2))
        badge = ctk.CTkLabel(head, text=str(num), width=30, height=30, corner_radius=15,
                             fg_color=TEAL, text_color="white", font=self.f_badge)
        badge.pack(side="left")
        ctk.CTkLabel(head, text="  " + title, font=self.f_h, text_color=INK).pack(side="left")
        inner = ctk.CTkFrame(card, fg_color=CARD)
        inner.pack(fill="x", padx=18, pady=(6, 16))
        return inner

    # ---------- header ----------
    def _build_header(self):
        h = ctk.CTkFrame(self.root, fg_color=TEAL, corner_radius=0, height=76)
        h.pack(fill="x"); h.pack_propagate(False)
        left = ctk.CTkFrame(h, fg_color=TEAL); left.pack(side="left", padx=22, pady=14)
        ctk.CTkLabel(left, text="Nhập & Sửa hàng loạt bài đăng",
                     font=self.f_title, text_color="white").pack(anchor="w")
        ctk.CTkLabel(left, text=f"Công cụ hỗ trợ DHB Reup Facebook  ·  phiên bản {core.VERSION}",
                     font=self.f_muted, text_color="#CFF5EF").pack(anchor="w")
        self.pill = ctk.CTkLabel(h, text=" đang kiểm tra… ", font=self.f_btn,
                                 corner_radius=13, fg_color="#0B6E66", text_color="white",
                                 height=28)
        self.pill.pack(side="right", padx=22)

    # ---------- body (cuon duoc) ----------
    def _build_body(self):
        self.body = ctk.CTkScrollableFrame(self.root, fg_color=MINT)
        self.body.pack(side="top", fill="both", expand=True, padx=10, pady=(12, 0))

        self.info_lbl = ctk.CTkLabel(self.body, text="Đang tải…", font=self.f_body, text_color=MUTED)
        self.info_lbl.pack(anchor="w", pady=(0, 10), padx=6)

        # (1)
        c1 = self._card(1, "Tạo dự án mới từ Excel")
        self._action(c1, "Mở file Excel mẫu để điền", self.open_template, "teal",
                     "Mở (hoặc tự tạo) file mau_bai_dang.xlsx để bạn điền danh sách bài: "
                     "nguồn, caption, trang đăng, lịch đăng… Điền xong nhớ LƯU lại.", width=250)
        self._action(c1, "Kiểm tra file Excel (chưa ghi gì)", self.action_check, "outline",
                     "Quét trước khi tạo: liệt kê dòng lỗi — vd page dùng chung 2 tài khoản mà "
                     "chưa điền cột ‘Tài khoản’, hoặc tên trang gõ sai. Không ghi gì cả.", width=280)
        self._action(c1, "Tạo / ghi vào dự án", self.action_create, "primary",
                     "Đọc file Excel đã điền và tạo bài vào dự án trong tool. Cột “Dự án” gom bài "
                     "vào dự án cùng tên; nếu trùng tên dự án cũ sẽ hỏi Ghi đè hay Nối thêm.", width=250)

        # (2)
        c2 = self._card(2, "Sửa dự án có sẵn   (xuất → sửa Excel → nạp lại)")
        r2a = ctk.CTkFrame(c2, fg_color=CARD); r2a.pack(fill="x")
        ctk.CTkLabel(r2a, text="Chọn dự án", font=self.f_body, text_color=INK, width=80,
                     anchor="w").pack(side="left")
        self.opt_proj = ctk.CTkOptionMenu(r2a, variable=self.proj_var, values=["(chưa có)"],
                                          font=self.f_body, width=300, height=40, corner_radius=9,
                                          fg_color="white", text_color=INK, button_color=TEAL,
                                          button_hover_color=TEAL_D, dropdown_font=self.f_body)
        self.opt_proj.pack(side="left", padx=10)
        self._btn(r2a, "Xuất ra Excel để sửa", self.action_export, "teal").pack(side="left")
        self._desc(c2, "Tạo file sua_<tên>.xlsx chứa TOÀN BỘ bài của dự án đã chọn, để bạn sửa "
                       "hàng loạt (đổi caption / nguồn / trang / lịch…).")
        r2b = ctk.CTkFrame(c2, fg_color=CARD); r2b.pack(fill="x")
        ctk.CTkLabel(r2b, text="File đã sửa", font=self.f_body, text_color=INK, width=80,
                     anchor="w").pack(side="left")
        self.opt_sua = ctk.CTkOptionMenu(r2b, variable=self.sua_var, values=["(chưa có)"],
                                         font=self.f_body, width=300, height=40, corner_radius=9,
                                         fg_color="white", text_color=INK, button_color=TEAL,
                                         button_hover_color=TEAL_D, dropdown_font=self.f_body)
        self.opt_sua.pack(side="left", padx=10)
        self._btn(r2b, "Nạp lại (ghi đè dự án)", self.action_apply, "primary").pack(side="left")
        self._desc(c2, "Ghi các thay đổi trong file đã sửa trở lại chính dự án đó "
                       "(bản cũ được tự backup trước khi ghi).")

        # (3)
        c3 = self._card(3, "Tiện ích")
        self._action(c3, "Cập nhật danh sách trang / tài khoản", self.action_refresh, "outline",
                     "Khi bạn thêm/xóa fanpage hay tài khoản trong tool, bấm để cập nhật lại danh "
                     "sách trong các file Excel — giữ nguyên dữ liệu bạn đã điền.", width=330)
        self._action(c3, "Tải lại danh sách", self.reload_lists, "outline",
                     "Đọc lại số dự án / trang / file sửa mới nhất để hiển thị trong cửa sổ này.", width=210)
        self._action(c3, "Mở thư mục", lambda: os.startfile(HERE), "outline",
                     "Mở thư mục chứa công cụ, các file Excel (mẫu, đã sửa) và thư mục backup.", width=170)

        self._bind_fast_scroll()

    def _bind_fast_scroll(self):
        """Cuon chuot muot & nhanh hon (mac dinh customtkinter: 1 don vi = 1px -> rat cham).
           Canvas co yscrollincrement=1 nen 1 'unit' = 1px; cuon ~PX_PER_NOTCH px moi nac."""
        try:
            canvas = self.body._parent_canvas
        except Exception:
            return
        PX_PER_NOTCH = 52   # cuon ~52px moi nac (chuot do phan giai cao van muot vi ti le theo delta)
        def _wheel(event):
            step = int(-1 * event.delta / 120 * PX_PER_NOTCH)
            if step == 0:
                step = -1 if event.delta > 0 else 1
            canvas.yview_scroll(step, "units")
            return "break"     # chay truoc handler mac dinh -> khong cong don, khong bi cham
        def _apply(w):
            try: w.bind("<MouseWheel>", _wheel)
            except Exception: pass
            for c in w.winfo_children():
                _apply(c)
        try: canvas.bind("<MouseWheel>", _wheel)
        except Exception: pass
        _apply(self.body)

    # ---------- log (ghim day) ----------
    def _build_log(self):
        wrap = ctk.CTkFrame(self.root, fg_color=MINT)
        wrap.pack(side="bottom", fill="x", padx=18, pady=(4, 14))
        ctk.CTkLabel(wrap, text="Nhật ký", font=self.f_muted, text_color=MUTED).pack(anchor="w")
        self.log = ctk.CTkTextbox(wrap, height=100, font=self.f_log, corner_radius=12, fg_color=CARD,
                                  text_color=INK, border_width=1, border_color=BORDER, wrap="word")
        self.log.pack(fill="x", pady=(4, 0))
        self.log.configure(state="disabled")

    # ================= nghiep vu =================
    def load_core(self):
        try:
            core.check_paths()
            self.pages, self.profiles = core.load_accounts()
            self.groups = core.load_option_groups()
            self.ready = True
        except core.DhbError as e:
            self.ready = False
            self.info_lbl.configure(text="Lỗi đường dẫn tool — sửa biến TOOL_DATA trong import_bai_dang.py")
            messagebox.showerror("Lỗi", str(e)); return
        except Exception as e:
            self.ready = False
            messagebox.showerror("Lỗi", repr(e)); return
        self.reload_lists()
        self.logln(f"Sẵn sàng — phiên bản {core.VERSION}. Chọn một thao tác ở trên.")

    def reload_lists(self):
        if not self.ready: return
        fangroups = core.load_fangroups()
        n_page = sum(len(v) for v in self.pages.values())
        projs = core.list_projects()
        self.info_lbl.configure(
            text=f"{n_page} trang     •     {len(self.profiles)} tài khoản     •     "
                 f"{len(fangroups)} nhóm     •     {len(projs)} dự án")
        pvals = [f"{n}  ({c} bài)" for n, c in projs] or ["(chưa có dự án)"]
        self._proj_names = [n for n, c in projs]
        self.opt_proj.configure(values=pvals)
        if projs and self.proj_var.get() not in pvals:
            self.proj_var.set(pvals[0])
        sua = sorted(f for f in os.listdir(HERE)
                     if f.lower().startswith("sua_") and f.lower().endswith(".xlsx"))
        self.opt_sua.configure(values=sua or ["(chưa có file sửa)"])
        if sua and self.sua_var.get() not in sua:
            self.sua_var.set(sua[0])
        self._update_tool_status()

    def _tool_running(self):
        try:
            core.ensure_tool_closed(); return False
        except core.DhbError:
            return True
        except Exception:
            return False

    def _update_tool_status(self):
        if self._tool_running():
            self.pill.configure(text="  ● Tool đang MỞ — hãy tắt trước khi ghi  ",
                                fg_color=AMBER_BG, text_color=AMBER_TX)
        else:
            self.pill.configure(text="  ● Tool đã tắt — sẵn sàng  ",
                                fg_color=GREEN_BG, text_color=GREEN_TX)

    def _guard_write(self):
        if self._tool_running():
            messagebox.showwarning("Tool đang mở",
                "Tool DHB Reup ĐANG CHẠY.\nHãy tắt hẳn tool rồi thử lại (tránh hỏng dữ liệu).")
            self._update_tool_status()
            return False
        return True

    def _run(self, fn):
        try:
            fn()
        except core.DhbError as e:
            messagebox.showerror("Lỗi", str(e))
        except Exception as e:
            self.logln("LỖI: " + repr(e))
            self.logln(traceback.format_exc())
            messagebox.showerror("Lỗi không mong đợi", repr(e))
        finally:
            self._update_tool_status()

    def _ask_choice(self, title, heading, note_lines, buttons):
        """Hop thoai tuy chinh: cac nut co NHAN RO RANG (khong phai Yes/No/Cancel).
           buttons: list (label, value, kind). Tra ve value da chon, hoac None."""
        dlg = ctk.CTkToplevel(self.root)
        dlg.title(title)
        dlg.configure(fg_color=CARD)
        dlg.resizable(False, False)
        result = {"v": None}
        def choose(v):
            result["v"] = v
            try: dlg.grab_release()
            except Exception: pass
            dlg.destroy()
        self._dlg_choose = choose   # de test/tu dong hoa co the goi
        dlg.protocol("WM_DELETE_WINDOW", lambda: choose(None))

        ctk.CTkLabel(dlg, text="  Lưu ý  ", font=self.f_btn, text_color=AMBER_TX,
                     fg_color=AMBER_BG, corner_radius=8, height=26).pack(anchor="w", padx=22, pady=(18, 8))
        ctk.CTkLabel(dlg, text=heading, font=self.f_h, text_color=INK,
                     wraplength=430, justify="left").pack(anchor="w", padx=22, pady=(0, 10))
        nb = ctk.CTkFrame(dlg, fg_color="#FFF8EC", corner_radius=10, border_width=1, border_color="#F2D9A8")
        nb.pack(fill="x", padx=22, pady=(0, 16))
        for i, ln in enumerate(note_lines):
            ctk.CTkLabel(nb, text=ln, font=self.f_muted, text_color="#8A5A00",
                         wraplength=400, justify="left", anchor="w").pack(
                         anchor="w", padx=12,
                         pady=(10 if i == 0 else 3, 10 if i == len(note_lines) - 1 else 0))
        br = ctk.CTkFrame(dlg, fg_color=CARD); br.pack(fill="x", padx=22, pady=(0, 18))
        for (label, val, kind) in buttons:   # pack side=right: phan tu dau list -> ngoai cung ben phai
            b = self._btn(br, label, (lambda vv=val: choose(vv)), kind)
            b.configure(width=120)
            b.pack(side="right", padx=(8, 0))
        dlg.transient(self.root); dlg.update_idletasks()
        w = 480; h = max(250, dlg.winfo_reqheight())
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        x = max(px + (pw - w) // 2, 0); y = max(py + (ph - h) // 2, 0)
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.lift(); dlg.grab_set(); dlg.wait_window()
        return result["v"]

    # --- hanh dong ---
    def open_template(self):
        def go():
            if not os.path.exists(core.EXCEL):
                core.make_template_excel(core.EXCEL, self.pages, self.profiles, self.groups)
                self.logln("Chưa có mẫu → đã tạo mau_bai_dang.xlsx")
            os.startfile(core.EXCEL)
            self.logln("Đã mở file mẫu. Điền xong nhớ LƯU lại.")
        self._run(go)

    def action_create(self):
        def go():
            if not self._guard_write(): return
            if not os.path.exists(core.EXCEL):
                core.make_template_excel(core.EXCEL, self.pages, self.profiles, self.groups)
                self.logln("Chưa có mẫu → đã tạo mau_bai_dang.xlsx. Hãy điền rồi bấm lại.")
                os.startfile(core.EXCEL); return
            posts_raw = core.read_excel(core.EXCEL)
            if not posts_raw:
                messagebox.showinfo("Trống", "File Excel chưa có dòng nào."); return
            # Chỉ hỏi 'tên dự án mặc định' khi CÓ dòng để TRỐNG cột 'Dự án'.
            need_default = any(not str(r.get("project") or "").strip() for r in posts_raw)
            project = "NhapHangLoat"
            if need_default:
                dlg = ctk.CTkInputDialog(
                    title="Tên dự án mặc định",
                    text="Có bài chưa ghi cột ‘Dự án’. Nhập tên dự án mặc định cho những bài đó\n"
                         "(để trống = NhapHangLoat):")
                project = dlg.get_input()
                if project is None: return
                project = project.strip() or "NhapHangLoat"
            tmpl, generic = core.build_templates(core.harvest_posts())
            fangroups = core.load_fangroups()
            by_proj, errors = {}, []
            for i, row in enumerate(posts_raw):
                pj = str(row.get("project") or "").strip() or project
                try:
                    lst = by_proj.setdefault(pj, [])
                    lst.append(core.build_post(row, len(lst), self.pages, self.profiles,
                                               fangroups, self.groups, tmpl, generic))
                except Exception as e:
                    errors.append(f"Dòng {i+2}: {e}")
            if errors:
                self.logln("Bỏ qua dòng lỗi:")
                for e in errors: self.logln("   " + e)
            if not by_proj:
                messagebox.showwarning("Không có bài", "Không có bài hợp lệ để nhập."); return
            if errors and not messagebox.askyesno(
                    "Có dòng cần sửa",
                    f"Có {len(errors)} dòng lỗi sẽ bị BỎ QUA (xem Nhật ký) — "
                    "hay gặp: page dùng chung mà chưa chọn cột ‘Tài khoản’.\n\n"
                    "Vẫn ghi các dòng hợp lệ còn lại?\n"
                    "(Chọn No để huỷ, sửa file Excel rồi làm lại.)"):
                self.logln("Đã huỷ để sửa file Excel."); return

            def exists(pj):
                p = os.path.join(core.PROJECT_DIR, pj + ".dhb")
                return os.path.exists(p) and os.path.getsize(p) > 0
            trung = [pj for pj in by_proj if exists(pj)]
            append = False
            if trung:
                choice = self._ask_choice(
                    "Trùng tên dự án",
                    "Dự án đã có sẵn: " + ", ".join(trung) + ".\nBạn muốn xử lý thế nào?",
                    ["•  Ghi đè — xóa hết bài cũ, thay bằng bài mới (bản cũ được tự backup).",
                     "•  Nối thêm — giữ nguyên bài cũ, chỉ thêm bài mới vào cuối."],
                    [("Ghi đè", "overwrite", "primary"),
                     ("Nối thêm", "append", "teal"),
                     ("Huỷ", None, "outline")])
                if choice is None: self.logln("Đã huỷ."); return
                append = (choice == "append")

            total = 0
            for pj, lst in by_proj.items():
                core.write_project(pj, lst, append)
                tag = (("nối thêm" if append else "ghi đè") if pj in trung else "mới")
                self.logln(f"   • {pj}: {len(lst)} bài  ({tag})")
                total += len(lst)
            self.logln(f"HOÀN TẤT: {total} bài vào {len(by_proj)} dự án. Mở tool → load dự án → chạy.")
            self.reload_lists()
            messagebox.showinfo("Xong", f"Đã nhập {total} bài vào {len(by_proj)} dự án.")
        self._run(go)

    def action_check(self):
        """Quét file Excel, liệt kê dòng lỗi (KHÔNG ghi gì) — vd page dùng chung chưa chọn tài khoản."""
        def go():
            if not os.path.exists(core.EXCEL):
                messagebox.showinfo("Chưa có file",
                    "Chưa có file mau_bai_dang.xlsx.\nBấm ‘Mở file Excel mẫu để điền’ trước."); return
            posts_raw = core.read_excel(core.EXCEL)
            if not posts_raw:
                messagebox.showinfo("Trống", "File Excel chưa có dòng nào."); return
            tmpl, generic = core.build_templates(core.harvest_posts())
            fangroups = core.load_fangroups()
            ok, problems = 0, []
            for i, row in enumerate(posts_raw):
                try:
                    core.build_post(row, i, self.pages, self.profiles, fangroups,
                                    self.groups, tmpl, generic)
                    ok += 1
                except Exception as e:
                    problems.append(f"   Dòng {i+2}: {e}")
            self.logln(f"── Kiểm tra Excel: {len(posts_raw)} dòng → {ok} OK, {len(problems)} cần sửa ──")
            for p in problems:
                self.logln(p)
            if problems:
                messagebox.showwarning("Có dòng cần sửa",
                    f"{len(problems)}/{len(posts_raw)} dòng có vấn đề (chi tiết ở Nhật ký).\n\n"
                    "Hay gặp nhất: page dùng chung 2 tài khoản (Sheepie.decor / Bedding Decor) "
                    "mà chưa điền cột ‘Tài khoản’.")
            else:
                messagebox.showinfo("Tốt", f"Tất cả {ok} dòng hợp lệ — sẵn sàng tạo dự án.")
        self._run(go)

    def action_export(self):
        def go():
            if not self._proj_names:
                messagebox.showinfo("Chưa có dự án", "Tool chưa có dự án nào để xuất."); return
            vals = list(self.opt_proj.cget("values"))
            sel = self.proj_var.get()
            i = vals.index(sel) if sel in vals else 0
            if i >= len(self._proj_names): i = 0
            name = self._proj_names[i]
            out, n = core.export_project(name, self.groups)
            self.logln(f"Đã xuất {n} bài → {os.path.basename(out)}")
            self.reload_lists()
            self.sua_var.set(os.path.basename(out))
            if messagebox.askyesno("Mở file?",
                    f"Đã xuất {n} bài ra:\n{os.path.basename(out)}\n\nMở file để sửa ngay?"):
                os.startfile(out)
        self._run(go)

    def action_apply(self):
        def go():
            if not self._guard_write(): return
            fname = self.sua_var.get()
            if not fname or fname.startswith("("):
                messagebox.showinfo("Chọn file", "Hãy chọn 1 file sửa để nạp."); return
            name = fname[4:-5]
            choice = self._ask_choice(
                "Xác nhận nạp lại",
                f"Ghi đè toàn bộ dự án “{name}” bằng nội dung file {fname}?",
                ["•  Bài cũ trong dự án sẽ được thay hết bằng bài trong file.",
                 "•  Bản cũ được tự backup trước khi ghi (khôi phục được)."],
                [("Ghi đè", "ok", "primary"), ("Huỷ", None, "outline")])
            if choice != "ok": return
            target, n, n_run, nerr = core.apply_project(name, self.pages, self.profiles, self.groups, append=False)
            self.logln(f"Đã ghi {n} bài vào “{name}” ({n_run} bài sẽ chạy). Lỗi bỏ qua: {nerr}")
            messagebox.showinfo("Xong", f"Đã nạp {n} bài vào “{name}”.\n{n_run} bài sẽ chạy (Not start).")
            self.reload_lists()
        self._run(go)

    def action_refresh(self):
        def go():
            core.do_refresh(self.groups)
            self.reload_lists()
            messagebox.showinfo("Xong", "Đã cập nhật danh sách trang/tài khoản trong các file Excel.")
        self._run(go)


def main():
    root = ctk.CTk()
    try:
        App(root)
    except Exception:
        err = traceback.format_exc()
        try:
            with open(os.path.join(HERE, "gui_error.log"), "w", encoding="utf-8") as f:
                f.write(err)
        except Exception: pass
        try: messagebox.showerror("Lỗi khởi động", err)
        except Exception: pass
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
