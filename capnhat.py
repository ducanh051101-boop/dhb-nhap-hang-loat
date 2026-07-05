# -*- coding: utf-8 -*-
"""
Tu dong cap nhat code tu GitHub (fail-safe).
- Chi cap nhat cac FILE CODE, KHONG dong toi du lieu (Excel, backup, .dhb...).
- Loi mang / tai loi / file hong -> BO QUA, chay ban hien tai (khong bao gio lam hong tool).
Chay: nhay dup Chay_GUI.bat se tu goi file nay truoc khi mo giao dien.
"""
import os, re, sys, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))

# Repo cong khai chua ban moi nhat (chi la code, khong co du lieu/FB cua ai)
RAW   = "https://raw.githubusercontent.com/ducanh051101-boop/dhb-nhap-hang-loat/main/"
FILES = ["import_bai_dang.py", "gui.py", "capnhat.py"]   # cac file tu cap nhat


def _ver(text):
    m = re.search(r'VERSION\s*=\s*"([^"]+)"', text or "")
    return m.group(1) if m else "0"

def _local_version():
    try:
        with open(os.path.join(HERE, "import_bai_dang.py"), encoding="utf-8") as f:
            return _ver(f.read())
    except Exception:
        return "0"

def _newer(remote, local):
    def parts(v):
        return [int(x) if x.isdigit() else 0 for x in re.split(r"[.\-_]", str(v))]
    try:
        return parts(remote) > parts(local)
    except Exception:
        return str(remote) > str(local)

def _fetch(name, timeout=8):
    req = urllib.request.Request(RAW + name,
                                 headers={"Cache-Control": "no-cache", "User-Agent": "dhb-updater"})
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8")


def check_and_update():
    """Tra ve True neu VUA cap nhat (nen chay lai code moi). Moi loi -> False."""
    try:
        remote_core = _fetch("import_bai_dang.py")
    except Exception:
        return False                                  # khong co mang -> bo qua
    if not _newer(_ver(remote_core), _local_version()):
        return False                                  # da moi nhat
    # Tai het cac file + KIEM TRA COMPILE truoc khi ghi de (an toan)
    newfiles = {"import_bai_dang.py": remote_core}
    try:
        for f in FILES:
            if f not in newfiles:
                newfiles[f] = _fetch(f)
        for f, data in newfiles.items():
            compile(data, f, "exec")                  # file tai ve phai chay duoc
    except Exception:
        return False                                  # tai/compile loi -> KHONG ghi de
    ok = False
    for f, data in newfiles.items():
        try:
            with open(os.path.join(HERE, f), "w", encoding="utf-8", newline="\n") as fp:
                fp.write(data)
            ok = True
        except Exception:
            pass
    return ok


if __name__ == "__main__":
    try:
        if check_and_update():
            print("[capnhat] Da cap nhat phien ban moi.")
        else:
            print("[capnhat] Dang dung ban moi nhat (hoac khong co mang).")
    except Exception as e:
        print("[capnhat] Bo qua:", e)
