"""
Helper for running LibreOffice (soffice) in different environments.
On Linux sandboxed VMs: applies AF_UNIX socket shim via LD_PRELOAD.
On Windows/macOS: runs normally without shim.
"""

import os
import platform
import socket
import subprocess
import tempfile
from pathlib import Path


def get_soffice_env() -> dict:
    env = os.environ.copy()

    if platform.system() == "Windows":
        # Adicionar diretório do LibreOffice ao PATH para subprocess achar soffice.exe
        lo_dirs = [
            r"C:\Program Files\LibreOffice\program",
            r"C:\Program Files (x86)\LibreOffice\program",
        ]
        for lo_dir in lo_dirs:
            if os.path.isdir(lo_dir) and lo_dir not in env.get("PATH", ""):
                env["PATH"] = lo_dir + os.pathsep + env.get("PATH", "")
                break
        return env

    # Linux / macOS: modo headless sem display
    env["SAL_USE_VCLPLUGIN"] = "svp"

    if _needs_shim():
        shim = _ensure_shim()
        if shim:
            env["LD_PRELOAD"] = str(shim)

    return env


_SHIM_SO = Path(tempfile.gettempdir()) / "lo_socket_shim.so"


def _needs_shim() -> bool:
    if platform.system() != "Linux":
        return False
    try:
        s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        s.close()
        return False
    except OSError:
        return True


def _ensure_shim():
    if _SHIM_SO.exists():
        return _SHIM_SO

    try:
        src = Path(tempfile.gettempdir()) / "lo_socket_shim.c"
        src.write_text(_SHIM_SOURCE, encoding="utf-8")
        result = subprocess.run(
            ["gcc", "-shared", "-fPIC", "-o", str(_SHIM_SO), str(src), "-ldl"],
            capture_output=True,
            timeout=30,
        )
        if result.returncode == 0:
            return _SHIM_SO
    except Exception:
        pass
    return None


_SHIM_SOURCE = r"""
#define _GNU_SOURCE
#include <dlfcn.h>
#include <sys/socket.h>
#include <errno.h>

typedef int (*socket_fn)(int, int, int);

int socket(int domain, int type, int protocol) {
    socket_fn real_socket = (socket_fn)dlsym(RTLD_NEXT, "socket");
    if (domain == AF_UNIX) {
        int sv[2];
        if (socketpair(AF_UNIX, type, 0, sv) == 0) {
            return sv[0];
        }
        errno = EAFNOSUPPORT;
        return -1;
    }
    return real_socket(domain, type, protocol);
}
"""
