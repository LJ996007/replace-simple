"""在独立 Windows 桌面上运行 pytest，测试窗口不弹到用户桌面、不抢焦点。

用法：python tests/hidden_pytest.py <pytest 参数…>
仅 Windows 有效；创建独立桌面失败时自动退回普通执行。
"""
import os
import subprocess
import sys
import tempfile

if os.name != 'nt':  # 其他平台退回普通执行
    raise SystemExit(subprocess.call([sys.executable, '-m', 'pytest', *sys.argv[1:]]))

import ctypes
import msvcrt
from ctypes import wintypes

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
user32 = ctypes.WinDLL('user32', use_last_error=True)

GENERIC_ALL = 0x10000000
STARTF_USESTDHANDLES = 0x100
HANDLE_FLAG_INHERIT = 0x1
INFINITE = 0xFFFFFFFF


class STARTUPINFO(ctypes.Structure):
    _fields_ = [('cb', wintypes.DWORD), ('lpReserved', wintypes.LPWSTR),
                ('lpDesktop', wintypes.LPWSTR), ('lpTitle', wintypes.LPWSTR),
                ('dwX', wintypes.DWORD), ('dwY', wintypes.DWORD),
                ('dwXSize', wintypes.DWORD), ('dwYSize', wintypes.DWORD),
                ('dwXCountChars', wintypes.DWORD), ('dwYCountChars', wintypes.DWORD),
                ('dwFillAttribute', wintypes.DWORD), ('dwFlags', wintypes.DWORD),
                ('wShowWindow', wintypes.WORD), ('cbReserved2', wintypes.WORD),
                ('lpReserved2', ctypes.c_void_p),
                ('hStdInput', wintypes.HANDLE), ('hStdOutput', wintypes.HANDLE),
                ('hStdError', wintypes.HANDLE)]


class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [('hProcess', wintypes.HANDLE), ('hThread', wintypes.HANDLE),
                ('dwProcessId', wintypes.DWORD), ('dwThreadId', wintypes.DWORD)]


def _inheritable(handle):
    kernel32.SetHandleInformation(handle, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT)


def run_hidden(arguments):
    desktop = user32.CreateDesktopW('ZCodePytest', None, None, 0, GENERIC_ALL, None)
    if not desktop:
        print('无法创建独立桌面，退回普通执行。', file=sys.stderr)
        return subprocess.call([sys.executable, '-m', 'pytest', *arguments])
    with tempfile.TemporaryFile(mode='w+', encoding='utf-8', errors='replace') as output:
        out_handle = msvcrt.get_osfhandle(output.fileno())
        _inheritable(out_handle)
        startup = STARTUPINFO()
        startup.cb = ctypes.sizeof(STARTUPINFO)
        startup.lpDesktop = 'winsta0\\ZCodePytest'
        startup.dwFlags = STARTF_USESTDHANDLES
        startup.hStdInput = msvcrt.get_osfhandle(sys.stdin.fileno())
        startup.hStdOutput = out_handle
        startup.hStdError = out_handle
        command = ' '.join([f'"{sys.executable}"', '-m', 'pytest',
                            *(f'"{a}"' for a in arguments)])
        info = PROCESS_INFORMATION()
        if not kernel32.CreateProcessW(None, command, None, None, True, 0, None, None,
                                       ctypes.byref(startup), ctypes.byref(info)):
            print('独立桌面启动失败：%d，退回普通执行。' % ctypes.get_last_error(), file=sys.stderr)
            user32.CloseDesktop(desktop)
            return subprocess.call([sys.executable, '-m', 'pytest', *arguments])
        kernel32.WaitForSingleObject(info.hProcess, INFINITE)
        code = wintypes.DWORD()
        kernel32.GetExitCodeProcess(info.hProcess, ctypes.byref(code))
        kernel32.CloseHandle(info.hThread)
        kernel32.CloseHandle(info.hProcess)
        user32.CloseDesktop(desktop)
        output.seek(0)
        sys.stdout.write(output.read())
        sys.stdout.flush()
        return code.value


if __name__ == '__main__':
    raise SystemExit(run_hidden(sys.argv[1:]))
