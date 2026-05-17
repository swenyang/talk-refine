"""Windows platform implementation."""

import os
import sys
import time
import logging
import threading

logger = logging.getLogger("talkrefine")


_hotkey_manager = None


def register_hotkey(key: str, callback):
    """Register a global hotkey (uses Win32 API, survives lock/unlock)."""
    global _hotkey_manager
    if _hotkey_manager is None:
        from talkrefine.platform.hotkeys import HotkeyManager
        _hotkey_manager = HotkeyManager()
    _hotkey_manager.register(key, callback)


def start_hotkey_listener():
    """Start the hotkey message loop. Call after all register_hotkey() calls."""
    if _hotkey_manager:
        _hotkey_manager.start()


def wait_forever():
    """Block the current thread until interrupted."""
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass


def _open_clipboard_with_retry(max_retries=5, delay=0.05):
    """Open the clipboard with retries (another process may hold it briefly)."""
    import win32clipboard
    for i in range(max_retries):
        try:
            win32clipboard.OpenClipboard()
            return
        except Exception:
            if i == max_retries - 1:
                raise
            time.sleep(delay)


def _save_clipboard():
    """Save clipboard contents using Win32 API. Returns list of (format, data) or None on failure."""
    try:
        import win32clipboard
    except ImportError:
        return None

    # Formats safe to round-trip (data is bytes/str, not OS handles)
    SAFE_FORMATS = {
        win32clipboard.CF_UNICODETEXT,  # Unicode text
        win32clipboard.CF_TEXT,         # ANSI text
        win32clipboard.CF_OEMTEXT,      # OEM text
        win32clipboard.CF_DIB,          # Device-independent bitmap (screenshots)
        win32clipboard.CF_HDROP,        # File list
    }
    # CF_DIBV5 may not be defined in all pywin32 versions
    CF_DIBV5 = 17
    SAFE_FORMATS.add(CF_DIBV5)

    saved = []
    try:
        _open_clipboard_with_retry()
    except Exception:
        logger.debug("Failed to open clipboard for saving")
        return None
    try:
        fmt = 0
        while True:
            fmt = win32clipboard.EnumClipboardFormats(fmt)
            if fmt == 0:
                break
            if fmt not in SAFE_FORMATS:
                # For unknown/private formats, try if data is bytes
                try:
                    data = win32clipboard.GetClipboardData(fmt)
                    if isinstance(data, (bytes, str)):
                        saved.append((fmt, data))
                    else:
                        logger.debug("Skipping clipboard format %d: non-serializable type %s", fmt, type(data).__name__)
                except Exception as e:
                    logger.debug("Skipping clipboard format %d: %s", fmt, e)
                continue
            try:
                data = win32clipboard.GetClipboardData(fmt)
                saved.append((fmt, data))
            except Exception as e:
                logger.debug("Skipping clipboard format %d: %s", fmt, e)
    finally:
        win32clipboard.CloseClipboard()
    return saved


def _restore_clipboard(saved):
    """Restore clipboard contents using Win32 API."""
    import win32clipboard

    try:
        _open_clipboard_with_retry()
    except Exception:
        logger.debug("Failed to open clipboard for restoring")
        return
    try:
        win32clipboard.EmptyClipboard()
        for fmt, data in saved:
            try:
                win32clipboard.SetClipboardData(fmt, data)
            except Exception as e:
                logger.debug("Failed to restore clipboard format %d: %s", fmt, e)
    finally:
        win32clipboard.CloseClipboard()


def paste_text(text: str, preserve_clipboard: bool = True):
    """Paste text at current cursor position."""
    import pyautogui

    saved_clipboard = None
    if preserve_clipboard:
        saved_clipboard = _save_clipboard()
        if saved_clipboard is None:
            # win32clipboard unavailable, fall back to pyperclip text-only
            try:
                import pyperclip
                saved_clipboard = pyperclip.paste()
            except Exception:
                saved_clipboard = None

    # Set text to clipboard
    try:
        import win32clipboard
        _open_clipboard_with_retry()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
    except ImportError:
        import pyperclip
        pyperclip.copy(text)

    time.sleep(0.15)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.15)

    if preserve_clipboard and saved_clipboard is not None:
        try:
            if isinstance(saved_clipboard, list):
                _restore_clipboard(saved_clipboard)
            elif isinstance(saved_clipboard, str):
                # pyperclip fallback: text-only restore
                import pyperclip
                pyperclip.copy(saved_clipboard)
        except Exception:
            pass


def copy_text(text: str):
    """Copy text to clipboard."""
    import pyperclip
    pyperclip.copy(text)


def _create_shortcut_ps(shortcut_path, target, arguments, working_dir, description,
                        icon_path=None):
    """Create .lnk shortcut using PowerShell (fallback when pywin32 unavailable)."""
    import subprocess
    ps_lines = [
        '$ws = New-Object -ComObject WScript.Shell',
        f'$sc = $ws.CreateShortcut("{shortcut_path}")',
        f'$sc.TargetPath = "{target}"',
        f"$sc.Arguments = '{arguments}'",
        f'$sc.WorkingDirectory = "{working_dir}"',
        f'$sc.Description = "{description}"',
        '$sc.WindowStyle = 7',
    ]
    if icon_path and os.path.exists(icon_path):
        ps_lines.append(f'$sc.IconLocation = "{icon_path}"')
    ps_lines.append('$sc.Save()')
    ps_script = '; '.join(ps_lines)
    subprocess.run(["powershell", "-Command", ps_script],
                   capture_output=True, timeout=10)


def setup_autostart(enable: bool = True):
    """Add/remove from Windows startup folder."""
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    vbs_path = os.path.join(project_root, "scripts", "start_hidden.vbs")
    ico_path = os.path.join(project_root, "assets", "talkrefine.ico")

    startup_folder = os.path.join(
        os.environ.get("APPDATA", ""),
        r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    shortcut_path = os.path.join(startup_folder, "TalkRefine.lnk")

    # Clean up old registry entry if present
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_SET_VALUE)
        try:
            winreg.DeleteValue(key, "TalkRefine")
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
    except Exception:
        pass

    if enable:
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            sc = shell.CreateShortCut(shortcut_path)
            sc.TargetPath = "wscript.exe"
            sc.Arguments = f'"{vbs_path}"'
            sc.WorkingDirectory = project_root
            sc.Description = "TalkRefine"
            sc.WindowStyle = 7
            if os.path.exists(ico_path):
                sc.IconLocation = ico_path
            sc.Save()
        except ImportError:
            # Fallback: use PowerShell to create shortcut
            _create_shortcut_ps(shortcut_path, "wscript.exe",
                                f'"{vbs_path}"', project_root,
                                "TalkRefine", ico_path)
        except Exception as e:
            logger.warning("Startup shortcut failed: %s", e)
    else:
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)


def create_start_menu_shortcut():
    """Create a Start Menu shortcut."""
    start_menu = os.path.join(
        os.environ.get("APPDATA", ""),
        r"Microsoft\Windows\Start Menu\Programs"
    )
    shortcut_path = os.path.join(start_menu, "TalkRefine.lnk")
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    vbs_path = os.path.join(project_root, "scripts", "start_hidden.vbs")
    ico_path = os.path.join(project_root, "assets", "talkrefine.ico")

    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(shortcut_path)
        sc.TargetPath = "wscript.exe"
        sc.Arguments = f'"{vbs_path}"'
        sc.WorkingDirectory = project_root
        sc.Description = "TalkRefine - Voice to refined text"
        sc.WindowStyle = 7
        if os.path.exists(ico_path):
            sc.IconLocation = ico_path
        sc.Save()
        return shortcut_path
    except ImportError:
        _create_shortcut_ps(shortcut_path, "wscript.exe",
                            f'"{vbs_path}"', project_root,
                            "TalkRefine - Voice to refined text", ico_path)
        return shortcut_path
    except Exception as e:
        logger.warning("Failed to create shortcut: %s", e)
        return None


def remove_start_menu_shortcut():
    """Remove Start Menu shortcut."""
    shortcut_path = os.path.join(
        os.environ.get("APPDATA", ""),
        r"Microsoft\Windows\Start Menu\Programs",
        "TalkRefine.lnk"
    )
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
