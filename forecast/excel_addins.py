"""
Excel add-in management via xlwings / COM.

Provides functions to disable and re-enable IntexLINK and Bloomberg
add-ins programmatically.  This fixes #NAME? errors by ensuring
add-ins discover formulas fresh when re-enabled.
"""

from __future__ import annotations

import time

from forecast.config import CONFIG


def toggle_com_addins(app, enable: bool) -> None:
    """Enable or disable the IntexLINK and Bloomberg COM add-ins.

    Args:
        app: An xlwings App instance (connected to a running Excel).
        enable: True to enable (Connect), False to disable (Disconnect).
    """
    action = "Enabling" if enable else "Disabling"

    try:
        com_addins = app.api.COMAddIns
    except Exception as e:
        print(f"  WARNING: Could not access COM Add-ins: {e}")
        return

    addin_names = [
        CONFIG.excel.intex_com_addin,
        CONFIG.excel.bloomberg_com_addin,
    ]

    for name in addin_names:
        if not name:
            continue
        try:
            addin = com_addins(name)
            addin.Connect = enable
            print(f"  {action} COM add-in: {name}")
        except Exception as e:
            print(f"  WARNING: Could not toggle '{name}': {e}")


def toggle_xll_addins(app, enable: bool) -> None:
    """Register or unregister XLL add-ins.

    Args:
        app: An xlwings App instance.
        enable: True to register, False to unregister.
    """
    xll_paths = [
        CONFIG.excel.intex_xll_path,
        CONFIG.excel.bloomberg_xll_path,
    ]

    for path in xll_paths:
        if not path:
            continue
        try:
            if enable:
                app.api.RegisterXLL(path)
                print(f"  Registered XLL: {path}")
            else:
                app.api.UnregisterXLL(path)
                print(f"  Unregistered XLL: {path}")
        except Exception as e:
            print(f"  WARNING: Could not toggle XLL '{path}': {e}")


def disable_addins(app) -> None:
    """Disable both COM and XLL add-ins for IntexLINK and Bloomberg."""
    print("  Disabling Excel add-ins...")
    toggle_com_addins(app, enable=False)
    toggle_xll_addins(app, enable=False)
    time.sleep(2)  # Let Excel process the changes.


def enable_addins(app) -> None:
    """Re-enable both COM and XLL add-ins for IntexLINK and Bloomberg."""
    print("  Re-enabling Excel add-ins...")
    toggle_xll_addins(app, enable=True)
    toggle_com_addins(app, enable=True)
    time.sleep(5)  # Give add-ins time to initialize after re-enabling.
