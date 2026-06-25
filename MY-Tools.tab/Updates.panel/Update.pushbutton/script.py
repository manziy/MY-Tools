# -*- coding: utf-8 -*-

import os
import shutil
import subprocess
import tempfile

from pyrevit import forms
from pyrevit.loader import sessionmgr


BAT_NAME = "update_tools.bat"


def read_log(log_path):
    if not os.path.exists(log_path):
        return ""

    try:
        with open(log_path, "r") as f:
            return f.read()
    except Exception:
        return ""


def get_log_tail(content, max_chars=3000):
    if len(content) > max_chars:
        return content[-max_chars:]
    return content


button_folder = os.path.dirname(__file__)
source_bat = os.path.join(button_folder, BAT_NAME)

if not os.path.exists(source_bat):
    forms.alert(
        "Cannot find update_tools.bat.\n\nExpected location:\n{}".format(source_bat),
        title="Update Failed"
    )
    raise SystemExit


# Copy bat to temp first because the update may replace the toolbar folder
temp_folder = tempfile.mkdtemp(prefix="pyrevit_toolbar_update_")
temp_bat = os.path.join(temp_folder, BAT_NAME)

try:
    shutil.copy2(source_bat, temp_bat)
except Exception as ex:
    forms.alert(
        "Failed to copy update_tools.bat to temp folder.\n\n{}".format(ex),
        title="Update Failed"
    )
    raise SystemExit


# Put log in the same temp folder for this update run
log_path = os.path.join(temp_folder, "pyRevit_Install_Log.txt")

try:
    exit_code = subprocess.call(
        ["cmd.exe", "/d", "/c", temp_bat],
        cwd=temp_folder
    )

except Exception as ex:
    forms.alert(
        "Failed to run update_tools.bat.\n\n{}".format(ex),
        title="Update Failed"
    )
    raise SystemExit


log_content = read_log(log_path)
log_tail = get_log_tail(log_content)

# Important:
# Some bat commands may return error code 1 even after successful install.
# If the log clearly says SUCCESS, treat it as success.
install_success = (
    exit_code == 0 or
    "SUCCESS" in log_content or
    "Installation complete" in log_content
)


if install_success:
    forms.alert(
        "Toolbar updated successfully.\n\npyRevit will reload now.",
        title="Update Complete"
    )

    try:
        sessionmgr.load_session()
    except Exception as ex:
        forms.alert(
            "Toolbar was updated, but pyRevit reload failed.\n\n"
            "Please click Reload in pyRevit manually.\n\n{}".format(ex),
            title="Reload Failed"
        )

else:
    message = "Toolbar update failed.\n\n"
    message += "Error code: {}\n\n".format(exit_code)
    message += "Log file:\n{}\n\n".format(log_path)

    if log_tail:
        message += "Last part of log:\n"
        message += log_tail

    forms.alert(message, title="Update Failed")