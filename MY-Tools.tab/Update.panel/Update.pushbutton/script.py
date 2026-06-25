# -*- coding: utf-8 -*-

import os
import shutil
import subprocess
import tempfile

from pyrevit import forms
from pyrevit.loader.sessionmgr import load_session


# --------------------------------------------------
# SETTINGS
# --------------------------------------------------

BAT_NAME = "update_tools.bat"


# --------------------------------------------------
# FIND BAT FILE
# --------------------------------------------------

button_folder = os.path.dirname(__file__)
source_bat = os.path.join(button_folder, BAT_NAME)

if not os.path.exists(source_bat):
    forms.alert(
        "Cannot find update_tools.bat.\n\nExpected location:\n{}".format(source_bat),
        title="Update Failed"
    )
    raise SystemExit


# --------------------------------------------------
# COPY BAT TO TEMP BEFORE RUNNING
# --------------------------------------------------
# Important:
# The updater may replace this toolbar folder.
# So we copy the bat to temp first, then run the temp copy.

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


# --------------------------------------------------
# RUN BAT
# --------------------------------------------------

try:
    process = subprocess.Popen(
        ["cmd.exe", "/c", temp_bat],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False
    )

    stdout, stderr = process.communicate()
    exit_code = process.returncode

except Exception as ex:
    forms.alert(
        "Failed to run update_tools.bat.\n\n{}".format(ex),
        title="Update Failed"
    )
    raise SystemExit


# --------------------------------------------------
# HANDLE RESULT
# --------------------------------------------------

if exit_code == 0:
    forms.alert(
        "Toolbar updated successfully.\n\npyRevit will reload now.",
        title="Update Complete"
    )

    try:
        load_session()
    except Exception as ex:
        forms.alert(
            "Toolbar was updated, but pyRevit reload failed.\n\n"
            "Please click Reload in pyRevit manually.\n\n{}".format(ex),
            title="Reload Failed"
        )

else:
    log_path = os.path.join(os.environ.get("TEMP", ""), "pyRevit_Install_Log.txt")

    message = "Toolbar update failed.\n\n"
    message += "Error code: {}\n\n".format(exit_code)

    if log_path and os.path.exists(log_path):
        message += "Please check the log file:\n{}\n\n".format(log_path)

    if stderr:
        try:
            message += "Error message:\n{}".format(stderr.decode("utf-8", "ignore"))
        except:
            message += "Error message:\n{}".format(stderr)

    forms.alert(message, title="Update Failed")