import os
import sys
import json
import subprocess


def create_vbs_launcher(vbs_path, exe_path, working_dir):
    """
    Create hidden VBS launcher.
    """
    with open(vbs_path, "w") as f:
        f.write(
            f'Dim sh\n'
            f'Set sh = CreateObject("WScript.Shell")\n'
            f'sh.CurrentDirectory = "{working_dir}"\n'
            f'sh.Run Chr(34) & "{exe_path}" & Chr(34), 0, False\n'
        )

# windows_user, windows_pass,
def apply_advanced_task_settings(task_name, log=None):
    """
    Apply advanced Task Scheduler settings using PowerShell.
    No credentials needed — task runs as SYSTEM.
    
    # schtasks.exe cannot configure ALL advanced settings directly.
    # stop existing instance # battery conditions # missed task handling
    # So we use PowerShell immediately AFTER creation.
    # THIS FIXES:
    # Power issues

    # Disables: # “Start only on AC” # “Stop on battery”
    # Missed schedule recovery # Enables: # Run task as soon as possible after missed start
    # Hung process recovery # Sets: # Stop existing instance # This is VERY important.
    """

    settings_cmd = (
        f'powershell -Command "'
        f'$settings = New-ScheduledTaskSettingsSet '
        f'-AllowStartIfOnBatteries '
        f'-DontStopIfGoingOnBatteries '
        f'-StartWhenAvailable '
        f'-WakeToRun '
        f'-MultipleInstances IgnoreNew '
        f'-ExecutionTimeLimit (New-TimeSpan -Days 0); '
        f'Set-ScheduledTask '
        f'-TaskName \'{task_name}\' '
        f'-Settings $settings'
        f'"'
        )

    result = subprocess.run(
        settings_cmd,
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        if log:
            log.warning(f"[TASKS] Failed advanced settings for {task_name}")
            log.warning(result.stderr)
        else:
            print(result.stderr)

        return False

    return True

# windows_user,
# windows_pass,
def create_or_update_task(
    task_name,
    task_cmd,
    schedule,
    log=None
):
    """
    Delete, recreate, and apply advanced settings to a scheduled task.

    # Args:
    #     task_name    : Display name e.g. "KRA Auto Checker"
    #     task_cmd     : Command to run e.g. 'wscript.exe "C:\\...\\run.vbs"'
    #     schedule     : schtasks schedule string e.g. "/sc daily /st 19:00"
    #     windows_user : Windows account e.g. "DESKTOP-ABC\\john"
    #     windows_pass : Windows account password
    #     log          : Optional logger instance

    No windows_user or windows_pass needed — runs as SYSTEM.
    SYSTEM + highest privilege = runs locked, logged out, on battery.
    PowerShell can modify these tasks without credentials later.
    """

    subprocess.run(
        f'schtasks /delete /tn "{task_name}" /f',
        shell=True,
        capture_output=True
    )

    create_cmd = (
    f'schtasks /create '
    f'/tn "{task_name}" '
    f'/tr "{task_cmd}" '
    f'{schedule} '
    f'/f '
    f'/ru SYSTEM'
    )

    result = subprocess.run(
        create_cmd,
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:

        if log:
            log.warning(f"[TASKS] Failed creating {task_name}")
            log.warning(result.stderr)
        else:
            print(result.stderr)

        return False

    advanced_ok = apply_advanced_task_settings(
        task_name,
        log
    )
    # windows_user,
    # windows_pass,

    if advanced_ok:

        if log:
            log.info(f"[TASKS] Created {task_name}")
        else:
            print(f"  Created: {task_name}")

    return advanced_ok


def change_heartbeat_interval(interval_minutes, log=None):
    """
    Change heartbeat interval WITHOUT recreating task.
    Preserves:
    - credentials
    - battery settings
    - advanced settings
    - hidden settings
    """


    ps_cmd = (
        'powershell -Command "'
        f'$task = Get-ScheduledTask -TaskName \'Station Heartbeat\'; '
        f'$task.Triggers[0].Repetition.Interval = \'PT{interval_minutes}M\'; '
        f'Set-ScheduledTask -InputObject $task"'
    )

    result = subprocess.run(
        ps_cmd,
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:

        if log:
            log.warning("[TASKS] Failed changing heartbeat interval")
            log.warning(result.stderr)

        return False

    if log:
        log.info(
            f"[TASKS] Station Heartbeat interval updated to "
            f"{interval_minutes} minute(s)"
        )

    return True



def change_kra_schedule_time(kra_check_time, log=None):
    """
    Change KRA checker schedule time WITHOUT recreating task.
    """

    ps_cmd = (
        'powershell -Command "'
        f'$task = Get-ScheduledTask -TaskName \'KRA Auto Checker\'; '
        f'$task.Triggers[0].StartBoundary = \'2026-01-01T{kra_check_time}:00\'; '
        f'Set-ScheduledTask -InputObject $task"'
    )

    result = subprocess.run(
        ps_cmd,
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0 or result.stderr.strip():

        if log:
            log.warning("[TASKS] Failed changing KRA schedule")
            log.warning(result.stderr)

        return False

    if log:
        log.info(
            f"[TASKS] KRA Auto Checker schedule updated to "
            f"{kra_check_time}"
        )

    return True
