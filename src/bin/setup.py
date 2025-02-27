import sys
import subprocess
from win32com.client import Dispatch
import winreg
from os import getenv, path, makedirs
from shutil import copyfile
from json import dumps
from psutil import process_iter, NoSuchProcess, Process, AccessDenied, ZombieProcess
from time import sleep, time
from rich.console import Console
from config import Config
from src.utilities.rpc.assets import DiscordAssets
from src.utilities.cli import (
    indent,
    print_divider,
    fatal_error,
    get_input,
    get_rich_presence_install_location,
    get_shortcut_preference,
    get_synthriders_install_location,
    get_startup_preference,
    get_promote_preference,
    get_keep_running_preference,
)

console = Console()

DEFAULT_SYNTHRIDERS_INSTALL_LOCATION = r"C:\Program Files (x86)\Steam\steamapps\common\SynthRiders"
DEFAULT_RICH_PRESENCE_INSTALL_LOCATION = (
    rf"{getenv('LOCALAPPDATA')}\Synth Riders DiscordRPC"
)
LARGE_DIVIDER = r"""
    .-.-.   .-.-.   .-.-.   .-.-.   .-.-.   .-.-.   .-.-.   .-.   
    / / \ \ / / \ \ / / \ \ / / \ \ / / \ \ / / \ \ / / \ \ / / 
    `-'   `-`-'   `-`-'   `-`-'   `-`-'   `-`-'   `-`-'   `-`-'      
"""
ASCII_ART = r"""
  _________             __  .__      __________.__    .___
 /   _____/__.__. _____/  |_|  |__   \______   \__| __| _/___________  ______
 \_____  <   |  |/    \   __\  |  \   |       _/  |/ __ |/ __ \_  __ \/  ___/
 /        \___  |   |  \  | |   Y  \  |    |   \  / /_/ \  ___/|  | \/\___ \
/_______  / ____|___|  /__| |___|  /  |____|_  /__\____ |\___  >__|  /____  >
        \/\/         \/          \/          \/        \/    \/          \/
________  __                              .________________________________
\______ \ |__| ______ ____  ___________  __| _/\______   \______   \_   ___ \
 |    |  \|  |/  ___// ___\/  _ \_  __ \/ __ |  |       _/|     ___/    \  \/
 |    `   \  |\___ \\  \__(  <_> )  | \/ /_/ |  |    |   \|    |   \     \____
/_______  /__/____  >\___  >____/|__|  \____ |  |____|_  /|____|    \______  /
        \/        \/     \/                 \/         \/                  \/
"""


def print_welcome_message(console: Console) -> None:
    """
    Print a welcome message to the console

    :param console: The console to use for output
    """
    console.print(
        "\n\n",
        ASCII_ART,
        indent(
            "\n\n",
            "[blue]Welcome![/blue] Please follow the instructions below to set up the program.",
            "Source code for this program can be found at https://github.com/6uhrmittag/Synth-Riders-DiscordRPC",
            "This tool is based on: https://github.com/xAkre/Wuthering-Waves-RPC",
            "[red]Please note that this program is not affiliated with Synth Riders or its developers.[/red]",
        ),
        highlight=False,
    )


def get_config(console: Console) -> dict:
    """
    Get the configuration options from the user

    :param console: The console to use for input and output
    """
    synthriders_install_location = get_input(
        console,
        "WSynth Riders Install Location",
        lambda: get_synthriders_install_location(console, DEFAULT_SYNTHRIDERS_INSTALL_LOCATION),
    )

    config = {
        "version": Config.VERSION,
        "synthriders_install_location": synthriders_install_location,
        "rich_presence_install_location": get_input(
            console,
            "Rich Presence Install Location",
            lambda: get_rich_presence_install_location(
                console, DEFAULT_RICH_PRESENCE_INSTALL_LOCATION
            ),
        ),
        "startup_preference": get_input(
            console,
            "Launch on Startup Preference",
            lambda: get_startup_preference(console),
        ),
        "keep_running_preference": get_input(
            console,
            "Keep Running Preference",
            lambda: get_keep_running_preference(console),
        ),
        "shortcut_preference": get_input(
            console,
            "Create Shortcut Preference",
            lambda: get_shortcut_preference(console),
        ),
        "promote_preference": get_input(
            console,
            "Promote Preference",
            lambda: get_promote_preference(console),
        ),
        "discord_application_id": Config.APPLICATION_ID,
        "discord_application_logo_large": DiscordAssets.LARGE_IMAGE,
        "discord_application_logo_small": DiscordAssets.SMALL_IMAGE,
        "synthriders_websocket_host": Config.WEBSOCKET_HOST,
        "synthriders_websocket_port": Config.WEBSOCKET_PORT,
        "image_upload_url": Config.IMAGE_UPLOAD_URL,
    }


    return config


def stop_running_process(console, process_name, timeout=5):
    """
    Stops all running instances of a process and its child processes before uninstalling.

    :param console: The console to use for output.
    :param process_name: The name of the executable to terminate.
    :param timeout: Time (in seconds) to wait before force-killing the process.
    """
    found = False

    def terminate_process_tree(process):
        """ Recursively terminate a process and its child processes. """
        try:
            children = process.children(recursive=True)  # Get all child processes
            for child in children:
                console.print(
                    indent(f"Stopping child process {child.name()} (PID: {child.pid})..."),
                    style="yellow")
                child.terminate()

            process.terminate()  # Gracefully terminate the parent

            # Wait for all processes to exit
            for _ in range(timeout * 2):  # Check every 0.5s for `timeout` seconds
                if not process.is_running():
                    console.print(
                        indent(f"{process.name()} (PID: {process.pid}) stopped."),
                        style="green"
                    )
                    return
                sleep(0.5)

            # Force kill if still running
            console.print(
                indent(f"{process.name()} (PID: {process.pid}) did not terminate in time, forcing shutdown..."),
                style="red"
            )
            process.kill()

            for child in children:  # Ensure child processes are also killed
                if child.is_running():
                    child.kill()

            console.print(
                indent(f"{process.name()} forcefully stopped."),
                style="green"
                )

        except (NoSuchProcess, AccessDenied, ZombieProcess):
            pass

    for process in process_iter(attrs=["pid", "name"]):
        try:
            if process.info["name"].lower() == process_name.lower():
                found = True
                console.print(
                    indent(f"{process_name} is running! Stopping {process_name} (PID: {process.info['pid']})..."),
                    style="yellow",
                )
                terminate_process_tree(Process(process.info["pid"]))

        except (NoSuchProcess, AccessDenied, ZombieProcess):
            continue



def create_config_folder(console: Console, config: dict) -> None:
    """
    Create the config folder in the install location

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(
            indent("Creating the config folder in the install location..."),
            spinner="dots",
        ):
            makedirs(path.join(config["rich_presence_install_location"], "config"))
            console.print(indent("Config folder created."), style="green")
    except Exception as e:
        fatal_error(
            console, indent(f"An error occurred while creating the config folder"), e
        )


def write_config_to_file(console: Console, config: dict) -> None:
    """
    Write the configuration to a file

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with open(
            path.join(
                config["rich_presence_install_location"], "config", "config.json"
            ),
            "w",
        ) as f:
            f.write(dumps(config, indent=4))

        console.print(indent("Configuration written to file."), style="green")
    except Exception as e:
        fatal_error(
            console, indent(f"An error occurred while writing the config to a file"), e
        )


def copy_main_exe_to_install_location(console: Console, config: dict) -> None:
    """
    Copy the main executable to the install location

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(
            indent("Copying the main executable to the install location..."),
            spinner="dots",
        ):
            copyfile(
                path.join(sys._MEIPASS, Config.MAIN_EXECUTABLE_NAME),
                path.join(
                    config["rich_presence_install_location"],
                    Config.MAIN_EXECUTABLE_NAME,
                ),
            )
            console.print(
                indent("Main executable copied to install location."), style="green"
            )
    except Exception as e:
        fatal_error(
            console,
            indent(
                f"An error occurred while copying the main executable to the install location",
            ),
            e,
        )


def copy_uninstall_exe_to_install_location(console: Console, config: dict) -> None:
    """
    Copy the uninstall executable to the install location

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(
            indent("Copying the uninstall executable to the install location..."),
            spinner="dots",
        ):
            copyfile(
                path.join(sys._MEIPASS, Config.UNINSTALL_EXECUTABLE_NAME),
                path.join(
                    config["rich_presence_install_location"],
                    Config.UNINSTALL_EXECUTABLE_NAME,
                ),
            )
            console.print(
                indent("Uninstall executable copied to install location."),
                style="green",
            )
    except Exception as e:
        fatal_error(
            console,
            indent(
                f"An error occurred while copying the uninstall executable to the install location",
            ),
            e,
        )


def add_exe_to_windows_apps(console: Console, config: dict) -> None:
    """
    Add the executable to the Windows App list

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(
            indent("Adding the executable to the Windows App list..."), spinner="dots"
        ):
            programs_folder = path.join(
                getenv("APPDATA"),
                "Microsoft/Windows/Start Menu/Programs",
            )

            if not path.exists(programs_folder):
                makedirs(programs_folder)

            main_exe_shortcut_path = path.join(
                programs_folder, Config.MAIN_EXECUTABLE_NAME.replace(".exe", ".lnk")
            )
            main_exe_shortcut_target = path.join(
                config["rich_presence_install_location"], Config.MAIN_EXECUTABLE_NAME
            )

            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(main_exe_shortcut_path)
            shortcut.TargetPath = main_exe_shortcut_target
            shortcut.Save()

            uninstall_exe_shortcut_path = path.join(
                programs_folder,
                Config.UNINSTALL_EXECUTABLE_NAME.replace(".exe", ".lnk"),
            )
            uninstall_exe_shortcut_target = path.join(
                config["rich_presence_install_location"],
                Config.UNINSTALL_EXECUTABLE_NAME,
            )

            shortcut = shell.CreateShortcut(uninstall_exe_shortcut_path)
            shortcut.TargetPath = uninstall_exe_shortcut_target
            shortcut.Save()

            console.print(
                indent("Executable added to Windows App list."), style="green"
            )
    except Exception as e:
        console.print(
            indent(
                "An error occurred while adding the executable to the Windows App list:"
            ),
            style="red",
        )
        console.print_exception()
        console.print(
            indent("The executable will not be added to the Windows App list.")
        )
        console.print(indent("Setup will continue..."))


def launch_exe_on_startup(console: Console, config: dict) -> None:
    """
    Launch the executable on system startup

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(
            indent("Setting the executable to launch on startup..."), spinner="dots"
        ):
            shortcut_target = path.join(
                config["rich_presence_install_location"],
                Config.MAIN_EXECUTABLE_NAME,
            )

            # Thank god ChatGPT for this
            create_task_command = [
                "schtasks",
                "/create",
                "/tn",
                "Synth Riders DiscordRPC",
                "/tr",
                f'"{shortcut_target}"',
                "/sc",
                "onlogon",
                "/rl",
                "highest",
                "/f",
            ]

            subprocess.run(create_task_command, check=True, stdout=subprocess.DEVNULL)
            console.print(indent("Executable set to launch on startup."), style="green")
    except Exception as e:
        console.print(
            indent(
                "An error occurred while setting the executable to launch on startup:"
            ),
            style="red",
        )
        console.print_exception()
        console.print(
            indent("The executable will not launch on startup. Setup"), style="red"
        )
        console.print(indent("Setup will continue..."))


def add_to_startup_registry(console, config):
    """
    Adds the executable to Windows startup via the Registry (without admin rights).

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status("Adding application to startup via registry...", spinner="dots"):
            shortcut_target = path.join(
                config["rich_presence_install_location"],
                Config.MAIN_EXECUTABLE_NAME,
            )

            key = winreg.HKEY_CURRENT_USER
            subkey = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "SynthRidersRPC"

            with winreg.OpenKey(key, subkey, 0, winreg.KEY_SET_VALUE) as reg_key:
                winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, f'"{shortcut_target}"')

            console.print("Executable successfully added to startup (Registry, no admin needed)", style="green")

    except Exception as e:
        console.print("Failed to add application to startup (Registry).", style="red")
        console.print_exception()

def create_windows_shortcut(console: Console, config: dict) -> None:
    """
    Create a desktop shortcut for the executable

    :param console: The console to use for output
    :param config: The configuration options
    """
    try:
        with console.status(indent("Creating a desktop shortcut..."), spinner="dots"):
            shortcut_path = path.join(
                path.expanduser("~/Desktop"),
                Config.MAIN_EXECUTABLE_NAME.replace(".exe", ".lnk"),
            )
            shortcut_target = path.join(
                config["rich_presence_install_location"], Config.MAIN_EXECUTABLE_NAME
            )
            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = shortcut_target
            shortcut.Save()
            console.print(indent("Desktop shortcut created."), style="green")
    except Exception as e:
        console.print(
            indent("An error occurred while creating the desktop shortcut:"),
            style="red",
        )
        console.print_exception()
        console.print(indent("The desktop shortcut will not be created."), style="red")
        console.print(indent("Setup will continue..."))


print_welcome_message(console)
stop_running_process(console, Config.MAIN_EXECUTABLE_NAME, timeout=5)
config = get_config(console)
print_divider(console, "[green]Options Finalised[/green]", "green")
create_config_folder(console, config)
write_config_to_file(console, config)
copy_main_exe_to_install_location(console, config)
copy_uninstall_exe_to_install_location(console, config)
add_exe_to_windows_apps(console, config)
if config["startup_preference"]:
    add_to_startup_registry(console, config)
    # launch_exe_on_startup(console, config)
if config["shortcut_preference"]:
    create_windows_shortcut(console, config)
print_divider(console, "[green]Setup Completed[/green]", "green")
console.show_cursor(False)

# For some reason using console.input() here doesn't work, so I'm using input() instead
input(indent("Press Enter to exit..."))
exit(0)
