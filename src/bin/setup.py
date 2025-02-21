import sys
import subprocess
from win32com.client import Dispatch
from os import getenv, path, makedirs
from shutil import copyfile
from json import dumps
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
    rf"{getenv('LOCALAPPDATA')}\Synth Riders RPC"
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
        \/\/         \/          \/          \/        \/    \/           \/
                       _____________________________
                       \______   \______   \_   ___ \
                        |       _/|     ___/    \  \/
                        |    |   \|    |   \     \____
                        |____|_  /|____|    \______  /
                               \/                  \/
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
            "[blue]Welcome! Please follow the instructions below to set up the program.[/blue]",
            "Source code for this program can be found at https://github.com/6uhrmittag/Synth-Riders-RPC",
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
    }


    return config


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
                "Synth Riders RPC",
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
config = get_config(console)
print_divider(console, "[green]Options Finalised[/green]", "green")
create_config_folder(console, config)
write_config_to_file(console, config)
copy_main_exe_to_install_location(console, config)
copy_uninstall_exe_to_install_location(console, config)
add_exe_to_windows_apps(console, config)
if config["startup_preference"]:
    launch_exe_on_startup(console, config)
if config["shortcut_preference"]:
    create_windows_shortcut(console, config)
print_divider(console, "[green]Setup Completed[/green]", "green")
console.show_cursor(False)

# For some reason using console.input() here doesn't work, so I'm using input() instead
input(indent("Press Enter to exit..."))
exit(0)
