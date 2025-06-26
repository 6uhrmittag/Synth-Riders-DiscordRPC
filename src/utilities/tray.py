import os
import sys
import webbrowser
import subprocess
from threading import Thread
from pystray import Icon, Menu, MenuItem
from PIL import Image
from config import Config


class SystemTray:
    def __init__(self, presence_instance=None):
        self.presence = presence_instance
        self.running = True
        self.update_cycle = 15  # Default update cycle in seconds
        
        # Load icon
        self.icon_image = self._load_icon()
        
        # Create menu
        menu = Menu(
            MenuItem(f"Synth Riders DiscordRPC v{Config.VERSION}", enabled=False),
            MenuItem("──────────────", enabled=False),
            MenuItem(f"Update Cycle: {self.update_cycle}s", self._show_update_cycle_menu),
            MenuItem("Open Settings", self._open_settings),
            MenuItem("──────────────", enabled=False),
            MenuItem("Exit", self._exit_app)
        )
        
        self.icon = Icon(
            name="SynthRidersRPC",
            title="Synth Riders DiscordRPC",
            icon=self.icon_image,
            menu=menu
        )

    def _load_icon(self):
        """Load the application icon"""
        try:
            # Get the script directory
            if getattr(sys, 'frozen', False):
                # Running as executable
                base_path = sys._MEIPASS
            else:
                # Running as script
                base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            
            icon_path = os.path.join(base_path, "assets", "logo.ico")
            
            if os.path.exists(icon_path):
                return Image.open(icon_path)
            else:
                # Fallback to a simple colored square if icon not found
                return Image.new('RGB', (64, 64), color='blue')
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
            # Create a simple fallback icon
            return Image.new('RGB', (64, 64), color='blue')

    def _show_update_cycle_menu(self, icon, item):
        """Show submenu for update cycle options"""
        # Create a simple menu with common update cycle options
        submenu = Menu(
            MenuItem("5 seconds", lambda: self._set_update_cycle(5)),
            MenuItem("10 seconds", lambda: self._set_update_cycle(10)),
            MenuItem("15 seconds (default)", lambda: self._set_update_cycle(15)),
            MenuItem("30 seconds", lambda: self._set_update_cycle(30)),
            MenuItem("60 seconds", lambda: self._set_update_cycle(60))
        )
        
        # Update the main menu with the new cycle info
        self._update_menu()

    def _set_update_cycle(self, seconds):
        """Set the update cycle for the RPC"""
        self.update_cycle = seconds
        if self.presence:
            # Update the presence instance's update cycle
            self.presence.update_cycle = seconds
        self._update_menu()

    def _update_menu(self):
        """Update the system tray menu"""
        menu = Menu(
            MenuItem(f"Synth Riders DiscordRPC v{Config.VERSION}", enabled=False),
            MenuItem("──────────────", enabled=False),
            MenuItem(f"Update Cycle: {self.update_cycle}s", self._show_update_cycle_options),
            MenuItem("Open Settings", self._open_settings),
            MenuItem("──────────────", enabled=False),
            MenuItem("Exit", self._exit_app)
        )
        self.icon.menu = menu

    def _show_update_cycle_options(self, icon, item):
        """Show update cycle options in a simpler way"""
        # For now, just cycle through common values
        cycles = [5, 10, 15, 30, 60]
        current_index = cycles.index(self.update_cycle) if self.update_cycle in cycles else 2
        next_index = (current_index + 1) % len(cycles)
        self._set_update_cycle(cycles[next_index])

    def _open_settings(self, icon, item):
        """Open the settings file or directory"""
        try:
            # Try to find the config file location
            if getattr(sys, 'frozen', False):
                # Running as executable - config should be next to executable
                config_dir = os.path.join(os.path.dirname(sys.executable), "config")
            else:
                # Development mode - check common locations
                config_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Synth Riders DiscordRPC', 'config')
                if not os.path.exists(config_dir):
                    config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "config")

            config_file = os.path.join(config_dir, "config.json")
            
            if os.path.exists(config_file):
                # Open the config file with default editor
                if sys.platform == "win32":
                    os.startfile(config_file)
                elif sys.platform == "darwin":
                    subprocess.run(["open", config_file])
                else:
                    subprocess.run(["xdg-open", config_file])
            elif os.path.exists(config_dir):
                # Open the config directory
                if sys.platform == "win32":
                    os.startfile(config_dir)
                elif sys.platform == "darwin":
                    subprocess.run(["open", config_dir])
                else:
                    subprocess.run(["xdg-open", config_dir])
            else:
                # Fallback: open project GitHub page
                webbrowser.open("https://github.com/6uhrmittag/Synth-Riders-DiscordRPC")
                
        except Exception as e:
            print(f"Error opening settings: {e}")
            # Fallback: open project GitHub page
            webbrowser.open("https://github.com/6uhrmittag/Synth-Riders-DiscordRPC")

    def _exit_app(self, icon, item):
        """Exit the application"""
        self.running = False
        if self.presence:
            try:
                # Clean shutdown of presence
                self.presence.stop()
            except:
                pass
        icon.stop()
        sys.exit(0)

    def run(self):
        """Run the system tray"""
        self.icon.run()

    def stop(self):
        """Stop the system tray"""
        self.running = False
        self.icon.stop()