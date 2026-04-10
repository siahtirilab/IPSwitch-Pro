import ctypes
import ipaddress
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

import pystray
from PIL import Image


APP_TITLE = "IPSwitch-Pro"
APP_VERSION = "1.4.1"
NEW_PROFILE_LABEL = "New Profile"
GITHUB_URL = "https://github.com/siahtirilab/IPSwitch-Pro"


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


APP_ICON_FILE = resource_path("assets/ip-address.png")
APP_DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / APP_TITLE
APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
PROFILE_FILE = APP_DATA_DIR / "profiles.json"
PING_PROFILE_FILE = APP_DATA_DIR / "ping_profiles.json"


class ProfileStore:
    def __init__(self, path: Path):
        self.path = path

    def load(self) -> list[dict]:
        if not self.path.exists():
            return []
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        return data if isinstance(data, list) else []

    def save(self, profiles: list[dict]) -> None:
        self.path.write_text(json.dumps(profiles, indent=2), encoding="utf-8")


class NetworkManager:
    @staticmethod
    def run_command(command: list[str]) -> tuple[bool, str]:
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            return False, str(exc)

        output = ((completed.stdout or "") + "\n" + (completed.stderr or "")).strip()
        return completed.returncode == 0, output

    @staticmethod
    def is_admin() -> bool:
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    def list_adapters(self) -> list[str]:
        success, output = self.run_command(["netsh", "interface", "show", "interface"])
        adapters = self._parse_netsh_adapters(output) if success else []
        if adapters:
            return adapters

        ps_command = [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-NetAdapter | Sort-Object Name | Select-Object -ExpandProperty Name",
        ]
        success, output = self.run_command(ps_command)
        adapters = [line.strip() for line in output.splitlines() if line.strip()]
        return sorted(set(adapters)) if success else []

    @staticmethod
    def _parse_netsh_adapters(output: str) -> list[str]:
        parsed = []
        for line in output.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("Admin") or stripped.startswith("---"):
                continue
            parts = stripped.split()
            if len(parts) >= 4:
                parsed.append(" ".join(parts[3:]))
        return sorted(set(parsed))

    def get_adapter_details(self, adapter_name: str) -> dict:
        if not adapter_name:
            return {}

        success, output = self.run_command(["netsh", "interface", "ipv4", "show", "config", f'name="{adapter_name}"'])
        details = self._parse_netsh_adapter_details(output) if success else {}
        if details:
            return details

        safe_name = adapter_name.replace("'", "''")
        script = (
            f"$cfg = Get-NetIPConfiguration -InterfaceAlias '{safe_name}' -ErrorAction SilentlyContinue; "
            "if ($cfg) { "
            "$ip = if ($cfg.IPv4Address) { $cfg.IPv4Address.IPAddress } else { '' }; "
            "$prefix = if ($cfg.IPv4Address) { $cfg.IPv4Address.PrefixLength } else { '' }; "
            "$mask = if ($prefix -ne '') { "
            "$bits = [uint32]0; "
            "for ($i = 0; $i -lt [int]$prefix; $i++) { $bits = $bits -bor (1 -shl (31 - $i)) }; "
            "$maskBytes = [BitConverter]::GetBytes([uint32]$bits); "
            "[Array]::Reverse($maskBytes); "
            "([System.Net.IPAddress]::new($maskBytes)).ToString() "
            "} else { '' }; "
            "$dhcp = if ($cfg.NetIPv4Interface) { $cfg.NetIPv4Interface.Dhcp } else { '' }; "
            "Write-Output ('IP=' + $ip); "
            "Write-Output ('MASK=' + $mask); "
            "Write-Output ('DHCP=' + $dhcp) "
            "}"
        )
        success, output = self.run_command(["powershell", "-NoProfile", "-Command", script])
        if not success:
            return {}

        details = {}
        for line in output.splitlines():
            if "=" in line:
                key, value = line.split("=", 1)
                details[key.strip().lower()] = value.strip()
        return details

    @staticmethod
    def _parse_netsh_adapter_details(output: str) -> dict:
        details = {}
        for line in output.splitlines():
            stripped = line.strip()
            if stripped.startswith("DHCP enabled:"):
                value = stripped.split(":", 1)[1].strip()
                details["dhcp"] = "Enabled" if value.lower().startswith("yes") else "Disabled"
            elif stripped.startswith("IP Address:"):
                details["ip"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("Subnet Prefix:") and "mask " in stripped:
                details["mask"] = stripped.rsplit("mask ", 1)[1].rstrip(")")
        return details

    def apply_static(self, adapter_name: str, ip_address: str, subnet_mask: str) -> tuple[bool, str]:
        command = [
            "netsh",
            "interface",
            "ipv4",
            "set",
            "address",
            f'name="{adapter_name}"',
            "source=static",
            f"address={ip_address}",
            f"mask={subnet_mask}",
            "gateway=none",
        ]
        return self.run_command(command)

    def set_dhcp(self, adapter_name: str) -> tuple[bool, str]:
        command = [
            "netsh",
            "interface",
            "ipv4",
            "set",
            "address",
            f'name="{adapter_name}"',
            "source=dhcp",
        ]
        return self.run_command(command)


class IPv4Input(ttk.Frame):
    def __init__(self, parent, style_name: str):
        super().__init__(parent, style=style_name)
        self.vars = [tk.StringVar() for _ in range(4)]
        self.entries: list[tk.Entry] = []
        self.configure(style=style_name)

        for index, var in enumerate(self.vars):
            entry = tk.Entry(
                self,
                textvariable=var,
                width=3,
                justify="center",
                relief="flat",
                bd=0,
                font=("Segoe UI Semibold", 9),
                bg="#f6f7fb",
                fg="#24324a",
                highlightthickness=1,
                highlightbackground="#d8e1f0",
                highlightcolor="#1e90ff",
                insertbackground="#24324a",
            )
            entry.grid(row=0, column=index * 2, padx=(0, 0), pady=0, ipady=5)
            entry.bind("<KeyRelease>", lambda event, idx=index: self._on_key(event, idx))
            entry.bind("<FocusOut>", lambda _event, idx=index: self._normalize_octet(idx))
            self.entries.append(entry)
            if index < 3:
                dot = ttk.Label(self, text=".", style="Dot.TLabel")
                dot.grid(row=0, column=index * 2 + 1, padx=8)

    def _on_key(self, event, index: int) -> None:
        value = self.vars[index].get()
        digits = "".join(char for char in value if char.isdigit())[:3]
        if digits != value:
            self.vars[index].set(digits)

        if len(digits) == 3 and index < 3 and event.keysym not in {"BackSpace", "Left"}:
            self.entries[index + 1].focus_set()
            self.entries[index + 1].selection_range(0, tk.END)

    def _normalize_octet(self, index: int) -> None:
        value = self.vars[index].get().strip()
        if value == "":
            return
        try:
            octet = int(value)
        except ValueError:
            self.vars[index].set("")
            return
        if octet < 0 or octet > 255:
            self.vars[index].set("")
        else:
            self.vars[index].set(str(octet))

    def get(self) -> str:
        return ".".join(var.get().strip() for var in self.vars)

    def set(self, value: str) -> None:
        parts = value.split(".") if value else []
        for index, var in enumerate(self.vars):
            var.set(parts[index] if index < len(parts) else "")

    def clear(self) -> None:
        self.set("")

    def set_state(self, disabled: bool) -> None:
        state = "disabled" if disabled else "normal"
        bg = "#eaedf4" if disabled else "#f6f7fb"
        fg = "#8b97aa" if disabled else "#24324a"
        for entry in self.entries:
            entry.configure(state=state, bg=bg, fg=fg)


class PingRunner:
    @staticmethod
    def ping_once(target: str) -> tuple[bool, str]:
        try:
            completed = subprocess.run(
                ["ping", "-n", "1", "-w", "1000", target],
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            return False, str(exc)

        output = ((completed.stdout or "") + "\n" + (completed.stderr or "")).strip()
        match = re.search(r"time[=<]\s*(\d+)\s*ms", output, flags=re.IGNORECASE)
        if match:
            return True, f"{match.group(1)} ms"
        if "TTL=" in output.upper():
            return True, "<1 ms"
        return False, "Timeout"


class IpChangerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x520")
        self.minsize(760, 500)
        self.configure(bg="#0e1726")
        self.app_icon = None
        self._set_app_icon()

        self.store = ProfileStore(PROFILE_FILE)
        self.ping_store = ProfileStore(PING_PROFILE_FILE)
        self.network = NetworkManager()
        self.ping_runner = PingRunner()
        self.profiles = self.store.load()
        self.ping_profiles = self.ping_store.load()
        self.selected_profile_index = None
        self.selected_ping_profile_index = None
        self.connected = False
        self.active_adapter = None
        self.ping_running = False
        self.ping_thread = None
        self.ping_stop_event = threading.Event()
        self.adapter_queue = queue.Queue()
        self.tray_queue = queue.Queue()
        self.tray_icon = None
        self.tray_thread = None
        self.tray_notification_shown = False

        self.profile_selector_var = tk.StringVar(value=NEW_PROFILE_LABEL)
        self.profile_name_var = tk.StringVar()
        self.adapter_var = tk.StringVar()
        self.use_dhcp_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Disconnected")
        self.detail_var = tk.StringVar(value="Profiles only fill the form. Press Connect to apply settings.")
        self.current_ip_var = tk.StringVar(value="-")
        self.current_mask_var = tk.StringVar(value="-")
        self.current_mode_var = tk.StringVar(value="-")
        self.connect_label_var = tk.StringVar(value="Connect")
        self.ping_profile_selector_var = tk.StringVar(value=NEW_PROFILE_LABEL)
        self.ping_profile_name_var = tk.StringVar()
        self.ping_target_var = tk.StringVar()
        self.ping_duration_var = tk.StringVar(value="10")
        self.ping_button_var = tk.StringVar(value="Start Ping")
        self.ping_status_var = tk.StringVar(value="Idle")
        self.ping_target_status_var = tk.StringVar(value="-")
        self.ping_reply_status_var = tk.StringVar(value="-")
        self.ping_left_status_var = tk.StringVar(value="-")

        self.profile_selector_combo = None
        self.ping_profile_selector_combo = None
        self.ping_console = None
        self.hero_badge = None
        self.adapter_combo = None
        self.ip_input = None
        self.subnet_input = None
        self.delete_button = None
        self.delete_ping_button = None

        self._configure_style()
        self._build_ui()
        self.refresh_profile_selector()
        self.refresh_ping_profile_selector()
        self._toggle_ip_fields()
        self.status_var.set("Loading...")
        self.detail_var.set("Loading network adapters in the background.")
        self.after(100, self.refresh_adapters_async)
        self.after(100, self.poll_tray_queue)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Unmap>", self.on_minimize)
        self.create_tray_icon()

    def _set_app_icon(self) -> None:
        if not APP_ICON_FILE.exists():
            return
        try:
            self.app_icon = tk.PhotoImage(file=str(APP_ICON_FILE))
            self.iconphoto(True, self.app_icon)
        except tk.TclError:
            self.app_icon = None

    def on_minimize(self, _event=None) -> None:
        if self.state() != "iconic":
            return
        self.after(0, self.hide_to_tray)

    def hide_to_tray(self) -> None:
        self.withdraw()
        self.create_tray_icon()
        if not self.tray_notification_shown:
            self.notify_tray("IPSwitch-Pro is still running in the system tray.")
            self.tray_notification_shown = True

    def create_tray_icon(self) -> None:
        if self.tray_icon is not None:
            self.tray_icon.update_menu()
            return
        image = self.create_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("Show", lambda _icon, _item: self.tray_queue.put("show"), default=True),
            pystray.MenuItem("Connect", lambda _icon, _item: self.tray_queue.put("connect"), enabled=lambda _item: not self.connected),
            pystray.MenuItem("Disconnect", lambda _icon, _item: self.tray_queue.put("disconnect"), enabled=lambda _item: self.connected),
            pystray.MenuItem("Exit", lambda _icon, _item: self.tray_queue.put("exit")),
        )
        self.tray_icon = pystray.Icon(APP_TITLE, image, APP_TITLE, menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def create_tray_image(self) -> Image.Image:
        color = "#1d4ed8" if self.connected else "#64748b"
        image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        from PIL import ImageDraw, ImageFont
        draw = ImageDraw.Draw(image)
        draw.ellipse((6, 6, 58, 58), fill=color)
        try:
            font = ImageFont.truetype("segoeuib.ttf", 20)
        except OSError:
            font = ImageFont.load_default()
        draw.text((32, 32), "IP", fill="white", font=font, anchor="mm")
        return image

    def poll_tray_queue(self) -> None:
        try:
            command = self.tray_queue.get_nowait()
        except queue.Empty:
            self.after(100, self.poll_tray_queue)
            return

        if command == "show":
            self.show_from_tray()
        elif command == "connect":
            self.tray_connect()
        elif command == "disconnect":
            self.tray_disconnect()
        elif command == "exit":
            self.exit_from_tray()

        self.after(100, self.poll_tray_queue)

    def show_from_tray(self) -> None:
        self.deiconify()
        self.state("normal")
        self.lift()
        self.focus_force()

    def tray_connect(self) -> None:
        if self.connected:
            return
        if self.selected_profile_index is None:
            self.show_from_tray()
            messagebox.showwarning("Select Profile", "Please select a network profile first.")
            return
        self.connect_profile()
        self.refresh_tray_menu()

    def tray_disconnect(self) -> None:
        if self.connected and self.active_adapter:
            self.disconnect_adapter(self.active_adapter, show_message=False)
            self.refresh_tray_menu()

    def exit_from_tray(self) -> None:
        self.remove_tray_icon()
        self.on_close()

    def remove_tray_icon(self) -> None:
        if self.tray_icon is None:
            return
        icon = self.tray_icon
        self.tray_icon = None
        icon.stop()

    def refresh_tray_menu(self) -> None:
        if self.tray_icon is not None:
            self.tray_icon.icon = self.create_tray_image()
            self.tray_icon.update_menu()

    def notify_tray(self, message: str) -> None:
        if self.tray_icon is None:
            return
        try:
            self.tray_icon.notify(message, APP_TITLE)
        except Exception:
            pass

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#0e1726")
        style.configure("Panel.TFrame", background="#122033")
        style.configure("Hero.TFrame", background="#162a44")
        style.configure("Card.TFrame", background="#f8fafc")
        style.configure("Accent.TFrame", background="#13243b")
        style.configure("InputWrap.TFrame", background="#f6f7fb")

        style.configure("Title.TLabel", background="#0e1726", foreground="#f8fafc", font=("Segoe UI Semibold", 17))
        style.configure("Subtitle.TLabel", background="#0e1726", foreground="#94a3b8", font=("Segoe UI", 8))
        style.configure("Section.TLabel", background="#f8fafc", foreground="#23314e", font=("Segoe UI Semibold", 11))
        style.configure("Body.TLabel", background="#f8fafc", foreground="#5b6980", font=("Segoe UI", 8))
        style.configure("Value.TLabel", background="#f8fafc", foreground="#0f172a", font=("Segoe UI Semibold", 8))
        style.configure("Link.TLabel", background="#f8fafc", foreground="#1d4ed8", font=("Segoe UI Semibold", 8, "underline"))
        style.configure("StatusBig.TLabel", background="#122033", foreground="#f8fafc", font=("Segoe UI Semibold", 15))
        style.configure("HeroBadgeConnected.TLabel", background="#1d4ed8", foreground="#ffffff", font=("Segoe UI Semibold", 10))
        style.configure("HeroBadgeDisconnected.TLabel", background="#64748b", foreground="#f8fafc", font=("Segoe UI Semibold", 10))
        style.configure("HeroTitle.TLabel", background="#162a44", foreground="#e0f2fe", font=("Segoe UI Semibold", 10))
        style.configure("HeroSub.TLabel", background="#162a44", foreground="#93c5fd", font=("Segoe UI", 7))
        style.configure("Dot.TLabel", background="#f8fafc", foreground="#6b7c93", font=("Segoe UI Semibold", 12))
        style.configure("DarkSection.TLabel", background="#122033", foreground="#dbeafe", font=("Segoe UI Semibold", 9))
        style.configure("DarkBody.TLabel", background="#122033", foreground="#bfd3ea", font=("Segoe UI", 8))
        style.configure("DarkValue.TLabel", background="#122033", foreground="#ffffff", font=("Segoe UI Semibold", 8))

        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 8),
            padding=(10, 7),
            background="#1d4ed8",
            foreground="#ffffff",
            borderwidth=0,
        )
        style.map("Primary.TButton", background=[("active", "#1e40af")])

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 8),
            padding=(10, 7),
            background="#e9eef7",
            foreground="#334155",
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#dbe4f0")])

        style.configure(
            "Danger.TButton",
            font=("Segoe UI Semibold", 8),
            padding=(10, 7),
            background="#e11d48",
            foreground="#ffffff",
            borderwidth=0,
        )
        style.map("Danger.TButton", background=[("active", "#be123c")])
        style.configure(
            "Success.TButton",
            font=("Segoe UI Semibold", 8),
            padding=(10, 7),
            background="#059669",
            foreground="#ffffff",
            borderwidth=0,
        )
        style.map("Success.TButton", background=[("active", "#047857")])

        style.configure(
            "App.TCombobox",
            fieldbackground="#f6f7fb",
            background="#f6f7fb",
            foreground="#24324a",
            bordercolor="#d8e1f0",
            lightcolor="#d8e1f0",
            darkcolor="#d8e1f0",
            arrowcolor="#1d4ed8",
            padding=6,
        )
    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top_bar = ttk.Frame(self, style="App.TFrame", padding=(18, 12, 18, 8))
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.grid_columnconfigure(0, weight=1)

        ttk.Label(top_bar, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top_bar, text="Manage IP profiles, switch networks fast, and run live ping tests", style="Subtitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        main = ttk.Frame(self, style="App.TFrame", padding=(14, 0, 14, 14))
        main.grid(row=1, column=0, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=3)
        main.grid_rowconfigure(0, weight=1)

        self._build_left_panel(main)
        self._build_right_panel(main)

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        left = ttk.Frame(parent, style="Panel.TFrame", padding=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(2, weight=1)
        left.grid_rowconfigure(3, weight=1)

        hero = ttk.Frame(left, style="Hero.TFrame", padding=(10, 8))
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        hero.grid_columnconfigure(1, weight=1)
        self.hero_badge = ttk.Label(hero, text="IP", style="HeroBadgeDisconnected.TLabel", anchor="center", width=4)
        self.hero_badge.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 10))
        ttk.Label(hero, text="Network Control", style="HeroTitle.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(hero, text="Switch profiles safely", style="HeroSub.TLabel").grid(row=1, column=1, sticky="w", pady=(2, 0))

        ttk.Label(left, textvariable=self.status_var, style="StatusBig.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 6))

        status_card = ttk.Frame(left, style="Card.TFrame", padding=10)
        status_card.grid(row=2, column=0, sticky="nsew")
        status_card.grid_columnconfigure(1, weight=1)

        ttk.Label(status_card, text="Current Status", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(status_card, textvariable=self.detail_var, style="Body.TLabel", wraplength=210, justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(5, 8)
        )
        ttk.Label(status_card, text="IP Address", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Label(status_card, textvariable=self.current_ip_var, style="Value.TLabel").grid(row=2, column=1, sticky="e", pady=3)
        ttk.Label(status_card, text="Subnet Mask", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Label(status_card, textvariable=self.current_mask_var, style="Value.TLabel").grid(row=3, column=1, sticky="e", pady=3)
        ttk.Label(status_card, text="Mode", style="Body.TLabel").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Label(status_card, textvariable=self.current_mode_var, style="Value.TLabel").grid(row=4, column=1, sticky="e", pady=3)

        ping_card = ttk.Frame(left, style="Panel.TFrame", padding=8)
        ping_card.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
        ping_card.grid_columnconfigure(1, weight=1)
        ttk.Label(ping_card, text="Live Ping", style="DarkSection.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(ping_card, text="Status", style="DarkBody.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 2))
        ttk.Label(ping_card, textvariable=self.ping_status_var, style="DarkValue.TLabel").grid(row=1, column=1, sticky="e", pady=(6, 2))
        ttk.Label(ping_card, text="Target", style="DarkBody.TLabel").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(ping_card, textvariable=self.ping_target_status_var, style="DarkValue.TLabel").grid(row=2, column=1, sticky="e", pady=2)
        ttk.Label(ping_card, text="Last", style="DarkBody.TLabel").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Label(ping_card, textvariable=self.ping_reply_status_var, style="DarkValue.TLabel").grid(row=3, column=1, sticky="e", pady=2)
        ttk.Label(ping_card, text="Left", style="DarkBody.TLabel").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Label(ping_card, textvariable=self.ping_left_status_var, style="DarkValue.TLabel").grid(row=4, column=1, sticky="e", pady=2)

        action_row = ttk.Frame(left, style="Panel.TFrame")
        action_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        action_row.grid_columnconfigure((0, 1), weight=1)

        ttk.Button(action_row, textvariable=self.connect_label_var, style="Primary.TButton", command=self.toggle_connection).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(action_row, text="Refresh", style="Secondary.TButton", command=self.refresh_current_status).grid(
            row=0, column=1, sticky="ew"
        )

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        right = ttk.Frame(parent, style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        tabs = ttk.Notebook(right)
        tabs.grid(row=0, column=0, sticky="nsew")

        network_tab = ttk.Frame(tabs, style="Card.TFrame", padding=14)
        ping_tab = ttk.Frame(tabs, style="Card.TFrame", padding=14)
        about_tab = ttk.Frame(tabs, style="Card.TFrame", padding=14)
        tabs.add(network_tab, text="Network")
        tabs.add(ping_tab, text="Ping")
        tabs.add(about_tab, text="About / Help")

        self._build_network_tab(network_tab)
        self._build_ping_tab(ping_tab)
        self._build_about_tab(about_tab)

    def _build_network_tab(self, editor: ttk.Frame) -> None:
        editor.grid_columnconfigure(0, weight=1)

        ttk.Label(editor, text="Profile Editor", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(editor, text="Selecting a profile only fills the form. Nothing is applied until you press Connect.", style="Body.TLabel", wraplength=340).grid(
            row=1, column=0, sticky="w", pady=(4, 8)
        )

        fields = ttk.Frame(editor, style="Card.TFrame")
        fields.grid(row=2, column=0, sticky="ew")
        fields.grid_columnconfigure(1, weight=1)

        self._labeled_profile_selector(fields, 0)
        self._labeled_entry(fields, "Profile Name", self.profile_name_var, 1)
        self._labeled_adapter_selector(fields, 2)
        self._labeled_ipv4(fields, "IP Address", "ip", 3)
        self._labeled_ipv4(fields, "Subnet Mask", "subnet", 4)

        ttk.Checkbutton(fields, text="Use DHCP instead of a static IP", variable=self.use_dhcp_var, command=self._toggle_ip_fields).grid(
            row=5, column=1, sticky="w", pady=(4, 0)
        )

        buttons = ttk.Frame(editor, style="Card.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        buttons.grid_columnconfigure((0, 1, 2), weight=1)

        ttk.Button(buttons, text="Save Profile", style="Primary.TButton", command=self.save_profile).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.delete_button = ttk.Button(buttons, text="Delete Profile", style="Danger.TButton", command=self.delete_profile)
        self.delete_button.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(buttons, text="New / Clear", style="Secondary.TButton", command=self.switch_to_new_profile).grid(row=0, column=2, sticky="ew", padx=(8, 0))

    def _build_ping_tab(self, editor: ttk.Frame) -> None:
        editor.grid_columnconfigure(0, weight=1)

        ttk.Label(editor, text="Ping Profiles", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(editor, text="Save targets and test duration, then start live ping from this tab.", style="Body.TLabel", wraplength=340).grid(
            row=1, column=0, sticky="w", pady=(4, 8)
        )

        fields = ttk.Frame(editor, style="Card.TFrame")
        fields.grid(row=2, column=0, sticky="ew")
        fields.grid_columnconfigure(1, weight=1)

        self._labeled_ping_profile_selector(fields, 0)
        self._labeled_entry(fields, "Profile Name", self.ping_profile_name_var, 1)
        self._labeled_entry(fields, "Target IP / Host", self.ping_target_var, 2)
        self._labeled_entry(fields, "Duration (sec)", self.ping_duration_var, 3)

        buttons = ttk.Frame(editor, style="Card.TFrame")
        buttons.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        buttons.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ttk.Button(buttons, text="Save Profile", style="Primary.TButton", command=self.save_ping_profile).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.delete_ping_button = ttk.Button(buttons, text="Delete Profile", style="Danger.TButton", command=self.delete_ping_profile)
        self.delete_ping_button.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(buttons, textvariable=self.ping_button_var, style="Success.TButton", command=self.toggle_ping).grid(row=0, column=2, sticky="ew", padx=8)
        ttk.Button(buttons, text="New / Clear", style="Secondary.TButton", command=self.switch_to_new_ping_profile).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        console_wrap = ttk.Frame(editor, style="Card.TFrame")
        console_wrap.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        console_wrap.grid_columnconfigure(0, weight=1)
        console_wrap.grid_rowconfigure(1, weight=1)
        ttk.Label(console_wrap, text="CMD Output", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.ping_console = tk.Text(
            console_wrap,
            height=7,
            bg="#050b14",
            fg="#9ef7c3",
            insertbackground="#9ef7c3",
            relief="flat",
            bd=0,
            font=("Consolas", 8),
            wrap="word",
        )
        self.ping_console.grid(row=1, column=0, sticky="nsew")
        self.ping_console.insert("end", "Ping output will appear here...\n")
        self.ping_console.tag_configure("ok", foreground="#9ef7c3")
        self.ping_console.tag_configure("error", foreground="#ff5c7a")
        self.ping_console.tag_configure("info", foreground="#93c5fd")
        self.ping_console.configure(state="disabled")

    def _build_about_tab(self, editor: ttk.Frame) -> None:
        editor.grid_columnconfigure(0, weight=1)

        ttk.Label(editor, text="Help & About", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        help_text = (
            "This app helps you save and apply Windows network IP profiles, switch adapters back to automatic DHCP, "
            "and run live ping tests from saved ping profiles.\n\n"
            "Network tab: create a profile, choose an adapter, enter IP/Subnet or DHCP, then press Connect on the left.\n\n"
            "Ping tab: save target IP/host profiles, choose a duration, then start a live CMD-style ping test."
        )
        ttk.Label(editor, text=help_text, style="Body.TLabel", wraplength=470, justify="left").grid(
            row=1, column=0, sticky="w", pady=(8, 14)
        )

        creator_card = ttk.Frame(editor, style="Card.TFrame")
        creator_card.grid(row=2, column=0, sticky="ew")
        creator_card.grid_columnconfigure(1, weight=1)

        ttk.Label(creator_card, text="Creator", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(creator_card, text="Name", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=3, padx=(0, 14))
        ttk.Label(creator_card, text="Siahtiri", style="Value.TLabel").grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(creator_card, text="Email", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=3, padx=(0, 14))
        ttk.Label(creator_card, text="siahtirim@gmail.com", style="Value.TLabel").grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(creator_card, text="Websites", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=3, padx=(0, 14))
        website_links = ttk.Frame(creator_card, style="Card.TFrame")
        website_links.grid(row=3, column=1, sticky="w", pady=3)
        self._link_label(website_links, "https://autonexit.com", "https://autonexit.com").grid(row=0, column=0, sticky="w")
        ttk.Label(website_links, text=" / ", style="Value.TLabel").grid(row=0, column=1, sticky="w")
        self._link_label(website_links, "https://poweren.ir", "https://poweren.ir").grid(row=0, column=2, sticky="w")
        ttk.Label(creator_card, text="GitHub (Check update)", style="Body.TLabel").grid(row=4, column=0, sticky="w", pady=3, padx=(0, 14))
        self._link_label(creator_card, GITHUB_URL, GITHUB_URL).grid(row=4, column=1, sticky="w", pady=3)

        open_source_text = (
            "This project is free and open-source. You can use it, modify it, "
            "and share it under the terms provided in the project repository."
        )
        ttk.Label(editor, text=open_source_text, style="Body.TLabel", wraplength=470, justify="left").grid(
            row=3, column=0, sticky="w", pady=(16, 0)
        )

    def _link_label(self, parent: ttk.Frame, text: str, url: str) -> ttk.Label:
        label = ttk.Label(parent, text=text, style="Link.TLabel", cursor="hand2")
        label.bind("<Button-1>", lambda _event: webbrowser.open_new_tab(url))
        return label

    def _labeled_profile_selector(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Profile", style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 14))
        self.profile_selector_combo = ttk.Combobox(parent, textvariable=self.profile_selector_var, style="App.TCombobox", state="readonly", font=("Segoe UI", 8))
        self.profile_selector_combo.grid(row=row, column=1, sticky="ew", pady=7)
        self.profile_selector_combo.bind("<<ComboboxSelected>>", self.on_profile_change)

    def _labeled_ping_profile_selector(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Profile", style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 14))
        self.ping_profile_selector_combo = ttk.Combobox(parent, textvariable=self.ping_profile_selector_var, style="App.TCombobox", state="readonly", font=("Segoe UI", 8))
        self.ping_profile_selector_combo.grid(row=row, column=1, sticky="ew", pady=7)
        self.ping_profile_selector_combo.bind("<<ComboboxSelected>>", self.on_ping_profile_change)

    def _labeled_adapter_selector(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Network Adapter", style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 14))
        self.adapter_combo = ttk.Combobox(parent, textvariable=self.adapter_var, style="App.TCombobox", state="readonly", font=("Segoe UI", 8))
        self.adapter_combo.grid(row=row, column=1, sticky="ew", pady=7)
        self.adapter_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_current_status())

    def _labeled_entry(self, parent: ttk.Frame, text: str, variable: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=text, style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 14))
        entry = ttk.Entry(parent, textvariable=variable, font=("Segoe UI", 8))
        entry.grid(row=row, column=1, sticky="ew", pady=5, ipady=4)

    def _labeled_ipv4(self, parent: ttk.Frame, text: str, kind: str, row: int) -> None:
        ttk.Label(parent, text=text, style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 14))
        holder = ttk.Frame(parent, style="InputWrap.TFrame", padding=6)
        holder.grid(row=row, column=1, sticky="w", pady=5)
        widget = IPv4Input(holder, "InputWrap.TFrame")
        widget.grid(row=0, column=0, sticky="w")
        if kind == "ip":
            self.ip_input = widget
        else:
            self.subnet_input = widget

    def refresh_profile_selector(self) -> None:
        names = [profile.get("name", "Unnamed Profile") for profile in self.profiles]
        self.profile_selector_combo["values"] = [NEW_PROFILE_LABEL, *names]
        if self.profile_selector_var.get() not in self.profile_selector_combo["values"]:
            self.profile_selector_var.set(NEW_PROFILE_LABEL)
        if self.delete_button is not None:
            self.delete_button.configure(state="normal" if self.selected_profile_index is not None else "disabled")

    def refresh_ping_profile_selector(self) -> None:
        names = [profile.get("name", "Unnamed Profile") for profile in self.ping_profiles]
        self.ping_profile_selector_combo["values"] = [NEW_PROFILE_LABEL, *names]
        if self.ping_profile_selector_var.get() not in self.ping_profile_selector_combo["values"]:
            self.ping_profile_selector_var.set(NEW_PROFILE_LABEL)
        if self.delete_ping_button is not None:
            self.delete_ping_button.configure(state="normal" if self.selected_ping_profile_index is not None else "disabled")

    def refresh_adapters(self) -> None:
        adapters = self.network.list_adapters()
        self.adapter_combo["values"] = adapters
        if adapters and self.adapter_var.get() not in adapters:
            self.adapter_var.set(adapters[0])
        elif not adapters:
            self.adapter_var.set("")
        self.refresh_current_status()

    def refresh_adapters_async(self) -> None:
        self.status_var.set("Loading...")
        self.detail_var.set("Loading network adapters in the background.")
        threading.Thread(target=self._load_adapters_worker, daemon=True).start()
        self.after(100, self._poll_adapter_queue)

    def _load_adapters_worker(self) -> None:
        adapters = self.network.list_adapters()
        self.adapter_queue.put(adapters)

    def _poll_adapter_queue(self) -> None:
        try:
            adapters = self.adapter_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_adapter_queue)
            return
        self._apply_loaded_adapters(adapters)

    def _apply_loaded_adapters(self, adapters: list[str]) -> None:
        self.adapter_combo["values"] = adapters
        if adapters and self.adapter_var.get() not in adapters:
            self.adapter_var.set(adapters[0])
        elif not adapters:
            self.adapter_var.set("")
        self.refresh_current_status()

    def refresh_current_status(self) -> None:
        self.update_connection_badge()
        adapter = (self.active_adapter if self.connected else self.adapter_var.get()).strip()
        if not adapter:
            self.status_var.set("Disconnected")
            self.detail_var.set("Choose a profile, then press Connect to apply its settings.")
            self.current_ip_var.set("-")
            self.current_mask_var.set("-")
            self.current_mode_var.set("-")
            self.connect_label_var.set("Connect")
            return

        details = self.network.get_adapter_details(adapter)
        self.connect_label_var.set("Disconnect" if self.connected else "Connect")
        if details:
            self.current_ip_var.set(details.get("ip") or "-")
            self.current_mask_var.set(details.get("mask") or "-")
            dhcp_enabled = (details.get("dhcp") or "").lower() == "enabled"
            self.current_mode_var.set("DHCP" if dhcp_enabled else "Static")
            self.status_var.set("Connected" if self.connected else "Ready")
            privilege = "Administrator" if self.network.is_admin() else "Standard user"
            if self.connected:
                self.detail_var.set(f"Profile is applied on {adapter}. Running as: {privilege}.")
            else:
                self.detail_var.set(f"Windows settings loaded for {adapter}. Nothing has been changed yet.")
            return

        self.status_var.set("Connected" if self.connected else "Ready")
        self.detail_var.set("Could not read the current IP details for this adapter.")
        self.current_ip_var.set("-")
        self.current_mask_var.set("-")
        self.current_mode_var.set("-")

    def update_connection_badge(self) -> None:
        if self.hero_badge is None:
            return
        badge_style = "HeroBadgeConnected.TLabel" if self.connected else "HeroBadgeDisconnected.TLabel"
        self.hero_badge.configure(style=badge_style)

    def on_profile_change(self, _event=None) -> None:
        label = self.profile_selector_var.get()
        if label == NEW_PROFILE_LABEL:
            self.switch_to_new_profile()
            return

        for index, profile in enumerate(self.profiles):
            if profile.get("name") == label:
                self.selected_profile_index = index
                self.load_profile(profile)
                self.refresh_profile_selector()
                return

    def on_ping_profile_change(self, _event=None) -> None:
        label = self.ping_profile_selector_var.get()
        if label == NEW_PROFILE_LABEL:
            self.switch_to_new_ping_profile()
            return

        for index, profile in enumerate(self.ping_profiles):
            if profile.get("name") == label:
                self.selected_ping_profile_index = index
                self.load_ping_profile(profile)
                self.refresh_ping_profile_selector()
                return

    def switch_to_new_profile(self) -> None:
        if self.connected and self.active_adapter:
            self.disconnect_adapter(self.active_adapter, show_message=False)
        self.selected_profile_index = None
        self.profile_selector_var.set(NEW_PROFILE_LABEL)
        self.clear_form(reset_profile_selector=False)
        self.refresh_profile_selector()

    def switch_to_new_ping_profile(self) -> None:
        self.stop_ping(show_message=False)
        self.selected_ping_profile_index = None
        self.ping_profile_selector_var.set(NEW_PROFILE_LABEL)
        self.clear_ping_form(reset_profile_selector=False)
        self.refresh_ping_profile_selector()

    def load_profile(self, profile: dict) -> None:
        self.profile_name_var.set(profile.get("name", ""))
        self.adapter_var.set(profile.get("adapter", ""))
        self.ip_input.set(profile.get("ip_address", ""))
        self.subnet_input.set(profile.get("subnet_mask", ""))
        self.use_dhcp_var.set(bool(profile.get("use_dhcp", False)))
        self._toggle_ip_fields()
        self.refresh_current_status()

    def load_ping_profile(self, profile: dict) -> None:
        self.ping_profile_name_var.set(profile.get("name", ""))
        self.ping_target_var.set(profile.get("target", ""))
        self.ping_duration_var.set(str(profile.get("duration", 10)))

    def _toggle_ip_fields(self) -> None:
        disabled = self.use_dhcp_var.get()
        self.ip_input.set_state(disabled)
        self.subnet_input.set_state(disabled)

    def _current_ip_value(self) -> str:
        return self.ip_input.get()

    def _current_subnet_value(self) -> str:
        return self.subnet_input.get()

    def validate_inputs(self) -> bool:
        if not self.profile_name_var.get().strip():
            messagebox.showerror("Validation Error", "Profile name is required.")
            return False
        if not self.adapter_var.get().strip():
            messagebox.showerror("Validation Error", "Please select a network adapter.")
            return False
        if self.use_dhcp_var.get():
            return True

        ip_value = self._current_ip_value()
        subnet_value = self._current_subnet_value()
        if "" in ip_value.split(".") or "" in subnet_value.split("."):
            messagebox.showerror("Validation Error", "Complete all IP and subnet fields.")
            return False

        try:
            ipaddress.IPv4Address(ip_value)
        except ipaddress.AddressValueError:
            messagebox.showerror("Validation Error", "Invalid IPv4 address.")
            return False

        try:
            ipaddress.IPv4Network(f"0.0.0.0/{subnet_value}")
        except ValueError:
            messagebox.showerror("Validation Error", "Invalid subnet mask.")
            return False

        return True

    def validate_ping_inputs(self) -> bool:
        if not self.ping_profile_name_var.get().strip():
            messagebox.showerror("Validation Error", "Ping profile name is required.")
            return False
        if not self.ping_target_var.get().strip():
            messagebox.showerror("Validation Error", "Enter an IP address or host name to ping.")
            return False
        try:
            duration = int(self.ping_duration_var.get().strip())
        except ValueError:
            messagebox.showerror("Validation Error", "Duration must be a whole number.")
            return False
        if duration <= 0:
            messagebox.showerror("Validation Error", "Duration must be greater than zero.")
            return False
        return True

    def save_profile(self) -> None:
        if not self.validate_inputs():
            return

        profile = {
            "name": self.profile_name_var.get().strip(),
            "adapter": self.adapter_var.get().strip(),
            "ip_address": self._current_ip_value(),
            "subnet_mask": self._current_subnet_value(),
            "use_dhcp": self.use_dhcp_var.get(),
        }

        existing_index = None
        for index, existing_profile in enumerate(self.profiles):
            if existing_profile.get("name", "").lower() == profile["name"].lower():
                existing_index = index
                break

        if self.selected_profile_index is not None and self.selected_profile_index < len(self.profiles):
            self.profiles[self.selected_profile_index] = profile
        elif existing_index is not None:
            self.profiles[existing_index] = profile
            self.selected_profile_index = existing_index
        else:
            self.profiles.append(profile)
            self.selected_profile_index = len(self.profiles) - 1

        self.store.save(self.profiles)
        self.profile_selector_var.set(profile["name"])
        self.refresh_profile_selector()
        messagebox.showinfo("Saved", "Profile saved. It will only apply when you press Connect.")

    def save_ping_profile(self) -> None:
        if not self.validate_ping_inputs():
            return

        profile = {
            "name": self.ping_profile_name_var.get().strip(),
            "target": self.ping_target_var.get().strip(),
            "duration": int(self.ping_duration_var.get().strip()),
        }

        existing_index = None
        for idx, item in enumerate(self.ping_profiles):
            if item.get("name", "").lower() == profile["name"].lower():
                existing_index = idx
                break

        if self.selected_ping_profile_index is not None and self.selected_ping_profile_index < len(self.ping_profiles):
            self.ping_profiles[self.selected_ping_profile_index] = profile
        elif existing_index is not None:
            self.ping_profiles[existing_index] = profile
            self.selected_ping_profile_index = existing_index
        else:
            self.ping_profiles.append(profile)
            self.selected_ping_profile_index = len(self.ping_profiles) - 1

        self.ping_store.save(self.ping_profiles)
        self.ping_profile_selector_var.set(profile["name"])
        self.refresh_ping_profile_selector()
        messagebox.showinfo("Saved", "Ping profile saved.")

    def toggle_connection(self) -> None:
        if self.connected:
            if self.active_adapter:
                self.disconnect_adapter(self.active_adapter, show_message=True)
            return

        self.connect_profile()

    def connect_profile(self) -> None:
        if not self.validate_inputs():
            return

        adapter = self.adapter_var.get().strip()
        if self.use_dhcp_var.get():
            success, output = self.network.set_dhcp(adapter)
        else:
            success, output = self.network.apply_static(adapter, self._current_ip_value(), self._current_subnet_value())

        if not success:
            self._show_apply_error(output)
            return

        self.connected = True
        self.active_adapter = adapter
        self.status_var.set("Connected")
        self.detail_var.set(f"Profile applied on {adapter}.")
        self.connect_label_var.set("Disconnect")
        self.current_mode_var.set("DHCP" if self.use_dhcp_var.get() else "Static")
        self.current_ip_var.set("-" if self.use_dhcp_var.get() else self._current_ip_value())
        self.current_mask_var.set("-" if self.use_dhcp_var.get() else self._current_subnet_value())
        self.update_connection_badge()
        self.refresh_tray_menu()

    def disconnect_adapter(self, adapter: str, show_message: bool) -> None:
        success, output = self.network.set_dhcp(adapter)
        if not success:
            self._show_apply_error(output, title="Disconnect Failed")
            return

        self.connected = False
        self.active_adapter = None
        self.status_var.set("Disconnected")
        self.detail_var.set("Windows network settings were returned to automatic DHCP.")
        self.connect_label_var.set("Connect")
        self.current_ip_var.set("-")
        self.current_mask_var.set("-")
        self.current_mode_var.set("DHCP")
        self.update_connection_badge()
        self.refresh_tray_menu()
        if show_message:
            self.status_var.set("Disconnected")

    def toggle_ping(self) -> None:
        if self.ping_running:
            self.stop_ping(show_message=True)
            return
        self.start_ping()

    def start_ping(self) -> None:
        if not self.validate_ping_inputs():
            return

        target = self.ping_target_var.get().strip()
        duration = int(self.ping_duration_var.get().strip())
        self.clear_ping_console()
        self.append_ping_console(f"Pinging {target} for {duration} seconds...", "info")
        self.ping_running = True
        self.ping_stop_event.clear()
        self.ping_button_var.set("Stop Ping")
        self.ping_status_var.set("Running")
        self.ping_target_status_var.set(target)
        self.ping_reply_status_var.set("Starting...")
        self.ping_left_status_var.set(f"{duration}s")

        self.ping_thread = threading.Thread(target=self._run_ping_session, args=(target, duration), daemon=True)
        self.ping_thread.start()

    def _run_ping_session(self, target: str, duration: int) -> None:
        end_time = time.time() + duration
        while not self.ping_stop_event.is_set():
            remaining = max(0, int(end_time - time.time()))
            if remaining <= 0:
                break
            success, result = self.ping_runner.ping_once(target)
            line = f"Reply from {target}: time={result}" if success else f"Request timed out for {target}"
            self.after(0, self._update_ping_status, target, "Online" if success else "Timeout", result, remaining, line, success)
            time.sleep(1)

        self.after(0, self._finish_ping_session)

    def _update_ping_status(self, target: str, status: str, result: str, remaining: int, line: str, success: bool) -> None:
        self.ping_status_var.set(status)
        self.ping_target_status_var.set(target)
        self.ping_reply_status_var.set(result)
        self.ping_left_status_var.set(f"{remaining}s")
        self.append_ping_console(line, "ok" if success else "error")

    def _finish_ping_session(self) -> None:
        if not self.ping_running:
            return
        self.ping_running = False
        self.ping_button_var.set("Start Ping")
        self.ping_status_var.set("Finished")
        self.ping_left_status_var.set("0s")
        self.append_ping_console("Ping session finished.", "info")

    def stop_ping(self, show_message: bool) -> None:
        if not self.ping_running:
            return
        self.ping_stop_event.set()
        self.ping_running = False
        self.ping_button_var.set("Start Ping")
        self.ping_status_var.set("Stopped")
        self.ping_left_status_var.set("-")
        self.append_ping_console("Ping session stopped.", "info")
        if show_message:
            messagebox.showinfo("Ping Stopped", "The ping session was stopped.")

    def clear_ping_console(self) -> None:
        if self.ping_console is None:
            return
        self.ping_console.configure(state="normal")
        self.ping_console.delete("1.0", "end")
        self.ping_console.configure(state="disabled")

    def append_ping_console(self, line: str, tag: str = "info") -> None:
        if self.ping_console is None:
            return
        self.ping_console.configure(state="normal")
        self.ping_console.insert("end", f"{line}\n", tag)
        lines = self.ping_console.get("1.0", "end-1c").splitlines()
        if len(lines) > 8:
            self.ping_console.delete("1.0", f"{len(lines) - 7}.0")
        self.ping_console.see("end")
        self.ping_console.configure(state="disabled")

    def _show_apply_error(self, output: str, title: str = "Apply Failed") -> None:
        error_text = output or "Unknown error"
        if not self.network.is_admin():
            error_text += "\n\nHint: Run the application as Administrator."
        messagebox.showerror(title, error_text)

    def delete_profile(self) -> None:
        if self.selected_profile_index is None or self.selected_profile_index >= len(self.profiles):
            messagebox.showwarning("Delete Profile", "Select a saved profile first.")
            return

        profile_name = self.profiles[self.selected_profile_index].get("name", "profile")
        if not messagebox.askyesno("Delete Profile", f"Delete '{profile_name}'?"):
            return

        del self.profiles[self.selected_profile_index]
        self.store.save(self.profiles)
        self.switch_to_new_profile()

    def delete_ping_profile(self) -> None:
        if self.selected_ping_profile_index is None or self.selected_ping_profile_index >= len(self.ping_profiles):
            messagebox.showwarning("Delete Profile", "Select a saved ping profile first.")
            return

        profile_name = self.ping_profiles[self.selected_ping_profile_index].get("name", "profile")
        if not messagebox.askyesno("Delete Profile", f"Delete '{profile_name}'?"):
            return

        del self.ping_profiles[self.selected_ping_profile_index]
        self.ping_store.save(self.ping_profiles)
        self.switch_to_new_ping_profile()

    def clear_form(self, reset_profile_selector: bool = True) -> None:
        self.profile_name_var.set("")
        self.ip_input.clear()
        self.subnet_input.clear()
        self.use_dhcp_var.set(False)
        if self.adapter_combo is not None and self.adapter_combo["values"]:
            self.adapter_var.set(self.adapter_combo["values"][0])
        else:
            self.adapter_var.set("")
        if reset_profile_selector:
            self.profile_selector_var.set(NEW_PROFILE_LABEL)
        self._toggle_ip_fields()
        self.refresh_current_status()

    def clear_ping_form(self, reset_profile_selector: bool = True) -> None:
        self.ping_profile_name_var.set("")
        self.ping_target_var.set("")
        self.ping_duration_var.set("10")
        if reset_profile_selector:
            self.ping_profile_selector_var.set(NEW_PROFILE_LABEL)

    def on_close(self) -> None:
        self.stop_ping(show_message=False)
        self.remove_tray_icon()
        self.destroy()


if __name__ == "__main__":
    app = IpChangerApp()
    app.mainloop()
