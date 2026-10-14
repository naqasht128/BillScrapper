#!/usr/bin/env python3
"""Bill Scrapper - Admin Serial Key Generator.

Generates new subscription serial keys via the Render Express API
and automatically copies them to the clipboard to send directly to buyers.

Supports both a Dark-Themed Tkinter GUI and CLI command-line mode.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import requests

# Default API Configuration (Can also be set via environment or GUI)
DEFAULT_API_URL = os.environ.get("LICENSE_API_URL", "https://your-render-app.onrender.com")
DEFAULT_ADMIN_SECRET = os.environ.get("ADMIN_SECRET", "")

# VS Code Dark+ Palette
BG_DARK = "#1e1e1e"
CARD_BG = "#252526"
BORDER_COLOR = "#383838"
TEXT_WHITE = "#ffffff"
TEXT_MAIN = "#cccccc"
TEXT_MUTED = "#8b949e"
ACCENT_GREEN = "#238636"
ACCENT_GREEN_HOVER = "#2ea043"
INPUT_BG = "#2d2d2d"
BTN_SEC_BG = "#3a3d41"
BTN_SEC_HOVER = "#45494e"
UI_FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FONT = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"

PLANS = [
    ("1_month", "1 Month  —  Rs. 300  (30 Days, 1 Device)"),
    ("3_months", "3 Months  —  Rs. 700  (90 Days, Up to 2 Devices)"),
    ("6_months", "6 Months  —  Rs. 1,200  (180 Days, Up to 2 Devices)"),
    ("1_year", "1 Year  —  Rs. 2,000  (365 Days, Up to 2 Devices)"),
]


def generate_key_api(api_url: str, admin_secret: str, plan_type: str) -> dict:
    """Call the remote API to generate a new key."""
    url = f"{api_url.rstrip('/')}/api/admin/generate"
    headers = {
        "Content-Type": "application/json",
        "x-admin-secret": admin_secret.strip(),
    }
    payload = {"plan_type": plan_type}
    resp = requests.post(url, json=payload, headers=headers, timeout=12)
    try:
        data = resp.json()
    except Exception:
        resp.raise_for_status()
        raise RuntimeError(f"Unexpected response from server: {resp.text}")

    if resp.status_code >= 400 or not data.get("success"):
        err = data.get("error") or data.get("message") or f"HTTP {resp.status_code}"
        raise RuntimeError(err)

    return data


class KeyGeneratorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bill Scrapper — Admin Key Generator")
        self.configure(bg=BG_DARK)
        self.geometry("560x650")
        self.minsize(520, 600)

        self._init_styles()
        self._build_ui()

    def _init_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure("Root.TFrame", background=BG_DARK)
        self.style.configure("Card.TFrame", background=CARD_BG)
        self.style.configure("Title.TLabel", background=BG_DARK, foreground=TEXT_WHITE, font=(UI_FONT, 14, "bold"))
        self.style.configure("Sub.TLabel", background=BG_DARK, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        self.style.configure("CardTitle.TLabel", background=CARD_BG, foreground=TEXT_WHITE, font=(UI_FONT, 10, "bold"))
        self.style.configure("CardText.TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.configure("CardMuted.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=(UI_FONT, 8))
        self.style.configure("TRadiobutton", background=CARD_BG, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.map("TRadiobutton", background=[("active", CARD_BG)], foreground=[("active", TEXT_WHITE)])

    def _build_ui(self):
        container = ttk.Frame(self, style="Root.TFrame", padding=(24, 20, 24, 20))
        container.pack(fill="both", expand=True)

        # Header
        ttk.Label(container, text="Serial Key Generator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(container, text="Generate customer license keys backed by Render & Supabase", style="Sub.TLabel").pack(anchor="w", pady=(2, 14))

        # API Configuration Card
        config_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        config_card.pack(fill="x", pady=(0, 14))

        ttk.Label(config_card, text="SERVER CONFIGURATION", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        # API URL
        ttk.Label(config_card, text="License Server URL:", style="CardText.TLabel").pack(anchor="w")
        self.api_url_var = tk.StringVar(value=DEFAULT_API_URL)
        api_entry = tk.Entry(
            config_card,
            textvariable=self.api_url_var,
            bg=INPUT_BG,
            fg=TEXT_WHITE,
            insertbackground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        api_entry.pack(fill="x", pady=(4, 10), ipady=4)

        # Admin Secret
        ttk.Label(config_card, text="Admin Secret (x-admin-secret):", style="CardText.TLabel").pack(anchor="w")
        self.secret_var = tk.StringVar(value=DEFAULT_ADMIN_SECRET)
        secret_box = ttk.Frame(config_card, style="Card.TFrame")
        secret_box.pack(fill="x", pady=(4, 0))

        self.secret_entry = tk.Entry(
            secret_box,
            textvariable=self.secret_var,
            bg=INPUT_BG,
            fg=TEXT_WHITE,
            show="*",
            insertbackground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        self.secret_entry.pack(side="left", fill="x", expand=True, ipady=4)

        self.show_secret_btn = tk.Button(
            secret_box,
            text="👁 Show",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 8),
            command=self._toggle_secret,
            padx=8,
            cursor="hand2",
        )
        self.show_secret_btn.pack(side="left", padx=(6, 0), ipady=2)

        # Plan Selection Card
        plan_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        plan_card.pack(fill="x", pady=(0, 14))

        ttk.Label(plan_card, text="SELECT SUBSCRIPTION PLAN", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 10))

        self.plan_var = tk.StringVar(value="1_month")
        for code, label in PLANS:
            r = ttk.Radiobutton(plan_card, text=label, value=code, variable=self.plan_var, style="TRadiobutton")
            r.pack(anchor="w", pady=3)

        # Generate Button
        self.gen_btn = tk.Button(
            container,
            text="⚡  GENERATE NEW SERIAL KEY",
            bg=ACCENT_GREEN,
            fg=TEXT_WHITE,
            activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 10, "bold"),
            command=self._on_generate,
            cursor="hand2",
            pady=8,
        )
        self.gen_btn.pack(fill="x", pady=(0, 14))

        # Result Display Card
        res_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        res_card.pack(fill="both", expand=True)

        ttk.Label(res_card, text="GENERATED SERIAL KEY", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 6))

        self.key_var = tk.StringVar(value="")
        key_box = ttk.Frame(res_card, style="Card.TFrame")
        key_box.pack(fill="x", pady=(0, 10))

        self.key_entry = tk.Entry(
            key_box,
            textvariable=self.key_var,
            font=(MONO_FONT, 13, "bold"),
            bg=INPUT_BG,
            fg="#4ec9b0",  # Soft cyan/teal highlight
            insertbackground=TEXT_WHITE,
            readonlybackground=INPUT_BG,
            relief="flat",
            justify="center",
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
            state="readonly",
        )
        self.key_entry.pack(side="left", fill="x", expand=True, ipady=6)

        self.copy_btn = tk.Button(
            key_box,
            text="📋 Copy",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9, "bold"),
            command=self._copy_key,
            cursor="hand2",
            padx=12,
        )
        self.copy_btn.pack(side="left", padx=(8, 0), ipady=4)

        # Details box
        self.details_var = tk.StringVar(value="Select a plan above and click 'Generate New Serial Key'.")
        self.details_lbl = ttk.Label(res_card, textvariable=self.details_var, style="CardText.TLabel", wraplength=480)
        self.details_lbl.pack(anchor="w", pady=(4, 0))

        self.status_var = tk.StringVar(value="")
        self.status_lbl = ttk.Label(res_card, textvariable=self.status_var, style="CardMuted.TLabel")
        self.status_lbl.pack(anchor="w", pady=(6, 0))

    def _toggle_secret(self):
        if self.secret_entry.cget("show") == "*":
            self.secret_entry.configure(show="")
            self.show_secret_btn.configure(text="🔒 Hide")
        else:
            self.secret_entry.configure(show="*")
            self.show_secret_btn.configure(text="👁 Show")

    def _on_generate(self):
        api_url = self.api_url_var.get().strip()
        secret = self.secret_var.get().strip()
        plan = self.plan_var.get().strip()

        if not api_url:
            messagebox.showerror("Missing Information", "Please enter your License Server URL.", parent=self)
            return
        if not secret:
            messagebox.showerror("Missing Information", "Please enter your Admin Secret.", parent=self)
            return

        self.gen_btn.configure(state="disabled", text="Generating Key...", bg="#1a5c28")
        self.update_idletasks()

        try:
            res = generate_key_api(api_url, secret, plan)
            key = res["serial_key"]
            self.key_var.set(key)

            # Auto-copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(key)

            # Details
            plan_label = res.get("plan_label", plan)
            days = res.get("duration_days", "")
            devices = res.get("max_devices", "")
            self.details_var.set(
                f"✅ Key Created & Copied to Clipboard!\n\n"
                f"• Plan: {plan_label}\n"
                f"• Validity: {days} Days (from first activation)\n"
                f"• Allowed Devices: {devices}\n"
                f"• Ready to paste & send to buyer."
            )
            self.status_var.set("Copied to clipboard automatically.")
            self.copy_btn.configure(text="✅ Copied!", fg="#4ec9b0")

            def reset_btn():
                self.copy_btn.configure(text="📋 Copy", fg=TEXT_MAIN)
            self.after(3000, reset_btn)

        except Exception as e:
            messagebox.showerror("Generation Error", f"Failed to generate serial key:\n\n{e}", parent=self)
            self.status_var.set(f"Error: {e}")
        finally:
            self.gen_btn.configure(state="normal", text="⚡  GENERATE NEW SERIAL KEY", bg=ACCENT_GREEN)

    def _copy_key(self):
        key = self.key_var.get().strip()
        if not key:
            return
        self.clipboard_clear()
        self.clipboard_append(key)
        self.copy_btn.configure(text="✅ Copied!", fg="#4ec9b0")
        self.status_var.set("Copied to clipboard.")
        self.after(2000, lambda: self.copy_btn.configure(text="📋 Copy", fg=TEXT_MAIN))


def cli_main():
    parser = argparse.ArgumentParser(description="Bill Scrapper Admin Serial Key Generator")
    parser.add_argument("--plan", choices=["1_month", "3_months", "6_months", "1_year"], help="Subscription plan type")
    parser.add_argument("--api", default=DEFAULT_API_URL, help="License API base URL")
    parser.add_argument("--secret", default=DEFAULT_ADMIN_SECRET, help="Admin secret token")
    parser.add_argument("--gui", action="store_true", help="Open Tkinter GUI")

    args = parser.parse_args()

    # If no plan provided or --gui flag passed, launch GUI
    if args.gui or not args.plan:
        app = KeyGeneratorGUI()
        app.mainloop()
        return

    # CLI mode
    try:
        print(f"[*] Contacting {args.api} to generate '{args.plan}' license key...")
        res = generate_key_api(args.api, args.secret, args.plan)
        key = res["serial_key"]
        print("\n========================================================")
        print(f"🎉 GENERATED SERIAL KEY: {key}")
        print(f"   Plan:      {res.get('plan_label', args.plan)}")
        print(f"   Duration:  {res.get('duration_days')} Days")
        print(f"   Devices:   Up to {res.get('max_devices')} Device(s)")
        print("========================================================\n")

        # Try copying to clipboard
        try:
            import tkinter as tk
            r = tk.Tk()
            r.withdraw()
            r.clipboard_clear()
            r.clipboard_append(key)
            r.update()
            r.destroy()
            print("📋 Key copied to clipboard automatically.")
        except Exception:
            pass

    except Exception as exc:
        print(f"❌ ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    cli_main()
