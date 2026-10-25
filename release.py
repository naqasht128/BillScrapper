#!/usr/bin/env python3
"""Bill Scrapper - Automated Release & Update Publisher.

One-command pipeline to:
1. Bump APP_VERSION in main.py
2. Stage and commit all pending code changes
3. Push commit to GitHub (main branch)
4. Create and push release tag (vX.Y.Z)
5. Automatically trigger GitHub Actions to build BillScrapper.exe and publish GitHub Release
6. Broadcast update notification to all installed desktop app users!

Usage:
    python release.py 1.2.1
    python release.py 1.2.1 -m "Fixed PESCO bill parsing bug and updated UI"
    python release.py  (Interactive mode)
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from pathlib import Path

# Fix Windows cp1252 console encoding for emojis and Unicode characters
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass



ROOT_DIR = Path(__file__).resolve().parent
MAIN_PY = ROOT_DIR / "main.py"


def run_cmd(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    """Run a system command and return output."""
    print(f"  → Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(ROOT_DIR), text=True, capture_output=True)
    if check and result.returncode != 0:
        print(f"\n❌ Error executing command: {' '.join(cmd)}")
        if result.stderr:
            print(f"Error details:\n{result.stderr.strip()}")
        if result.stdout:
            print(f"Output:\n{result.stdout.strip()}")
        sys.exit(result.returncode)
    return result


def get_current_version() -> str:
    """Read APP_VERSION from main.py."""
    if not MAIN_PY.exists():
        print(f"❌ Error: {MAIN_PY} not found.")
        sys.exit(1)

    content = MAIN_PY.read_text(encoding="utf-8")
    match = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', content)
    if not match:
        print("❌ Error: Could not find APP_VERSION in main.py")
        sys.exit(1)
    return match.group(1).strip()


def suggest_next_version(current: str) -> str:
    """Suggest next patch version (e.g. 1.2.0 -> 1.2.1)."""
    parts = current.split(".")
    try:
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    except Exception:
        return f"{current}.1"


def update_main_version(new_version: str):
    """Write new APP_VERSION into main.py."""
    content = MAIN_PY.read_text(encoding="utf-8")
    new_content, count = re.subn(
        r'APP_VERSION\s*=\s*["\'][^"\']+["\']',
        f'APP_VERSION = "{new_version}"',
        content
    )
    if count == 0:
        print("❌ Error: Failed to replace APP_VERSION in main.py")
        sys.exit(1)
    MAIN_PY.write_text(new_content, encoding="utf-8")
    print(f"  ✔ Updated main.py: APP_VERSION = \"{new_version}\"")


def update_installer_version(new_version: str):
    """Write new AppVersion into installer.iss if it exists."""
    installer_iss = ROOT_DIR / "installer.iss"
    if not installer_iss.exists():
        return
    content = installer_iss.read_text(encoding="utf-8")
    new_content, count = re.subn(
        r'#define\s+AppVersion\s+["\'][^"\']+["\']',
        f'#define AppVersion "{new_version}"',
        content
    )
    if count > 0:
        installer_iss.write_text(new_content, encoding="utf-8")
        print(f"  ✔ Updated installer.iss: AppVersion = \"{new_version}\"")


def generate_release_notes_file(new_version: str, msg: str):
    """Write clean RELEASE_NOTES.md so GitHub Releases does not insert compare links."""
    rn_file = ROOT_DIR / "RELEASE_NOTES.md"
    if msg and msg != f"Release v{new_version}":
        content = f"• {msg}"
    else:
        content = f"• Release v{new_version}: Performance improvements and stability updates"
    rn_file.write_text(content, encoding="utf-8")
    print("  ✔ Generated clean RELEASE_NOTES.md")


def main():
    parser = argparse.ArgumentParser(description="Bill Scrapper - Automated Release & Update Publisher")
    parser.add_argument("version", nargs="?", help="New version number (e.g. 1.2.1)")
    parser.add_argument("-m", "--message", default="", help="Release note / commit message")
    parser.add_argument("--skip-tests", action="store_true", help="Skip Python syntax pre-compilation check")
    args = parser.parse_args()

    print("\n" + "=" * 60)
    print("🚀  BILL SCRAPPER — AUTOMATED RELEASE PUBLISHER")
    print("=" * 60)

    current_ver = get_current_version()
    print(f"Current version in main.py: v{current_ver}")

    # Determine target version
    target_ver = args.version
    if not target_ver:
        suggested = suggest_next_version(current_ver)
        user_input = input(f"Enter new version number [default: {suggested}]: ").strip()
        target_ver = user_input if user_input else suggested

    # Clean version string (e.g. strip leading 'v')
    target_ver = target_ver.lstrip("v").strip()
    tag_name = f"v{target_ver}"

    if not re.match(r"^\d+\.\d+(\.\d+)?", target_ver):
        print(f"❌ Error: Invalid version format '{target_ver}'. Must be like 1.2.1 or 1.3")
        sys.exit(1)

    print(f"Target release version:    {tag_name}")

    # Release note
    msg = args.message.strip()
    if not msg:
        user_msg = input("Enter release notes / changes summary (optional): ").strip()
        msg = user_msg if user_msg else f"Release {tag_name}"

    # 1. Syntax check
    if not args.skip_tests:
        print("\n[Step 1/5] Verifying Python code syntax...")
        py_res = run_cmd([sys.executable, "-m", "py_compile", str(MAIN_PY)], check=False)
        if py_res.returncode != 0:
            print("❌ Syntax error detected in main.py! Aborting release.")
            sys.exit(1)
        print("  ✔ Python syntax check passed.")

    # 2. Update APP_VERSION in main.py and installer.iss
    print(f"\n[Step 2/5] Updating version in main.py and installer.iss to {target_ver}...")
    update_main_version(target_ver)
    update_installer_version(target_ver)
    generate_release_notes_file(target_ver, msg)

    # 3. Git Stage & Commit
    print("\n[Step 3/5] Staging code changes & committing...")
    run_cmd(["git", "add", "."])
    commit_msg = f"Release {tag_name}: {msg}"
    # Commit might return 1 if there's nothing new except version update
    run_cmd(["git", "commit", "-m", commit_msg], check=False)
    print("  ✔ Git commit created.")

    # 4. Push to main branch
    print("\n[Step 4/5] Pushing changes to GitHub main branch...")
    run_cmd(["git", "push", "origin", "main"])
    print("  ✔ Pushed to origin/main.")

    # 5. Create & Push Git Tag
    print(f"\n[Step 5/5] Creating and pushing tag '{tag_name}'...")
    # Delete local tag if it already exists from a failed attempt
    run_cmd(["git", "tag", "-d", tag_name], check=False)
    run_cmd(["git", "tag", "-a", tag_name, "-m", commit_msg])
    run_cmd(["git", "push", "origin", tag_name, "--force"])
    print(f"  ✔ Tag '{tag_name}' pushed to GitHub.")

    # 6. Optional: Register release on Vercel backend
    api_url = os.environ.get("LICENSE_API_URL", "https://bill-scrapper.vercel.app").rstrip("/")
    admin_secret = os.environ.get("ADMIN_SECRET", "")
    if admin_secret:
        try:
            import requests
            set_url = f"{api_url}/api/admin/set-update"
            r = requests.post(
                set_url,
                json={"version": target_ver, "release_notes": msg},
                headers={"x-admin-secret": admin_secret},
                timeout=5
            )
            if r.status_code == 200:
                print(f"  ✔ Update registered with Vercel API: v{target_ver}")
        except Exception:
            pass

    print("\n" + "=" * 60)
    print("🎉  RELEASE PIPELINE TRIGGERED SUCCESSFULLY!")
    print("=" * 60)
    print(f"\nWhat happens next:")
    print(f"1. GitHub Actions has started building Windows (Win 10/11 & Win 7) and Linux installers.")
    print(f"2. Upon completion, GitHub Actions will publish Release {tag_name} with:")
    print(f"   • BillScrapper-Setup.exe       (Windows 10 & 11 Setup Installer)")
    print(f"   • BillScrapper-Win7-Setup.exe  (Windows 7 & 8 Setup Installer + VC++ Runtime)")
    print(f"   • BillScrapper                 (Linux ARM64 Binary)")
    print(f"3. All users who open their installed Bill Scrapper app will automatically receive the update banner:")
    print(f"   '🚀 Update Available: Version {tag_name} is ready to install!'")
    print(f"4. Users click 'Update Now' to automatically download and upgrade in-place silently.\n")
    print("💡 Tip: To verify your build progress, visit:")
    print("   https://github.com/naqasht128/BillScrapper/actions\n")


if __name__ == "__main__":
    main()
