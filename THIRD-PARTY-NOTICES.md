# Third-party software in Unbound

Unbound.exe contains the following software, each under its own licence. None of it has been modified.

| Component | Version | Licence | Source |
|---|---|---|---|
| Python runtime | 3.15 | Python Software Foundation License 2.0 | https://www.python.org/downloads/source/ |
| Qt for Python (PySide6, Shiboken6) and the Qt libraries they include | 6.12.0 | GNU Lesser General Public License v3.0 | https://code.qt.io/cgit/pyside/pyside-setup.git/ and https://download.qt.io/official_releases/qt/ |
| libpff and its Python bindings (pypff) | 20261030 | GNU Lesser General Public License v3.0 or later | https://github.com/libyal/libpff |
| pywin32 | 312 | Python Software Foundation License 2.0 | https://github.com/mhammond/pywin32 |
| PyInstaller bootloader | 6.22.3 | GNU General Public License v2.0 or later, with an exception that allows it to start programs under any licence | https://github.com/pyinstaller/pyinstaller |

The versions are the ones pinned in `requirements.txt` and `requirements-dev.txt`.

## The LGPL libraries

Qt for Python and libpff are used under the GNU Lesser General Public License, version 3. The licence text is in [licenses/LGPL-3.0.txt](licenses/LGPL-3.0.txt), and the GNU General Public License it builds on is in [licenses/GPL-3.0.txt](licenses/GPL-3.0.txt). In a release, both texts follow this page in `Third-party-licenses.txt`.

Unbound uses these libraries as separate, unmodified files; Unbound.exe unpacks them alongside its own code when it starts. To run Unbound with a different or modified version of one of them, install that version and rebuild the EXE as described under "Build the EXE" in [DEVELOPMENT.md](DEVELOPMENT.md).

Qt itself includes further third-party components, listed at https://doc.qt.io/qt-6/licenses-used-in-qt.html.
