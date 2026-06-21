"""License management module for Bantane Shift Optimizer.

Provides license generation and validation using
Fernet (AES-128-CBC + HMAC-SHA256) encryption.

This module is deliberately kept free of GUI dependencies so that
``generate_license.py`` can import it without PySide6.
"""

import calendar
import datetime as dt
import hashlib
import logging
import platform
import secrets as _secrets
import string
from pathlib import Path
from typing import List, Optional, Tuple


def _get_app_dir() -> Path:
    """Return the directory that contains the running executable / script."""
    import sys
    # PyInstaller
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Nuitka / other bundlers: sys.executable is NOT a python interpreter
    try:
        exe_stem = Path(sys.executable).stem.lower()
        if exe_stem not in ("python", "python3", "pythonw", "python3w") \
                and not exe_stem.startswith("python3."):
            return Path(sys.executable).resolve().parent
    except Exception:
        pass
    return Path(__file__).resolve().parent


class LicenseManager:
    """Manages license generation and validation.

    License files are encrypted with Fernet (AES-128-CBC + HMAC-SHA256)
    using a key derived from a master secret.
    """

    _MASTER = b'BantaneShiftOpt2026!SecureMasterKey'

    def __init__(self, license_file: Optional[Path] = None):
        if license_file is None:
            self.license_file = self._find_license_file()
        else:
            self.license_file = Path(license_file)

    @staticmethod
    def _find_license_file() -> Path:
        """Search for .license file in app dir/files/ first, then app dir."""
        app_dir = _get_app_dir()
        # 1) files/ subfolder (default location)
        candidate = app_dir / 'files' / '.license'
        if candidate.exists():
            return candidate
        # 2) app dir itself
        candidate = app_dir / '.license'
        if candidate.exists():
            return candidate
        # 3) any immediate subdirectory
        for sub in app_dir.iterdir():
            if sub.is_dir():
                candidate = sub / '.license'
                if candidate.exists():
                    return candidate
        # Default to files/ subfolder
        return app_dir / 'files' / '.license'

    # -- cryptography helpers ------------------------------------------------

    @staticmethod
    def _derive_key(master: bytes, salt: bytes) -> bytes:
        """Derive a Fernet-compatible key from *master* and *salt*."""
        import base64
        dk = hashlib.pbkdf2_hmac(
            'sha256', master, salt, iterations=200_000, dklen=32
        )
        return base64.urlsafe_b64encode(dk)

    def _encrypt(self, data: str, salt: bytes) -> bytes:
        from cryptography.fernet import Fernet
        key = self._derive_key(self._MASTER, salt)
        return Fernet(key).encrypt(data.encode('utf-8'))

    def _decrypt(self, token: bytes, salt: bytes) -> str:
        from cryptography.fernet import Fernet
        key = self._derive_key(self._MASTER, salt)
        return Fernet(key).decrypt(token).decode('utf-8')

    @staticmethod
    def _hash_password(password: str, salt: str) -> str:
        return hashlib.sha256((password + salt).encode('utf-8')).hexdigest()

    # -- hardware fingerprint ------------------------------------------------

    @staticmethod
    def get_machine_fingerprint() -> str:
        """Return a stable fingerprint for the current machine.

        Combines hostname, all physical MAC addresses (sorted), and
        (on Windows) the system drive volume serial number.
        The result is a hex SHA-256 digest.

        Note: Uses all MAC addresses sorted to ensure deterministic results
        regardless of NIC enumeration order (uuid.getnode() can be
        non-deterministic when multiple NICs exist).
        """
        import socket
        parts: List[str] = []
        parts.append(socket.gethostname())

        # Collect all MAC addresses deterministically
        mac_addresses: List[str] = []
        if platform.system() == 'Windows':
            try:
                import subprocess as _sp
                out = _sp.check_output(
                    'getmac /FO CSV /NH',
                    shell=True, text=True, stderr=_sp.DEVNULL
                )
                for line in out.strip().splitlines():
                    # Each line: "AA-BB-CC-DD-EE-FF","...","..."
                    cols = line.split(',')
                    if cols:
                        mac = cols[0].strip().strip('"')
                        if mac and mac != 'N/A' and '-' in mac:
                            mac_addresses.append(mac.upper())
            except Exception:
                pass
        else:
            try:
                import subprocess as _sp
                out = _sp.check_output(
                    ['ip', 'link', 'show'],
                    text=True, stderr=_sp.DEVNULL
                )
                import re
                for m in re.finditer(r'link/ether\s+([0-9a-fA-F:]{17})', out):
                    mac_addresses.append(m.group(1).upper())
            except Exception:
                pass

        if not mac_addresses:
            # Fallback to uuid.getnode() if no MACs found
            import uuid as _uuid
            mac_addresses.append(str(_uuid.getnode()))

        mac_addresses.sort()
        parts.append('|'.join(mac_addresses))

        if platform.system() == 'Windows':
            try:
                import subprocess as _sp
                out = _sp.check_output(
                    'vol C:', shell=True, text=True, stderr=_sp.DEVNULL
                )
                for line in out.splitlines():
                    if 'Serial' in line or '\u30b7\u30ea\u30a2\u30eb' in line:
                        parts.append(line.strip())
                        break
            except Exception:
                pass
        raw = '|'.join(parts)
        return hashlib.sha256(raw.encode('utf-8')).hexdigest()

    # -- public API ----------------------------------------------------------

    @staticmethod
    def _generate_password(length: int = 16) -> str:
        """Generate a random password of adequate strength."""
        alphabet = string.ascii_letters + string.digits + string.punctuation
        # Ensure at least one of each category
        pw: List[str] = [
            _secrets.choice(string.ascii_uppercase),
            _secrets.choice(string.ascii_lowercase),
            _secrets.choice(string.digits),
            _secrets.choice(string.punctuation),
        ]
        pw += [_secrets.choice(alphabet) for _ in range(length - 4)]
        # Shuffle so guaranteed chars aren't always at the start
        import random
        rng = random.SystemRandom()
        rng.shuffle(pw)
        return ''.join(pw)

    @staticmethod
    def parse_expiration(value: str) -> dt.date:
        """Parse expiration input.

        Accepts:
          - ``yyyymm`` — valid until the last day of that month
          - ``dd``     — valid for *dd* days from tomorrow (creation day
                         excluded)
        """
        value = value.strip()
        if len(value) == 6 and value.isdigit():
            year = int(value[:4])
            month = int(value[4:])
            if month < 1 or month > 12:
                raise ValueError(f"月が不正です: {month}")
            last_day = calendar.monthrange(year, month)[1]
            return dt.date(year, month, last_day)
        elif value.isdigit():
            days = int(value)
            if days <= 0:
                raise ValueError("日数は1以上を指定してください。")
            # Creation day excluded: start counting from tomorrow
            return dt.date.today() + dt.timedelta(days=days)
        else:
            raise ValueError(
                "有効期間は yyyymm（例: 202612）または日数（例: 90）で"
                "指定してください。"
            )

    def generate_license(
        self,
        user_id: str,
        password: Optional[str] = None,
        expiration_days: Optional[int] = None,
        expiration_date: Optional[dt.date] = None,
        machine_fingerprint: Optional[str] = None,
    ) -> Tuple[str, str]:
        """Create and save a new license file.

        Returns ``(file_path, generated_password)``.

        If *password* is ``None`` a strong random password is generated.
        Exactly one of *expiration_days* or *expiration_date* should be set.
        """
        import json
        import os

        if password is None:
            password = self._generate_password()

        if expiration_date is not None:
            exp_str = expiration_date.strftime('%Y-%m-%d')
        elif expiration_days is not None:
            exp_str = (
                dt.datetime.now() + dt.timedelta(days=expiration_days)
            ).strftime('%Y-%m-%d')
        else:
            exp_str = (
                dt.datetime.now() + dt.timedelta(days=365)
            ).strftime('%Y-%m-%d')

        salt_hex = _secrets.token_hex(16)
        password_hash = self._hash_password(password, salt_hex)

        license_data = {
            'version': 4,
            'user_id': user_id,
            'password_hash': password_hash,
            'password': password,
            'salt': salt_hex,
            'expiration_date': exp_str,
            'created_at': dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

        enc_salt = os.urandom(16)
        token = self._encrypt(json.dumps(license_data), enc_salt)

        with open(self.license_file, 'wb') as f:
            f.write(enc_salt + token)

        return str(self.license_file), password

    def _read_license_data(self) -> dict:
        """Decrypt and return the license data dict.  Raises on failure."""
        import json

        raw = self.license_file.read_bytes()
        if len(raw) <= 16:
            raise ValueError('ライセンスファイルの形式が不正です。')
        enc_salt = raw[:16]
        token = raw[16:]
        if token[:1] == b'\n':
            token = token[1:]
        decrypted = self._decrypt(token, enc_salt)
        return json.loads(decrypted)

    def validate_license_auto(self) -> Tuple[bool, str]:
        """Auto-validate: file exists, decryptable, not expired.

        No user_id / password required.  Returns (ok, message).
        """
        if not self.license_file.exists():
            return False, 'ライセンスファイルが見つかりません。'
        try:
            data = self._read_license_data()
            expiration = dt.datetime.strptime(
                data['expiration_date'], '%Y-%m-%d'
            )
            if dt.datetime.now() > expiration:
                return False, (
                    f"ライセンスの有効期限が切れています。"
                    f"（有効期限: {data['expiration_date']}）"
                )
            days_remaining = (expiration - dt.datetime.now()).days
            user_id = data.get('user_id', '不明')
            return True, (
                f'ライセンス認証成功（{user_id}）。'
                f'残り{days_remaining}日有効です。'
            )
        except Exception as e:
            logging.error(f'License auto-validation error: {e}')
            return False, f'ライセンスファイルの読み込みに失敗しました: {e}'

    def validate_license(self, user_id: str, password: str) -> Tuple[bool, str]:
        """Validate credentials.  Returns (ok, message)."""
        if not self.license_file.exists():
            return False, 'ライセンスファイルが見つかりません。'

        try:
            data = self._read_license_data()

            if data.get('user_id') != user_id:
                return False, 'ユーザーIDが正しくありません。'

            pw_hash = self._hash_password(password, data['salt'])
            if pw_hash != data['password_hash']:
                return False, 'パスワードが正しくありません。'

            expiration = dt.datetime.strptime(
                data['expiration_date'], '%Y-%m-%d'
            )
            if dt.datetime.now() > expiration:
                return False, (
                    f"ライセンスの有効期限が切れています。"
                    f"（有効期限: {data['expiration_date']}）"
                )

            days_remaining = (expiration - dt.datetime.now()).days
            return True, f'ライセンス認証成功。残り{days_remaining}日有効です。'

        except Exception as e:
            logging.error(f'License validation error: {e}')
            return False, f'ライセンスファイルの読み込みに失敗しました: {e}'
