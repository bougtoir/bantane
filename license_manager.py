"""License management module for Bantane Shift Optimizer.

Provides hardware-bound license generation and validation using
Fernet (AES-128-CBC + HMAC-SHA256) encryption.

This module is deliberately kept free of GUI dependencies so that
``generate_license.py`` can import it without PySide6.
"""

import datetime as dt
import hashlib
import logging
import platform
from pathlib import Path
from typing import List, Optional, Tuple


def _get_app_dir() -> Path:
    """Return the directory that contains the running executable / script."""
    import sys
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


class LicenseManager:
    """Manages license validation with hardware-bound encryption.

    License files are encrypted with Fernet (AES-128-CBC + HMAC-SHA256)
    using a key derived from a master secret.  Each license is optionally
    bound to a machine fingerprint so copying the .license file to another
    PC will fail validation.
    """

    _MASTER = b'BantaneShiftOpt2026!SecureMasterKey'

    def __init__(self, license_file: Optional[Path] = None):
        if license_file is None:
            self.license_file = self._find_license_file()
        else:
            self.license_file = Path(license_file)

    @staticmethod
    def _find_license_file() -> Path:
        """Search for .license file in app dir and its subdirectories."""
        app_dir = _get_app_dir()
        # 1) app dir itself
        candidate = app_dir / '.license'
        if candidate.exists():
            return candidate
        # 2) files/ subfolder
        candidate = app_dir / 'files' / '.license'
        if candidate.exists():
            return candidate
        # Default to app dir (will show "not found" on validation)
        return app_dir / '.license'

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

    def generate_license(
        self,
        user_id: str,
        password: str,
        expiration_days: int = 365,
        machine_fingerprint: Optional[str] = None,
    ) -> str:
        """Create and save a new license file.  Returns the file path.

        If *machine_fingerprint* is ``None`` the current machine's fingerprint
        is used.  Pass an explicit fingerprint to generate a license for a
        remote machine (see ``generate_license.py``).
        """
        import json
        import os
        import secrets as _secrets

        salt_hex = _secrets.token_hex(16)
        password_hash = self._hash_password(password, salt_hex)
        expiration_date = (
            dt.datetime.now() + dt.timedelta(days=expiration_days)
        ).strftime('%Y-%m-%d')

        fp = machine_fingerprint or self.get_machine_fingerprint()

        license_data = {
            'version': 2,
            'user_id': user_id,
            'password_hash': password_hash,
            'salt': salt_hex,
            'expiration_date': expiration_date,
            'machine_fingerprint': fp,
            'created_at': dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

        enc_salt = os.urandom(16)
        token = self._encrypt(json.dumps(license_data), enc_salt)

        # Fixed-length format: exactly 16 salt bytes followed by token.
        # No delimiter needed — salt length is known at read time.
        with open(self.license_file, 'wb') as f:
            f.write(enc_salt + token)

        return str(self.license_file)

    def validate_license(self, user_id: str, password: str) -> Tuple[bool, str]:
        """Validate credentials and machine binding.  Returns (ok, message)."""
        import json

        if not self.license_file.exists():
            return False, '\u30e9\u30a4\u30bb\u30f3\u30b9\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002'

        try:
            raw = self.license_file.read_bytes()

            # v2 format: first 16 bytes = encryption salt, rest = Fernet token
            if len(raw) <= 16:
                return False, '\u30e9\u30a4\u30bb\u30f3\u30b9\u30d5\u30a1\u30a4\u30eb\u306e\u5f62\u5f0f\u304c\u4e0d\u6b63\u3067\u3059\u3002'
            enc_salt = raw[:16]
            token = raw[16:]
            # Strip legacy separator if present (v2 files written before fix)
            if token[:1] == b'\n':
                token = token[1:]
            decrypted = self._decrypt(token, enc_salt)
            license_data = json.loads(decrypted)

            if license_data.get('user_id') != user_id:
                return False, '\u30e6\u30fc\u30b6\u30fcID\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002'

            pw_hash = self._hash_password(password, license_data['salt'])
            if pw_hash != license_data['password_hash']:
                return False, '\u30d1\u30b9\u30ef\u30fc\u30c9\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002'

            # Machine binding check
            expected_fp = license_data.get('machine_fingerprint')
            if expected_fp:
                current_fp = self.get_machine_fingerprint()
                if current_fp != expected_fp:
                    return False, (
                        '\u3053\u306ePC\u306e\u30e9\u30a4\u30bb\u30f3\u30b9\u3067\u306f\u3042\u308a\u307e\u305b\u3093\u3002\n'
                        '\u7ba1\u7406\u8005\u306b\u73fe\u5728\u306e\u30de\u30b7\u30f3ID \u3092\u304a\u4f1d\u3048\u304f\u3060\u3055\u3044:\n'
                        f'{current_fp[:16]}...'
                    )

            expiration = dt.datetime.strptime(
                license_data['expiration_date'], '%Y-%m-%d'
            )
            if dt.datetime.now() > expiration:
                return False, (
                    f"\u30e9\u30a4\u30bb\u30f3\u30b9\u306e\u6709\u52b9\u671f\u9650\u304c\u5207\u308c\u3066\u3044\u307e\u3059\u3002"
                    f"\uff08\u6709\u52b9\u671f\u9650: {license_data['expiration_date']}\uff09"
                )

            days_remaining = (expiration - dt.datetime.now()).days
            return True, f'\u30e9\u30a4\u30bb\u30f3\u30b9\u8a8d\u8a3c\u6210\u529f\u3002\u6b8b\u308a{days_remaining}\u65e5\u6709\u52b9\u3067\u3059\u3002'

        except Exception as e:
            logging.error(f'License validation error: {e}')
            return False, f'\u30e9\u30a4\u30bb\u30f3\u30b9\u30d5\u30a1\u30a4\u30eb\u306e\u8aad\u307f\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {e}'
