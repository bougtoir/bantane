#!/usr/bin/env python3
"""管理者用ライセンス発行ツール

使い方:
    # ライセンスを発行（対話モード）
    python generate_license.py

    # コマンドライン引数で指定
    python generate_license.py --user-id admin --password secret123 --days 365
"""

import argparse
import sys
from pathlib import Path

# app.py と同じディレクトリにあることを前提
sys.path.insert(0, str(Path(__file__).parent))


def main():
    parser = argparse.ArgumentParser(
        description="ばんたね病院 シフト最適化アプリ - ライセンス発行ツール"
    )
    parser.add_argument("--user-id", help="ユーザーID")
    parser.add_argument("--password", help="パスワード")
    parser.add_argument(
        "--days",
        type=int,
        default=365,
        help="有効日数（デフォルト: 365日）",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="出力先ファイルパス（デフォルト: .license）",
    )
    args = parser.parse_args()

    # Lazy import so --help works without heavy deps
    from license_manager import LicenseManager

    # Interactive mode if required args are missing
    user_id = args.user_id
    password = args.password

    if not user_id:
        user_id = input("ユーザーID: ").strip()
        if not user_id:
            print("エラー: ユーザーIDを入力してください。", file=sys.stderr)
            sys.exit(1)

    if not password:
        import getpass

        password = getpass.getpass("パスワード: ")
        if not password:
            print("エラー: パスワードを入力してください。", file=sys.stderr)
            sys.exit(1)

    output_path = Path(args.output) if args.output else None
    manager = LicenseManager(license_file=output_path)

    path = manager.generate_license(
        user_id=user_id,
        password=password,
        expiration_days=args.days,
    )

    print()
    print("=" * 50)
    print("  ライセンス発行完了")
    print("=" * 50)
    print(f"  ファイル     : {path}")
    print(f"  ユーザーID   : {user_id}")
    print(f"  有効日数     : {args.days}日")
    print("=" * 50)
    print()
    print("配布手順:")
    print("  1. 上記の .license ファイルを対象PCへコピー")
    print("  2. アプリの実行ファイル（.exe）と同じフォルダに配置")
    print("  3. アプリ起動時にユーザーID / パスワードを入力")


if __name__ == "__main__":
    main()
