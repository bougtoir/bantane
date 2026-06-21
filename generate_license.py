#!/usr/bin/env python3
"""管理者用ライセンス発行ツール

使い方:
    # 対話モード（推奨）
    python generate_license.py

    # コマンドライン引数で指定
    python generate_license.py --user-id admin --expiration 202612
    python generate_license.py --user-id admin --expiration 90

有効期間の指定方法:
    yyyymm  → その年月の末日まで有効（例: 202612 → 2026年12月31日まで）
    dd      → 作成日を含まず、dd日後まで有効（例: 90 → 90日後まで）
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))


def main():
    parser = argparse.ArgumentParser(
        description="ばんたね病院 シフト最適化アプリ - ライセンス発行ツール"
    )
    parser.add_argument("--user-id", help="ユーザーID")
    parser.add_argument(
        "--expiration",
        default=None,
        help="有効期間: yyyymm（年月末日まで）または日数（dd）",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="出力先ファイルパス（デフォルト: files/.license）",
    )
    args = parser.parse_args()

    from license_manager import LicenseManager

    # --- User ID ---
    user_id = args.user_id
    if not user_id:
        user_id = input("ユーザーID: ").strip()
        if not user_id:
            print("エラー: ユーザーIDを入力してください。", file=sys.stderr)
            sys.exit(1)

    # --- Expiration ---
    expiration_input = args.expiration
    if not expiration_input:
        print()
        print("有効期間を指定してください:")
        print("  yyyymm  → その年月の末日まで有効（例: 202612）")
        print("  dd      → 作成日を含まず、dd日後まで有効（例: 90）")
        expiration_input = input("有効期間: ").strip()
        if not expiration_input:
            print("エラー: 有効期間を入力してください。", file=sys.stderr)
            sys.exit(1)

    try:
        exp_date = LicenseManager.parse_expiration(expiration_input)
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)

    # --- Generate ---
    output_path = Path(args.output) if args.output else None
    manager = LicenseManager(license_file=output_path)

    path, password = manager.generate_license(
        user_id=user_id,
        expiration_date=exp_date,
    )

    print()
    print("=" * 55)
    print("  ライセンス発行完了")
    print("=" * 55)
    print(f"  ファイル       : {path}")
    print(f"  ユーザーID     : {user_id}")
    print(f"  有効期限       : {exp_date.strftime('%Y年%m月%d日')}")
    print(f"  パスワード     : {password}")
    print("=" * 55)
    print()
    print("配布手順:")
    print("  1. 上記の .license ファイルを対象PCへコピー")
    print("  2. アプリの実行ファイル（.exe）と同じディレクトリの")
    print("     files フォルダに配置")
    print("  3. アプリ起動時に自動認証されます")
    print()
    print("※ パスワードは .license ファイル内に暗号化保存されています。")
    print("   ID/パスワード入力なしで起動できます。")


if __name__ == "__main__":
    main()
