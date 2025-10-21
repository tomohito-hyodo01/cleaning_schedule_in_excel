"""
メインエントリーポイント

CLIまたはGUIモードで起動
"""

import sys
import os
import argparse

# スクリプトのディレクトリをPythonパスに追加
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description="Excel画像変換ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # GUIモードで起動
  python main.py
  python main.py --gui
  
  # CLIモードで実行
  python main.py --cli input.png -o output.xlsx
        """
    )
    
    parser.add_argument(
        "--gui",
        action="store_true",
        help="GUIモードで起動（デフォルト）"
    )
    
    parser.add_argument(
        "--cli",
        action="store_true",
        help="CLIモードで起動"
    )
    
    # CLIモード用の引数（--cliが指定された場合のみ有効）
    parser.add_argument(
        "input",
        nargs="?",
        help="入力画像ファイル（CLIモード時）"
    )
    
    parser.add_argument(
        "-o", "--output",
        help="出力Excelファイル（CLIモード時）"
    )
    
    args, remaining = parser.parse_known_args()
    
    # モードの決定
    if args.cli or args.input:
        # CLIモード
        from src.cli import main as cli_main
        
        # 引数を再構築してCLIに渡す
        cli_args = []
        if args.input:
            cli_args.append(args.input)
        if args.output:
            cli_args.extend(["-o", args.output])
        cli_args.extend(remaining)
        
        sys.argv = [sys.argv[0]] + cli_args
        return cli_main()
    else:
        # GUIモード（デフォルト）
        from src.ui import main as gui_main
        return gui_main()


if __name__ == "__main__":
    sys.exit(main())

