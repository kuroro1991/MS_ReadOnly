@echo off
rem 読み取り専用ツールのパス
set tool_path=".\ReadOnly.vbs"

echo ＊＊＊＊＊　読み取り専用ツール　＊＊＊＊＊
echo
echo ・・・・・ファイルを開いています・・・・・

%tool_path% %1
