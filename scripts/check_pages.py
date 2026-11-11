# -*- coding: utf-8 -*-
"""
check_pages.py — 用 Word COM (pywin32) 计算 docx 页数（通用版）

不依赖任何写死路径，直接用当前 Python 运行。
需要: pip install pywin32 （仅 Windows + 已安装 Microsoft Word）

用法:
    python check_pages.py <docx路径>
输出:
    页数（整数） 或 ERR:错误信息
"""
import sys
import os
import time


def kill_word():
    """清理残留 Word 进程。"""
    try:
        import subprocess
        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                      capture_output=True, timeout=10)
    except:
        pass
    time.sleep(0.5)


def get_pages(path):
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        print("ERR:file not found: " + abs_path)
        return

    word = None
    doc = None

    for attempt in range(3):
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()

            if attempt > 0:
                kill_word()
                time.sleep(2)

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            doc = word.Documents.Open(abs_path)
            time.sleep(0.5)

            # 立即获取页数并打印
            pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            print(pages)

            # 清理（忽略错误）
            try:
                doc.Close(False)
            except:
                pass
            try:
                word.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return

        except ImportError:
            print("ERR:pywin32 not installed. Run: pip install pywin32")
            return
        except Exception as e:
            err_msg = str(e)
            try:
                if doc:
                    doc.Close(False)
            except:
                pass
            try:
                if word:
                    word.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

            if attempt < 2:
                kill_word()
                time.sleep(2)
                continue

            print("ERR:" + err_msg)
            return


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("ERR:usage: python check_pages.py <docx_path>")
        sys.exit(1)
    get_pages(sys.argv[1])
