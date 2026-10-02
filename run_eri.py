"""ERI 텍스트 복잡도 계산기 실행 파일.

Windows에서는 `ERI_실행.bat`을 더블클릭하는 것이 가장 쉽다 (파이썬·라이브러리 확인 포함).
`python run_eri.py`로 실행하면 창이 열린다. 명령줄 사용법은 `python run_eri.py --help` 참고.
"""
import sys

REQUIRED = {"kiwipiepy": "kiwipiepy", "openpyxl": "openpyxl", "numpy": "numpy"}


def check_environment() -> bool:
    """파이썬 버전과 필요한 라이브러리를 확인하고, 문제가 있으면 해결 방법을 알려 준다."""
    if sys.version_info < (3, 10):
        print(f"[오류] 파이썬 3.10 이상이 필요합니다. 현재 버전: {sys.version.split()[0]}")
        return False
    missing = []
    for module, package in REQUIRED.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(package)
    if missing:
        print("[오류] 필요한 라이브러리가 설치되어 있지 않습니다: " + ", ".join(missing))
        print("       아래 명령을 한 번 실행한 뒤 다시 시도하세요.")
        print(f'       "{sys.executable}" -m pip install -r requirements.txt')
        return False
    return True


def start() -> int:
    print(f"ERI 계산기 시작 (파이썬 {sys.version.split()[0]}: {sys.executable})")
    if not check_environment():
        return 1
    from eri.cli import main
    return main()


if __name__ == "__main__":
    sys.exit(start())
