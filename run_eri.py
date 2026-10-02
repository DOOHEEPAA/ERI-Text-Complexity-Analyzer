"""ERI 텍스트 복잡도 계산기 실행 파일.

더블클릭하거나 `python run_eri.py`로 실행하면 창이 열린다.
명령줄 사용법은 `python run_eri.py --help` 참고.
"""
import sys

from eri.cli import main

if __name__ == "__main__":
    sys.exit(main())
