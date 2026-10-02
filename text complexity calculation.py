"""이전 실행 파일 이름과의 호환용. 실제 코드는 eri 패키지에 있으며 run_eri.py와 같다."""
import sys

from eri.cli import main

if __name__ == "__main__":
    sys.exit(main())
