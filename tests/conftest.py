import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


@pytest.fixture(scope="session")
def kiwi():
    from kiwipiepy import Kiwi
    return Kiwi()


@pytest.fixture(scope="session")
def mini_vocab():
    """특허 예시 지문에 필요한 단어만 담은 작은 사전 (실제 국립국어원 등급 기준)."""
    from eri.vocab import Vocabulary
    words = {
        1: "가 가다 개구리 거의 것 겨울 그 그것 그래서 깊이 깨어나다 꽁꽁 꿀 날씨 다시 다행히 땅 땅속 뛰다 뛰어다니다 "
           "뜯다 마치 만큼 말 먹다 몸 물 물가 물론 봄 생기다 속 수 시간 심장 아무것 않다 어디 어떻다 얼다 얼어붙다 "
           "오다 있다 자다 잠 잠들다 잠자다 지내다 추워지다 풀 하다 흙 괜찮다 갈아입다 가격 곳 관계 기대 기술 깨지다 "
           "내리다 대하다 되다 들다 등 때 또는 모이다 물건 미래 반대 사다 사람 상품 서비스 시장 알다 어떤 오르다 "
           "이러하다 입장 팔다",
        2: "산토끼 몸속 균형 법칙 비용 발생하다 반면",
        3: "겨울옷 끈적끈적하다 감소하다 증가하다 일반적 결론적 욕구 공급 변화 생산",
        4: "공급량 수요 요인 거래하다",
        5: "수요자 공급자 재화 즉",
    }
    entries = {}
    for g, ws in words.items():
        for w in ws.split():
            entries.setdefault(w, []).append((g, ""))
    return Vocabulary(entries, "mini")
