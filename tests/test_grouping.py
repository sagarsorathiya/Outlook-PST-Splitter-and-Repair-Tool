import random
from pstsplitter.util import chunk_sequence


def test_chunk_sequence_basic():
    items = [10, 20, 30, 40]
    chunks = chunk_sequence(items, 50)
    assert chunks == [[10, 20], [30], [40]]


def test_chunk_sequence_oversize_item():
    items = [10, 200, 20]
    chunks = chunk_sequence(items, 100)
    assert chunks == [[10], [200], [20]]


def test_chunk_sequence_random_stability():
    random.seed(1)
    data = [random.randint(1, 50) for _ in range(30)]
    chunks = chunk_sequence(data, 100)
    # validate each chunk sum
    for ch in chunks:
        if len(ch) == 1 and ch[0] > 100:
            continue
        assert sum(ch) <= 100
