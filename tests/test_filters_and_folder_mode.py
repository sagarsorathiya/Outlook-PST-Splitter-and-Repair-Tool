import types
from pathlib import Path
from pstsplitter.splitter import _group_by_folder, group_items_by_size
from pstsplitter.outlook import MailItemInfo


def make_item(size: int, folder: str = '', received=None, sender=None):
    return MailItemInfo(entry_id=f'id-{size}-{folder}', subject='s', size=size, received=received, folder_path=folder, sender_email=sender)


def test_group_by_folder_basic():
    items = [
        make_item(10, 'Inbox'),
        make_item(5, 'Inbox/Sub'),
        make_item(2, 'Sent Items'),
        make_item(1, ''),
    ]
    groups = _group_by_folder(items)
    # Root group first, then Inbox, then Sent Items alphabetically
    assert len(groups) == 3
    top_names = []
    # Sort groups to ensure deterministic order (Root first, then alphabetical)
    sorted_groups = sorted(groups.items(), key=lambda x: (x[0] != 'Root', x[0]))
    for group_name, group_items in sorted_groups:
        fp = group_items[0].folder_path or ''
        top = fp.split('/',1)[0] if fp else 'root'
        top_names.append(top)
    assert top_names[0] == 'root'
    assert set(top_names[1:]) == {'Inbox','Sent Items'}


def test_group_items_by_size_with_oversize():
    items = [make_item(10), make_item(5000), make_item(20)]
    buckets = group_items_by_size(items, 100)
    assert len(buckets) == 3
    assert buckets[1][0].size == 5000
