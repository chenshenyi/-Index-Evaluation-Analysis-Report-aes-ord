from aes_tool.rank import ranking_list


def test_ranking_list():
    sortedlist = [1, 2, 2, 3, 4, 5]
    assert ranking_list(sortedlist, 'dense') == (
        1, 2, 2, 3, 4, 5), ranking_list(sortedlist, 'dense')
    assert ranking_list(sortedlist, 'simple') == (
        1, 2, 2, 4, 5, 6), ranking_list(sortedlist, 'simple')
