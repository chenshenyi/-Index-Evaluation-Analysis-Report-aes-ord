from aes_tool.rank import ranking_list


def test_ranking_list():
    sortedlist = [1, 2, 2, 3, 4, 5]
    assert ranking_list(sortedlist, 'average') == (
        1, 2.5, 2.5, 3, 5, 6), ranking_list(sortedlist, 'average')
    assert ranking_list(sortedlist, 'min') == (
        1, 2, 2, 4, 5, 6), ranking_list(sortedlist, 'min')
    assert ranking_list(sortedlist, 'max') == (
        1, 3, 3, 4, 5, 6), ranking_list(sortedlist, 'max')
    assert ranking_list(sortedlist) == (
        1, 2, 2, 4, 5, 6), ranking_list(sortedlist)
