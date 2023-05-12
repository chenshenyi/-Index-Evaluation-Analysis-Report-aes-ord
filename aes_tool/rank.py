from typing import Iterable, Optional, Literal, overload, Tuple


MODE = Literal['dense', 'simple']


def ranking_list(sortedlist: Iterable, operation_when_duplicate: Optional[MODE] = 'simple') -> Tuple[int]:
    """Return the ranking list of the sorted list.

    Args:
        sortedlist (Iterable): The sorted list.
        operation_when_duplicate (MODE, optional): The operation when duplicate. Defaults to 'min'.

    Returns:
        Tuple[float, int]: The ranking list.
    """

    assert is_sorted(sortedlist, 'both'), 'The list is not sorted.'

    if operation_when_duplicate == 'dense':
        return _ranking_list_dense(sortedlist)
    elif operation_when_duplicate == 'simple':
        return _ranking_list_simple(sortedlist)
    else:
        raise ValueError(
            f'Invalid operation_when_duplicate: {operation_when_duplicate}')


def _ranking_list_dense(sortedlist: Iterable) -> Tuple[int]:
    """Return the ranking list of the sorted list.

    Args:
        sortedlist (Iterable): The sorted list.

    Returns:
        Tuple[float, int]: The ranking list.
    """

    rankinglist = []
    for i in range(len(sortedlist)):
        if i == 0:
            rankinglist.append(1)
        elif sortedlist[i] == sortedlist[i-1]:
            rankinglist.append(rankinglist[i-1])
        else:
            rankinglist.append(rankinglist[i-1]+1)
    return tuple(rankinglist)


def _ranking_list_simple(sortedlist: Iterable) -> Tuple[int]:
    """Return the ranking list of the sorted list.

    Args:
        sortedlist (Iterable): The sorted list.

    Returns:
        Tuple[float, int]: The ranking list.
    """

    rankinglist = []
    for i in range(len(sortedlist)):
        if i == 0:
            rankinglist.append(1)
        elif sortedlist[i] == sortedlist[i-1]:
            rankinglist.append(rankinglist[i-1])
        else:
            rankinglist.append(i+1)
    return tuple(rankinglist)


def is_sorted(iterable: Iterable, descending: Literal[True, False, 'both']) -> bool:
    """Return True if the iterable is sorted.

    Args:
        iterable (Iterable): The iterable.
        descending (bool): Whether the iterable is sorted in descending order.

    Returns:
        bool: Whether the iterable is sorted.
    """

    # test if the elements of iterable is comparable
    try:
        iterable[0] < iterable[0]
    except TypeError:
        raise TypeError('The elements of iterable is not comparable.')

    if descending == 'both':
        return is_sorted(iterable, True) or is_sorted(iterable, False)
    elif descending:
        return all(iterable[i] >= iterable[i+1] for i in range(len(iterable)-1))
    else:
        return all(iterable[i] <= iterable[i+1] for i in range(len(iterable)-1))
