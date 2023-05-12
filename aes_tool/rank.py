from .sort import is_sorted
from typing import Iterable, Optional, Literal, overload, Tuple


MODE = Literal['average', 'min', 'max', 'simple']


@overload
def ranking_list(
    sortedlist, operation_when_duplicate: Literal['min', 'max', 'simple']) -> Tuple[int]: ...


@overload
def ranking_list(
    sortedlist, operation_when_duplicate: Literal['average']) -> Tuple[float, int]: ...


@overload
def ranking_list(
    sortedlist, operation_when_duplicate: MODE) -> Tuple[float, int]: ...


def ranking_list(sortedlist: Iterable, operation_when_duplicate: Optional[MODE] = 'min') -> Tuple[float, int]:
    """Return the ranking list of the sorted list.

    Args:
        sortedlist (Iterable): The sorted list.
        operation_when_duplicate (MODE, optional): The operation when duplicate. Defaults to 'min'.

    Returns:
        Tuple[float, int]: The ranking list.
    """

    assert is_sorted(sortedlist, 'both'), 'The list is not sorted.'

    if operation_when_duplicate == 'average':
        return _ranking_list_average(sortedlist)
    elif operation_when_duplicate == 'min':
        return _ranking_list_min(sortedlist)
    elif operation_when_duplicate == 'max':
        return _ranking_list_max(sortedlist)
    elif operation_when_duplicate == 'simple':
        return _ranking_list_simple(sortedlist)
    else:
        raise ValueError(
            f'Invalid operation_when_duplicate: {operation_when_duplicate}')
