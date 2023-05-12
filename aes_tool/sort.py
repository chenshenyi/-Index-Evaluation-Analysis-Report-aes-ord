from typing import Iterable, Literal


def is_sorted(iterable: Iterable, ascending: Literal[True, False, 'both']) -> bool:
    """Return True if the iterable is sorted.

    Args:
        iterable (Iterable): The iterable.
        ascending (bool): Whether the iterable is sorted in ascending order.

    Returns:
        bool: Whether the iterable is sorted.
    """

    # test if the elements of iterable is comparable
    try:
        iterable[0] < iterable[0]
    except TypeError:
        raise TypeError('The elements of iterable is not comparable.')

    if ascending == 'both':
        return is_sorted(iterable, True) or is_sorted(iterable, False)
    elif ascending:
        return all(iterable[i] <= iterable[i+1] for i in range(len(iterable)-1))
    else:
        return all(iterable[i] >= iterable[i+1] for i in range(len(iterable)-1))
