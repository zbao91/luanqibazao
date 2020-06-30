# encoding: utf-8
from operator import itemgetter



class List():
    def __init__(self):
        pass

    def get_common(self, list1, list2):
        """
            ref: https://kite.com/python/answers/how-to-find-common-elements-between-two-lists-in-python
            return common from 2 list
        """
        list1_as_set = set(list1)
        return list(list1_as_set.intersection(list2))

    def get_by_multi_indices(self, _list, list_indices):
        """
        ref: https://stackoverflow.com/questions/18272160/access-multiple-elements-of-list-knowing-their-index
        return list of element by list of indices
        """
        return itemgetter(*list_indices)(_list)


