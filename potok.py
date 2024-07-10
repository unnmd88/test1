# def make_condition(det_ranges, num_group=None):
#
#     det_collection = det_ranges.replace(' ', '').split(',')
#
#     condition_string = ''
#     for det_range in det_collection:
#         det_from, det_to = det_range.split('-')
#         if det_from.isdigit() and det_to.isdigit():
#             det_from, det_to = int(det_from), int(det_to)
#             for num_det in range(det_from, det_to + 1):
#                 if num_det != det_to:
#                     condition_string += f'ddr(D{num_det}) or '
#                 else:
#                     condition_string += f'ddr(D{num_det})'
#     print(condition_string)
#     # det_from, det_to

def make_condition(det_ranges: str, num_group: str = None):
    """ Функция формирует строку продления/запроса ДК Поток стандартного вида:
        ddr(D1) or ddr(D2) or .... ddr(Dn) или (ddr(D1) or ddr(D2) or .... ddr(Dn)) and mr(Gn)
        :param det_ranges -> диапазоны номеров детекторов вида 1-4, 5-12 и т.д
        :param num_group -> номер группы для функции mr
        :return -> строка продления/запроса стандартного вида
    """

    all_det_ranges = []

    det_collection = det_ranges.replace(' ', '').split(',')

    for num, det_range in enumerate(det_collection, 1):
        det_from, det_to = det_range.split('-')
        condition_string = ''

        if not det_from.isdigit() or not det_to.isdigit():
            return # Вернуть какое нибудь сообщение пользователю о некорректности введённых данных

        det_from, det_to = int(det_from), int(det_to)
        for num_det in range(det_from, det_to + 1):
            if num_det != det_to:
                condition_string += f'ddr(D{num_det}) or '
            else:
                condition_string += f'ddr(D{num_det})'
        all_det_ranges.append(condition_string)
    print(all_det_ranges)

    condition_string = ' or '.join(all_det_ranges)
    print(condition_string)

    if num_group is not None:
        if num_group.isdigit():
            condition_string = f'({condition_string}) and mr(G{num_group})'
        else:
            return # Вернуть какое нибудь сообщение пользователю о некорректности введённых данных
    print(condition_string)
    return condition_string

    # det_from, det_to

a = make_condition('2- 4,   12-18, 34-42')
""" Test commit num 2 """
b = make_condition('3-5, 9-14', '2')
