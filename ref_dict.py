def redirect(ref_instance, name):

    sta_dict = {'0465': 'FWD', '0693': 'AFT', '0699': 'AFT', '0849': 'AFT',
                '1041': 'AFT', '1365': 'FWD', '1473': 'AFT', '1618': 'AFT',
                '1618+95': 'AFT', '1623': 'FWD'}

    try:
        sta_check = int(ref_instance.sta_to_find)
    except ValueError:
        pass
    else:
        if sta_check < 465 and ref_instance.bin_order == 1:
            return name + '_FWD'
        elif sta_check < 465 and ref_instance.bin_order == ref_instance.irm_ln:
            return name + '_AFT'

    try:
        new_name = name + '_' + sta_dict[ref_instance.sta_to_find]
    except KeyError:
        return name
    else:
        return new_name