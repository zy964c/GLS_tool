def redirect(ref_instance, name):

    sta_dict = {'LH': {'0465': 'FWD', '0693': 'AFT', '0699': 'AFT', '0849': 'AFT',
                '1041': 'AFT', '1365': 'FWD', '1473': 'AFT', '1569': 'AFT',
                '1618+47': 'AFT', '1623': 'FWD'},
                'RH': {'0465': 'AFT', '0693': 'FWD', '0699': 'FWD', '0849': 'FWD',
                '1041': 'FWD', '1365': 'AFT', '1473': 'FWD', '1569': 'FWD',
                '1618+47': 'FWD', '1623': 'AFT'}}

    bushing = '1_FWD'
    side = 'FWD'

    if ref_instance.side_to_find == 'RH':
        side = 'AFT'
        bushing = '1_AFT'

    if '1X5005-210000' in [x[:x.find('##')] for x in ref_instance.all_irm_parts] and ref_instance.bin_order == ref_instance.irm_ln:
        bushing = ''

    try:
        sta_check = int(ref_instance.sta_to_find)
    except ValueError:
        pass
    else:
        if sta_check < 465 and ref_instance.bin_order == 1:
            return name + '_' + side + '_' + bushing
        elif sta_check < 465 and ref_instance.bin_order == ref_instance.irm_ln:
            return name + '_' + side

    try:
        name = name + '_' + sta_dict[ref_instance.side_to_find][ref_instance.sta_to_find]
    except KeyError:
        pass
    if bushing != '':
        return name + '_' + bushing
    else:
        return name

#if __name__ == "__main__":

    #ecs = Ref('787_9_KAL_ZB656', '0465', 'LH', 240, 1, 3, [], name='GLS_STA0561-0657_OB_LH_CAI')