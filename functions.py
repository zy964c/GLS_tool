# -*- coding: utf-8 -*-
"""
Created on Wed Dec 09 23:54:47 2015

@author: Roman
"""


def inch_to_mm(distance):

    """
    :param distance: distance in inches
    :return: distance in mm
    """

    return distance * 25.4


def mm_to_inch(distance):

    """
    :param distance: distance in mm
    :return: distance in inches
    """

    return distance / 25.4


def sta_value(coord, plug_value):

    """
    :param coord: x coordinate in airplane coordinate system in mm
    :param plug_value: 240 for -9, 456 for -10, 0 for -8
    :return: str with format 'STA....' or 'STA....+...'
    """

    STA = '0'
    if plug_value == 240:
        if round(coord, 1) <= round(inch_to_mm(609), 1):
            STA = '0' + str(int(round(coord / 25.4)))
        elif round(coord, 1) > round(inch_to_mm(609), 1) and coord <= round(inch_to_mm(609 + 120), 1):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif round(coord, 1) > round(inch_to_mm(609 + 120), 1) and coord <= round(inch_to_mm(1401 + 120), 1):
            if (coord / 25.4 - 120) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 120)))
            else:
                STA = str(int(round(coord / 25.4 - 120)))
        elif round(coord, 1) > round(inch_to_mm(1401 + 120), 1) and coord <= round(inch_to_mm((1401 + 120) + 120), 1):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 120))))
        elif round(coord, 1) > round(inch_to_mm(1401 + 240), 1):
            STA = str(int(round(coord / 25.4 - 240)))

    elif plug_value == 456:
        if round(coord, 1) <= round(inch_to_mm(609), 1):
            STA = '0' + str(int(round(coord / 25.4)))
        elif round(coord, 1) > round(inch_to_mm(609), 1) and coord <= round(inch_to_mm(609 + 240), 1):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif round(coord, 1) > round(inch_to_mm(609 + 240), 1) and coord <= round(inch_to_mm(1401 + 240), 1):
            if (coord / 25.4 - 240) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 240)))
            else:
                STA = str(int(round(coord / 25.4 - 240)))
        elif round(coord, 1) > round(inch_to_mm(1401 + 240), 1) and coord <= round(inch_to_mm((1401 + 240) + 120), 1):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 240))))
        elif round(coord, 1) > round(inch_to_mm(1401 + 360), 1) and coord <= round(inch_to_mm(1618 + 360), 1):
            STA = str(int(round(coord / 25.4 - 360)))
        elif round(coord, 1) > round(inch_to_mm(1618 + 360), 1) and coord <= round(inch_to_mm((1618 + 360) + 96), 1):
            STA = '1618+' + str(int(round(coord / 25.4 - (1618 + 360))))
        elif round(coord, 1) > round(inch_to_mm(1618 + 360 + 96), 1):
            STA = str(int(round(coord / 25.4 - (360 + 96))))

    elif plug_value == 0:
        if int(round(coord / 25.4)) < 1000:
            STA = '0' + str(int(round(coord / 25.4)))
        else:
            STA = str(int(round(coord / 25.4)))
    return STA


# path = '\\\\nw\\data\\irc-kmapi\\IRC_KBE\\Engineering_Automation\\ECS_Tool\\LIBRARY_NOGEOM_ICM2'
# path = '\\\\mow.boeing.ru\\dfs\\Home1\\zy964c\\My Documents\\ECS-tool\\LIBRARY_NOGEOM_ICM2'
