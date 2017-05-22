from pprint import pprint


class Annotation(object):

    views = ['FWD Facing AFT - Text Plane LH',
             'Inboard Facing Out - Lower Support - Text Plane LH',
             'Outboard Facing In - Middle Support - Text Plane LH',
             'FWD Facing AFT - Text Plane RH',
             'Inboard Facing Out - Lower Support - Text Plane RH',
             'Outboard Facing In - Middle Support - Text Plane RH',
             'FWD Facing AFT - Text Plane 41 LH',
             'Inboard Facing Out - Lower Support - Text Plane 41 LH',
             'Inboard Facing Out - Top Support - Text Plane 41 LH',
             'FWD Facing AFT - Text Plane 41 RH',
             'Inboard Facing Out - Lower Support - Text Plane 41 RH',
             'Inboard Facing Out - Top Support - Text Plane 41 RH',
             'Outboard Facing In - Middle Support - Text Plane 47 LH',
             'Inboard Facing Out - Top Support - Text Plane 47 LH',
             'Outboard Facing In - Middle Support - Text Plane 47 RH',
             'Inboard Facing Out - Top Support - Text Plane 47 RH',
             'Inboard Facing Out - Top Support - Text Plane LH',
             'Inboard Facing Out - Top Support - Text Plane RH',
             'Outboard Facing In - Middle Support - Text Plane 41 LH',
             'Outboard Facing In - Middle Support - Text Plane 41 RH',
             'FWD Facing AFT - Text Plane 47 LH',
             'Inboard Facing Out - Lower Support - Text Plane 47 LH',
             'FWD Facing AFT - Text Plane 47 RH',
             'Inboard Facing Out - Lower Support - Text Plane 47 RH']

    views_ctr = ['FWD Facing AFT - Text Plane',
                 'Outboard LH Facing In - Upper Support - Text Plane',
                 'Outboard RH Facing In - Upper Support - Text Plane',
                 'JD40 - Section View - Text Plane']

    def __init__(self, annot_name, annot_view, capture_name):
        self.annot_name = annot_name
        self.annot_view = annot_view
        self.capture_name = capture_name

    def get_annot_view(self):

        return self.annot_view

    def get_annot_name(self):

        return self.annot_name

    def get_capture_name(self):

        return self.capture_name


class AnnotationFactory(object):

    def __init__(self, irm_type=1):

        self.irm_type = irm_type
        self.annots = {'FL1': ['Inboard Facing Out - Lower Support', 'Wire Harness And Wire Splice Location'],
                       'FL2': ['Inboard Facing Out - Lower Support', 'Wire Harness And Wire Splice Location'],
                       'FL3': ['Inboard Facing Out - Top Support', 'Ceiling Light Marker Installation'],
                       'FL4': ['FWD Facing AFT', 'Sidewall Light'],
                       'FL5': ['Inboard Facing Out - Lower Support', ['Pigtail Upper', 'Pigtail Lower']],
                       'FL6': ['Outboard Facing In - Middle Support', 'Sidewall Light Marker Installation'],
                       'FL7': ['Inboard Facing Out - Lower Support', 'Mounting Blocks'],
                       'FL8': ['Inboard Facing Out - Lower Support', 'Wire Harness Tie'],
                       'FL9': ['Inboard Facing Out - Lower Support', 'Connector Installation'],
                       'FL10': ['Inboard Facing Out - Lower Support',
                                'Sidewall Light Termination Plug Marker Installation'],
                       'FL11': ['Inboard Facing Out - Top Support', 'Connector to P-Clamp Installation'],
                       'FL12': ['Inboard Facing Out - Top Support',
                                'Ceiling Light Termination Plug Marker Installation'],
                       'FL13': ['Inboard Facing Out - Lower Support', 'Wire Tie Installation'],
                       'FL15': ['Outboard Facing In - Middle Support', 'Riding Condition'],
                       'FL16': ['Inboard Facing Out - Top Support', 'NOFAR Termination Plug Marker Installation'],
                       'FL17': ['Outboard Facing In - Middle Support', 'Wire Harness of  1J5009-3010(XX) Installation'],
                       'FL18': ['Inboard Facing Out - Top Support', 'NOFAR Light Marker Installation'],
                       'FL19': ['Inboard Facing Out - Top Support', 'Power Supply Marker Installation'],
                       'FL20': ['Outboard LH Facing In - Upper Support', 'Wire Harness to Ceiling Light'],
                       'FL28': ['Inboard Facing Out - Lower Support', 'Wire Bundle'],
                       'FL29': ['Inboard Facing Out - Lower Support', 'Bushing Installation'],
                       'FL38': ['Inboard Facing Out - Lower Support', 'Bushing Installation'],
                       'JD 01': ['FWD Facing AFT', 'JD01 Ceiling Light Typical'],
                       'JD 02': ['Inboard Facing Out - Lower Support', 'JD02 Noise Seal Typical'],
                       'JD 03': ['Inboard Facing Out - Lower Support', 'JD03 Sidewall Light Bracket Typical'],
                       'JD 04': ['Inboard Facing Out - Lower Support', 'JD04 1510-1 Bracket Typical'],
                       'JD 05': ['Outboard Facing In - Middle Support', 'JD05 1510-2 Bracket Typical'],
                       'JD 06': ['Outboard Facing In - Middle Support', 'JD06 1510-3 Bracket Typical'],
                       'JD 07': ['Outboard Facing In - Middle Support', 'JD07 1510-22 Bracket Typical'],
                       'JD 08': ['Inboard Facing Out - Lower Support', 'JD08 1510-24 Bracket Typical'],
                       'JD 09': ['Inboard Facing Out - Lower Support', 'JD09 Ring Post to Bin Typical'],
                       'JD 10': ['Outboard Facing In - Middle Support',
                                 'JD10 Installation of Ring Posts (double) to Fairings.'],
                       'JD 11': ['Outboard Facing In - Middle Support', 'JD11 Ring Post to 1510-2 Bracket'],
                       'JD 12': ['Outboard Facing In - Middle Support', 'JD12 Ring Post to 1510-3 Bracket Typical'],
                       'JD 13': ['Outboard Facing In - Middle Support', 'JD13 Ring Post to 1510-22 Bracket Typical'],
                       'JD 14': ['FWD Facing AFT', 'JD14 Term Plug P-clamp Typical'],
                       'JD 15': ['Inboard Facing Out - Lower Support', 'JD15 Sight Block to Stowbin Typical'],
                       'JD 17': ['Inboard Facing Out - Lower Support', 'JD17 Single Wire Bushings Typical'],
                       'JD 19': ['Inboard Facing Out - Lower Support', 'JD19 Outboard Latch Typical'],
                       'JD 20': ['Inboard Facing Out - Lower Support', 'JD20 Outboard Latch to Risers'],
                       'JD 21': ['Inboard Facing Out - Lower Support', 'JD21 OFCR Outboard Latch Typical'],
                       'JD 22': ['Inboard Facing Out - Lower Support', 'JD22 Sidewall Light Bracket Typical'],
                       'JD 23': ['Inboard Facing Out - Lower Support', 'JD23 Sidewall Light Bracket Typical'],
                       'JD 24': ['Inboard Facing Out - Lower Support', 'JD24 Riser to Ceiling Lights'],
                       'JD 25': ['Inboard Facing Out - Lower Support', 'JD25 Outboard Latch To SLC Typical'],
                       'JD 26': ['Inboard Facing Out - Top Support', 'JD26 Outboard NOFAR Light Typical'],
                       'JD 27': ['Inboard Facing Out - Top Support', 'JD27 Power Supply Typical'],
                       'JD 28': ['Inboard Facing Out - Lower Support', 'JD28 Ring post to Rail Typical'],
                       'JD 29': ['Outboard Facing In - Middle Support', 'JD29 1510-1 Bracket to Rail Typical'],
                       'JD 31': ['Outboard Facing In - Middle Support', 'JD31 Saddle Clamp Typical'],
                       'JD 40': ['JD40 - Section View', 'Joint Definition 40'],
                       'JD 41': ['Outboard LH Facing In - Upper Support', 'Joint Definition 41'],
                       'JD 42': ['Outboard LH Facing In - Upper Support', 'Joint Definition 42'],
                       'sta': ['Inboard Facing Out - Lower Support', 'Reference Geometry'],
                       '1X5005-412(XXX) Alignment Feature': ['FWD Facing AFT', 'JD01 Ceiling Light Typical'],
                       '1X5005-412(XXX) Mounting Feature': ['FWD Facing AFT', 'JD01 Ceiling Light Typical'],
                       'Fixing Tower': ['FWD Facing AFT', 'Sidewall Light'],
                       '1X5005-420(X)00 Mounting Bracket': ['FWD Facing AFT', 'Sidewall Light'],
                       'Locking Tabs': ['FWD Facing AFT', 'Sidewall Light'],
                       'light_marker': ['Inboard Facing Out - Top Support', 'Ceiling Light Marker Installation'],
                       'term_plug_marker_sidewall': ['Inboard Facing Out - Lower Support',
                                                     'Sidewall Light Termination Plug Marker Installation'],
                       'term_plug_marker_ceiling': ['Inboard Facing Out - Top Support',
                                                    'Ceiling Light Termination Plug Marker Installation'],
                       'term_plug_marker_nofar': ['Inboard Facing Out - Top Support',
                                                  'NOFAR Termination Plug Marker Installation'],
                       'sidewall_light_marker': ['Outboard Facing In - Middle Support',
                                                 'Sidewall Light Marker Installation'],
                       'nofar_light_marker': ['Inboard Facing Out - Top Support',
                                              'NOFAR Light Marker Installation'],
                       'power_supply_marker': ['Inboard Facing Out - Top Support',
                                               'Power Supply Marker Installation']}

        self.annots_ctr = {'FL2': ['FWD Facing AFT - Text Plane', 'Wire Splice Location'],
                           'FL3': [['Outboard LH Facing In - Upper Support - Text Plane',
                                    'Outboard RH Facing In - Upper Support - Text Plane'],
                                  ['Ceiling Light Marker Installation LH', 'Ceiling Light Marker Installation RH']],
                           'FL5': [['Outboard LH Facing In - Upper Support - Text Plane',
                                    'Outboard RH Facing In - Upper Support - Text Plane'], 'Pigtail Upper'],
                           'FL8': ['FWD Facing AFT - Text Plane', 'Wire Harness Tie'],
                           'FL12': [['Outboard LH Facing In - Upper Support - Text Plane',
                                    'Outboard RH Facing In - Upper Support - Text Plane'],
                                   ['FL12_LH', 'FL12_RH']],
                           'FL16': ['FWD Facing AFT - Text Plane',
                                   ['NOFAR Termination Plug Marker Installation LH',
                                    'NOFAR Termination Plug Marker Installation RH']],
                           'FL17': ['FWD Facing AFT - Text Plane',
                                    'Wire Harness of  1J5009-3010(XX) Installation'],
                           'FL18': ['FWD Facing AFT - Text Plane', ['NOFAR Light Marker Installation LH',
                                                                    'NOFAR Light Marker Installation RH']],
                           'FL19': ['FWD Facing AFT - Text Plane', ['Power Supply Marker Installation LH',
                                                                    'Power Supply Marker Installation RH']],
                           'FL20': [['Outboard RH Facing In - Upper Support - Text Plane',
                                    'Outboard LH Facing In - Upper Support - Text Plane'],
                                    'Wire Harness to Ceiling Light'],
                           'FL21': [['Outboard LH Facing In - Upper Support - Text Plane',
                                    'Outboard RH Facing In - Upper Support - Text Plane'],
                                    'FL21 Typical for BACM14L'],
                           'FL22': ['JD40 - Section View - Text Plane', 'Wire Harness to Ceiling Light'],
                           'FL36': ['FWD Facing AFT - Text Plane', 'FL22 Typical'],
                           'JD 04': ['FWD Facing AFT - Text Plane', 'Joint Definition 04'],
                           'JD 09': ['FWD Facing AFT - Text Plane', 'Joint Definition 09'],
                           'JD 26': ['FWD Facing AFT - Text Plane', 'Joint Definition 26'],
                           'JD 40': ['JD40 - Section View - Text Plane', 'Joint Definition 40'],
                           'JD 41': ['Outboard LH Facing In - Upper Support - Text Plane', 'Joint Definition 41'],
                           'JD 42': ['Outboard LH Facing In - Upper Support - Text Plane', 'Joint Definition 42'],
                           'JD 45': ['FWD Facing AFT - Text Plane', 'Joint Definition 45'],
                           'JD 46': ['FWD Facing AFT - Text Plane', 'Joint Definition 46'],
                           'JD 47': ['FWD Facing AFT - Text Plane', 'Joint Definition 47'],
                           'JD 48': ['FWD Facing AFT - Text Plane', 'Joint Definition 48'],
                           'JD 49': ['FWD Facing AFT - Text Plane', 'Joint Definition 49'],
                           'JD 52': ['FWD Facing AFT - Text Plane', 'Joint Definition 52'],
                           'JD 53': ['FWD Facing AFT - Text Plane', 'Joint Definition 53'],
                           'sta': [['Outboard LH Facing In - Upper Support - Text Plane',
                                    'Outboard RH Facing In - Upper Support - Text Plane'],
                                   'Reference Geometry'],
                           'light_marker': [['Outboard LH Facing In - Upper Support - Text Plane',
                                            'Outboard RH Facing In - Upper Support - Text Plane'],
                                            ['Ceiling Light Marker Installation LH',
                                             'Ceiling Light Marker Installation RH']],
                           'term_plug_marker_ceiling': [['Outboard LH Facing In - Upper Support - Text Plane',
                                                        'Outboard RH Facing In - Upper Support - Text Plane'],
                                                        ['FL12_LH', 'FL12_RH']],
                           'term_plug_marker_nofar': ['FWD Facing AFT - Text Plane',
                                                      ['NOFAR Termination Plug Marker Installation LH',
                                                       'NOFAR Termination Plug Marker Installation RH']],
                           'nofar_light_marker': ['FWD Facing AFT - Text Plane',
                                                 ['NOFAR Light Marker Installation LH',
                                                  'NOFAR Light Marker Installation RH']],
                           'power_supply_marker': ['FWD Facing AFT - Text Plane',
                                                  ['Power Supply Marker Installation LH',
                                                   'Power Supply Marker Installation RH']]
                           }

    def get_annot_dict(self):
        collection = []
        views_dict1 = {}
        d = self.annots
        if self.irm_type == 2:
            d = self.annots_ctr
        if len(d) > 0:
            for key in d:
                a = Annotation(key, d[key][0], d[key][1])
                collection.append(a)
            for item in collection:
                views_dict1[item.get_annot_name()] = self.add_annot(item)
            return views_dict1
        else:
            return None

    def add_annot(self, a):

        if self.irm_type == 1:
            return {'LH': {'constant': a.get_annot_view() + ' - Text Plane LH',
                           '41': a.get_annot_view() + ' - Text Plane 41 LH',
                           '47': a.get_annot_view() + ' - Text Plane 47 LH'},
                    'RH': {'constant': a.get_annot_view() + ' - Text Plane RH',
                           '41': a.get_annot_view() + ' - Text Plane 41 RH',
                           '47': a.get_annot_view() + ' - Text Plane 47 RH'},
                    'capture': a.get_capture_name()}

        return {'CTR': {'constant': a.get_annot_view(),
                        '41': a.get_annot_view(),
                        '47': a.get_annot_view()},
                'capture': a.get_capture_name()}

    def get_view_number(self, annot_name, side, section):

        try:
            dict1 = self.get_annot_dict()
            view_dict = Annotation.views
            if self.irm_type == 2:
                view_dict = Annotation.views_ctr
            view_name = dict1[annot_name][side][section]
            if isinstance(view_name, list):
                result = []
                for view in view_name:
                    result.append(view_dict.index(view))
                return result
            return view_dict.index(view_name)
        except:
            return None

    def get_view_name(self, annot_name, side, section):

        dict1 = self.get_annot_dict()
        try:
            result = dict1[annot_name][side][section]
        except:
            return None
        return result

    def get_capture(self, annot_name):

        dict1 = self.get_annot_dict()
        try:
            result = dict1[annot_name]['capture']
        except:
            return None
        return result

if __name__ == "__main__":

    # assert views.index(inbd_facing_out_lwr_support_41_lh) == 7
    # pprint(views_dict)
    # assert view_index('JOINT DEFINITION 01', 'LH', 'constant') == 0
    # assert view_index('sta', 'LH', 'constant') == 1
    # assert views_dict['FL2']['capture'] == 'FL1 and FL2 Typical'
    annotations = AnnotationFactory(irm_type=2)
    c = []
    for key in annotations.annots_ctr:
        captures = annotations.annots_ctr.get(key)[-1]
        if isinstance(captures, list):
            for i in captures:
                c.append(i)
        else:
            c.append(captures)
    with open('ctr_captures.txt', 'w') as f:
        for j in c:
            f.write(j + '\n')

    #look_at_dict = annotations.get_annot_dict()
    #pprint(look_at_dict)
    # b = Annotation('FL1', 'Inboard Facing Out - Lower Support', 'FL1 and FL2 Typical')
    # print b.get_annot_name()
    #print annotations.get_view_number('sta', 'CTR', 'constant')
    #print annotations.get_view_name('sta', 'CTR', 'constant')
    print annotations.get_capture('JD 48')
