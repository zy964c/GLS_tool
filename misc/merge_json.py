import json
from jsonmerge import merge 


def merger():

    base_file = open('jd_vectors.json')
    head_file = open('jd_vectors_centers1.json')

    base = json.load(base_file)
    head = json.load(head_file)
    result = merge(base, head)

    with open('jd_vectors_3.30.json', 'w') as m:
        json.dump(result, m, sort_keys=True, indent=4, separators=(',', ': '))

    head_file.close()
    base_file.close()

merger()


