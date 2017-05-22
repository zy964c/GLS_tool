import json
from jsonmerge import merge 


def merger():

    base_file = open('flagnote.json')
    head_file = open('flagnote_ref.json')

    base = json.load(base_file)
    head = json.load(head_file)
    result = merge(base, head)

    with open('flagnote_new.json', 'w') as m:
        json.dump(result, m, sort_keys=True, indent=4, separators=(',', ': '))

    head_file.close()
    base_file.close()

merger()


