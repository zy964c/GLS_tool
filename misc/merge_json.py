import json
from jsonmerge import merge 

def merger():

    head_file = open(r'\\FIL-MOW01-01\787Payloads\IRC\SYSTEM_Int\personal_folder\Godov\GLS_clean\misc\lights_all\flagnote_ref.json')
    base_file = open(r'\\FIL-MOW01-01\787Payloads\IRC\SYSTEM_Int\personal_folder\Godov\GLS_clean\misc\lights_all\flagnote.json')
    
    head = json.load(head_file)
    base = json.load(base_file)
    result = merge(base, head)

    with open(r'\\FIL-MOW01-01\787Payloads\IRC\SYSTEM_Int\personal_folder\Godov\GLS_clean\misc\lights_all\vector_all.json', 'w') as m:
        json.dump(result, m, sort_keys=True, indent=4, separators=(',', ': '))

    head_file.close()
    base_file.close()


merger()


