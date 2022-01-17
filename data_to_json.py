import json


test_dict = {
    "d9201703-3230-435d-9576-2d89c4e0d543": "test",
}


with open('data.json', 'w') as outfile:
    json.dump(test_dict, outfile)
