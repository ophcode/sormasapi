"""
Work in progress:
- Does a database dump of all cases / persons in database
- When restarted, only updates cases / persons that have been changed since last use
- Potentially useful for automated statistics
"""

import sormasapi
import config
import time
import os
import json
from pprint import pprint

def initialize_json(filename, save_interval = 100):
    print("Getting all data... might take a while")
    casedict = {}
    uuids = sormasapi.get_all_case_uuids()
    print(str(len(uuids)) + " Fälle insgesamt")
    if os.path.exists(filename):
        with open(filename) as jsonfile:
            casedict = json.load(jsonfile)
    print(str(len(casedict)) + " Fälle aus Datei geladen")
    i = 0
    for uuid in uuids:
        if uuid not in casedict:
            case = sormasapi.query("cases",uuid)
            casedict[uuid] = case
            i+=1
            if i == save_interval:
                i=0
                print(str(len(casedict))),
                with open(filename, "w") as outfile:
                    json.dump(casedict, outfile)
                    
def initialize_person_json():
    pass

def load_and_update_json(filename, override_date=""):
    # 1) Reads json dump 
    # 2) Updates all cases updated since last dump 
    # 3) Adds new cases
    # override_date: force update since date / datetime (dd.mm.yyyy / dd.mm.yyyy HH:SS)
    casedict = {}
    with open(filename) as jsonfile:
        casedict = json.load(jsonfile)
    changedate = max([casedict[key]["changeDate"] for key in casedict])
    if override_date:
        changedate = sormasapi.datestring_to_int(override_date)
    new_cases = sormasapi.get_since("cases", dateint = changedate)
    for case in new_cases:
        casedict[case["uuid"]] = case
    with open(filename, "w") as outfile:
        json.dump(casedict, outfile)
    print("Done")
    return casedict
    
def 


if __name__== "__main__":
    os.environ['HTTPS_PROXY'] = config.proxy
    start_time = time.time()
    filename = "all_cases.json"
    #initialize_json(filename)
    casedict = load_and_update_json(filename)
    print(len(casedict))
    print("--- %s seconds ---" % (time.time() - start_time))
    
    