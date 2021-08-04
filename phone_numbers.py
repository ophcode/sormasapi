"""
Script to fix issue that contactPersonPhone / contactPersonEmail are not displayed anymore by default
Copies entries entered into those two fields to personContactDetails
Tries to avoid duplicate entries and unnecessary API queries, still does not work in one session for larger databases
"""


import sormasapi
import config
import os
from pprint import pprint
import time

def move_number(person):
    # Input: Json from "person" table
    # Moves contactPersonEmail & contactPersonPhone to personContactDetails
    # If contactPersonEmail & don't exist or are identical to existing entries nothing is changed.
    
    if "address" in person:
        if person["address"].get("contactPersonEmail","") or person["address"].get("contactPersonPhone",""):
            mailexists = False
            phoneexists = False
            if not "personContactDetails" in person:
                person["personContactDetails"] = []
            if person["address"].get("contactPersonEmail",""):
                primary = True
                for pcd in person["personContactDetails"]:
                    if pcd.get("personContactDetailType","") == "EMAIL" and pcd.get("primaryContact",False):
                        primary = False
                    if pcd.get("contactInformation","") == person["address"]["contactPersonEmail"]:
                        mailexists = True
                if not mailexists:    
                    email =   {"changeDate": sormasapi.now(),
                               "contactInformation": person["address"]["contactPersonEmail"],
                               "creationDate": sormasapi.now(),
                               "person": {"uuid": person["uuid"]},
                               "personContactDetailType": "EMAIL",
                               "primaryContact": primary,
                               "pseudonymized": False,
                               "thirdParty": False,
                               "uuid": sormasapi.generate_uuid(prefix="XXX")}
                    person["personContactDetails"].append(email)
            if person["address"].get("contactPersonPhone",""):
                primary = True
                for pcd in person["personContactDetails"]:
                    if pcd.get("personContactDetailType","") == "PHONE" and pcd.get("primaryContact",False):
                        primary = False
                    if pcd.get("contactInformation","") == person["address"]["contactPersonPhone"]:
                        phoneexists = True
                if not phoneexists:
                    phone =   {"changeDate": sormasapi.now(),
                               "contactInformation": person["address"]["contactPersonPhone"],
                               "creationDate": sormasapi.now(),
                               "person": {"uuid": person["uuid"]},
                               "personContactDetailType": "PHONE",
                               "primaryContact": primary,
                               "pseudonymized": False,
                               "thirdParty": False,
                               "uuid": sormasapi.generate_uuid(prefix="XXX")}
                    person["personContactDetails"].append(phone)
            if (person["address"].get("contactPersonEmail","") and not mailexists) or (person["address"].get("contactPersonPhone","") and not phoneexists):
                person["changeDate"] = sormasapi.now()
                return sormasapi.push("persons", person)
    return "."

def move_all_numbers(since="01.01.2020 12:00"):
    persons = sormasapi.get_since("persons",since)
    for person in persons:
        print(person["uuid"])
        print(move_number(person))

def test(personuuid):    
    json = sormasapi.query("persons",personuuid)
    print(move_number(personuuid))
    json = sormasapi.query("persons",personuuid)
    pprint(json)
    
if __name__ == "__main__":
    os.environ['HTTPS_PROXY'] = config.proxy
    
    # For testing purposes enter uuid of one person here
    # test("XXXXXX-XXXXXX-XXXXXX-XXXXXXXX")

    start_time = time.time()
    move_all_numbers("25.07.2021 12:00")
    print("--- %s seconds ---" % (time.time() - start_time))

