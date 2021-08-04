import datetime
import time
import config
import string
import random
import collections

from pprint import pprint

import requests
from requests.auth import HTTPBasicAuth

def generate_uuid(prefix="",suffix=""):
    #Generates a SORMAS-Style uuid. The first up to 6 and last up to 8 letters can be defined manually.
    alphabet= string.ascii_uppercase + string.digits
    uuid = "-".join([''.join(random.choice(alphabet) for _ in range(6)) for _ in range(4)])+''.join(random.choice(alphabet) for _ in range(2))
    return prefix[:6]+uuid[len(prefix[:6]):len(uuid)-len(suffix[:8])]+suffix[:8]

def dict_merge(dct, merge_dct):
    # Recursive dictionary merge
    # Copyright (C) 2016 Paul Durivage <pauldurivage+github@gmail.com>
    # 
    # This program is free software: you can redistribute it and/or modify
    # it under the terms of the GNU General Public License as published by
    # the Free Software Foundation, either version 3 of the License, or
    # (at your option) any later version.
    """ Recursive dict merge. Inspired by :meth:``dict.update()``, instead of
    updating only top-level keys, dict_merge recurses down into dicts nested
    to an arbitrary depth, updating keys. The ``merge_dct`` is merged into
    ``dct``.
    :param dct: dict onto which the merge is executed
    :param merge_dct: dct merged into dct
    :return: None
    """
    for k, v in iter(merge_dct.items()):
        if (k in dct and isinstance(dct[k], dict) and isinstance(merge_dct[k], collections.Mapping)):
            dict_merge(dct[k], merge_dct[k])
        elif (k in dct and isinstance(dct[k], list) and isinstance(merge_dct[k], collections.Sequence)):    # Modified
            dct[k]+= merge_dct[k]                                                                           # Lists will be concatenated, not overwritten
        else:
            dct[k] = merge_dct[k]

def timestamp_to_datestring(timestamp: int):
    """Convert Unix timestamp to DD.MM.YYYY HH:SS (31.12.2020 23:59) - default to empty string for empty dates"""
    if not timestamp:
        return ""
    return datetime.datetime.fromtimestamp(timestamp/1000).strftime("%d.%m.%Y")

def datestring_to_int(datestring):
    # Convert DD.MM.YYYY HH:SS (31.12.2020 23:59) to Unix timestamp (int)
    # Also works for 31.12.2020
    if len(datestring) == 10:
        datestring += " 00:00"
    return int(time.mktime(datetime.datetime.strptime(datestring,"%d.%m.%Y %H:%M").timetuple()))*1000

def excel_to_timestamp(exceldate):
    # Convert Excel date (str) to UNIX timestamp (int)
    exceldate = exceldate.split(",")[0]
    if exceldate == "":
        return ""
    dt =  datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(exceldate) - 2)
    return int(time.mktime(dt.timetuple()))*1000

def excelint_to_datetuple(exceldate):
    # Convert Excel date int to (YYYY,MM,DD) tuple
    dt = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(exceldate) - 2)
    return dt.year, dt.month, dt.day

def now():
    return round(time.time())*1000

def is_adult(datestring):
    # Input: DD.MM.YYYY
    # Sloppy adult check
    return (datetime.datetime.today() - datetime.datetime.strptime(datestring,"%d.%m.%Y")).days>=6574

def get_since(table, datestring = "", dateint = 0):
    # Return json list of all items from table since datestring. Date Format: 31.12.2020 23:59
    # Unix time (as int) also accepted
    # table: persons, tasks, cases, contacs, samples, pathogentests
    # Info: datestring corresponds to changedate
    t=""
    if datestring:
        t=str(datestring_to_int(datestring))
    elif dateint:
        t=str(dateint)
    url = config.sormasurl + table + "/all/" + t
    print(url)
    response = requests.get(url, auth=HTTPBasicAuth(config.restuser, config.restpw), verify=config.verify)
    print(response)
    return response.json()
 
def get_all_case_uuids():
    # Return all case uuids, active or archived. Includes deleted cases.
    active_case_uuids = get_uuids("cases")
    archived_case_uuids = requests.get(config.sormasurl+"/cases/archived/0", auth=HTTPBasicAuth(config.restuser, config.restpw), verify=config.verify).json()
    return active_case_uuids + archived_case_uuids
 
def query(table, uuid):
    # Get one item from table by uuid
    # Return None if uuid not in table
    # TODO exception handling if uuid not available
    url = config.sormasurl + table + "/query/"
    response = requests.post(url, auth=HTTPBasicAuth(config.restuser, config.restpw), json=[uuid], verify=config.verify)
    #if len(response.json())==0:
    #    print("Warning - uuid "+uuid+" not available in table "+table)
    #    return None
    return response.json()[0]

def query_uuidlist(table, uuidlist):
    # Accepts list of uuids, returns list of entries
    url = config.sormasurl + table + "/query/"
    response = requests.post(url, auth=HTTPBasicAuth(config.restuser, config.restpw), json=uuidlist, verify=config.verify)
    return response.json()

def push(table, json):
    #pprint(json)
    # Creates a new entry. Will overwrite existing entry if uuid exists.
    url = config.sormasurl + table + "/push/"
    response = requests.post(url, auth=HTTPBasicAuth(config.restuser, config.restpw), json=[json], verify=config.verify)
    return response

def get_uuids(table):
    # returns list of uuids
    url = config.sormasurl + table + "/uuids/"
    response = requests.get(url, auth=HTTPBasicAuth(config.restuser, config.restpw), verify=config.verify)
    return response.json()


def create_person():
    pass

def create_default_person(persondict, personuuid):
    # Creates new person with data specified in persondict
    # Uses defaultvalues specified in config.person_template_uuid (Except address/phone/mail)
    person = query("persons", config.person_template_uuid)
    person.pop("creationDate", None)
    person.pop("changeDate", None)
    person.pop("address", None)
    #person.update(persondict)
    dict_merge(person, persondict)
    person["uuid"]=personuuid
    return push("persons", person)

def update_person(personuuid, persondict, prioritize_existing = True):
    # Update person with values specified in persondict
    person = query("persons", personuuid)
    person["changeDate"] = now()
    # Numbers / mail adresses:
    """
    if "personContactDetails" in person:
        contactdetaillist = [x["contactInformation"] for x in person["personContactDetails"]] #TODO: remove special characters to compare phone numbers <
        if "personContactDetails" in persondict:
            for item in persondict["personContactDetails"]:
                if item["contactInformation"] not in contactdetaillist:
                    item["primaryContact"] = False
                    person["personContactDetails"].append(item)
        persondict.pop("personContactDetails", None)
    """
    if not prioritize_existing:
        dict_merge(person, persondict)
        pprint(person)
        return push("persons", person)
    else:
        dict_merge(persondict, person)
        pprint(persondict)
        return push("persons", persondict)

def create_contact():
    pass

def create_default_contact(contactdict, personuuid):
    # Creates new person with data specified in persondict
    # Uses defaultvalues specified in config.contact_template_uuid (Except creator / changedates)
    contactuuid = generate_uuid(prefix=contactdict["caseIdExternalSystem"].zfill(6))
    cjson = query("contacts", config.contact_template_uuid)
    cjson["uuid"] = contactuuid
    cjson["person"] = { "uuid" : personuuid }
    cjson["district"] = { "uuid" : cjson["district"]["uuid"] }
    cjson["region"] = { "uuid" : cjson["region"]["uuid"] }
    cjson.pop("followUpStatusChangeDate", None)
    cjson.pop("creationdate", None)
    cjson["additionalDetails"] = contactdict["additionalDetails"]
    cjson.pop("epiData", None)            #TODO: clone entries
    cjson.pop("vaccinationInfo", None)
    cjson["healthConditions"]["uuid"] = generate_uuid(suffix = "XX")
    cjson["healthConditions"]["creationDate"] = now()
    cjson.pop("followUpStatusChangeUser", None)
    cjson["reportingUser"] = config.useruuid
    #json.update(contactdict)
    dict_merge(cjson, contactdict)
    pprint(cjson)
    return push("contacts", cjson)

def create_task(associateduuid, context, creatorComment = "", assigneeReply = "", assigneeUseruuid = config.useruuid, creatorUseruuid = config.useruuid, suggestedStart = now(), dueDate = now()+86400000, priority = "NORMAL", taskStatus = "PENDING", taskType = "OTHER"):
    # Creates new task, using default values if not defined otherwise.
    # Taskcontext: "CASE", "CONTACT" or "EVENT"
    tjson = {'assigneeReply': assigneeReply,
        'assigneeUser': { 'uuid': assigneeUseruuid },
        'creatorComment': creatorComment,
        'creatorUser': { 'uuid': creatorUseruuid },
        'dueDate': dueDate,
        'priority': priority,
        'suggestedStart': suggestedStart,
        'taskContext': context,
        'taskStatus': taskStatus,
        'taskType': taskType,
        'uuid': generate_uuid()}
    if context == "CASE":
        tjson["caze"] = { "uuid" : associateduuid }
    elif context == "CONTACT":
        tjson["contact"] = { "uuid" : associateduuid }
    elif context == "EVENT":
        tjson["event"] = { "uuid" : associateduuid }
    return push("tasks", tjson)

def update_task(taskuuid, changedict = {}, commentprefix = ""):
    # Changes existing task
    # Commentprefix: Puts a comment on top of creatorComment field
    tjson = query("tasks",taskuuid)
    if "caze" in tjson:
        tjson["caze"] = { "uuid" : tjson["caze"]["uuid"]}
    elif "contact" in tjson:
        tjson["contact"] = { "uuid" : tjson["contact"]["uuid"]}
    elif "event" in tjson:
        tjson["event"] = { "uuid" : tjson["event"]["uuid"]}
    tjson["creatorUser"] = { "uuid" : tjson["creatorUser"]["uuid"] }
    tjson["assigneeUser"] = { "uuid" : tjson["assigneeUser"]["uuid"] }
    tjson.pop("contextReference", None)
    for key in changedict:
        tjson[key] = changedict[key]
    if "creatorComment" in tjson:
        tjson["creatorComment"] = commentprefix + tjson["creatorComment"]
    else:
        tjson["creatorComment"] = commentprefix
    return push("tasks",tjson)

def update_case(uuid, changedict = {}, commentprefix=""):
    # Updates an existing case
    # Careful! Might result in data loss
    json = query("cases", uuid)
    l = ["responsibleDistrict","responsibleRegion","responsibleCommunity","district", "region", "community", "reportingDistrict", "reportingUser", "followUpStatusChangeUser", "surveillanceOfficer", "classificationUser","pointOfEntry"]
    for key in l:
        if key in json:
            json[key] = {"uuid": json[key]["uuid"]}
    if "epiData" in json:
        if "exposures" in json["epiData"]:
            for exposure in json["epiData"]["exposures"]:
                if "reportingUser" in exposure:
                    exposure["reportingUser"]={"uuid": exposure["reportingUser"]["uuid"] }
            for activity in json["epiData"]["activitiesAsCase"]:
                if "reportingUser" in activity:
                    activity["reportingUser"]={"uuid": activity["reportingUser"]["uuid"] }
    for key in changedict:
        json[key]=changedict[key]
    if not "additionalDetails" in json and commentprefix:
        json["additionalDetails"]=""
    if commentprefix:
        json["additionalDetails"]=commentprefix+json["additionalDetails"]
    return push("cases",json)
    
def update_contact(uuid, changedict = {}, commentprefix=""):
    json = query("contacts", uuid)
    l = ["responsibleDistrict","responsibleRegion","responsibleCommunity","district", "region", "community", "reportingUser", "contactOfficer"] #"caze"
    for key in l:
        if key in json:
            json[key] = {"uuid": json[key]["uuid"]}
    if "epiData" in json:
        if "exposures" in json["epiData"]:
            for exposure in json["epiData"]["exposures"]:
                if "reportingUser" in exposure:
                    exposure["reportingUser"]={"uuid": exposure["reportingUser"]["uuid"] }
            for activity in json["epiData"]["activitiesAsCase"]:
                if "reportingUser" in activity:
                    activity["reportingUser"]={"uuid": activity["reportingUser"]["uuid"] }
    for key in changedict:
        json[key]=changedict[key]
    if not "additionalDetails" in json and commentprefix:
        json["additionalDetails"]=""
    if commentprefix:
        json["additionalDetails"]=commentprefix+json["additionalDetails"]
    return push("contacts",json)

def update(table, uuid, changedict, prioritize_existing=False):
    json = query(table, uuid)
    #json.update(changedict)
    json["changeDate"]=now()
    if prioritize_existing:
        dict_merge(changedict, json)
        return push(table, changedict)
    else:
        dict_merge(json, changedict)
        return push(table, json)

def get_case_samples(caseuuid):
    url = config.sormasurl + "samples/query/cases"
    response  = requests.post(url, auth=HTTPBasicAuth(config.restuser, config.restpw), json=[caseuuid], verify=config.verify)
    return response.json()

def get_pathogentests(sampleuuid):
    url = config.sormasurl + "pathogentests/query/samples"
    response  = requests.post(url, auth=HTTPBasicAuth(config.restuser, config.restpw), json=[sampleuuid], verify=config.verify)
    return response.json()
    
def get_earliest_positive_test(caseuuid):
    #Return earliest sampledatetime of all positive tests - unix timestamp int
    samplejson = [x for x in get_case_samples(caseuuid) if x['pathogenTestResult'] == "POSITIVE"]
    if len(samplejson)==0:
        return ""
    else:
        return min([x['sampleDateTime'] for x in samplejson])

def get_earliest_positive_PCR(caseuuid):
    #Return earliest sampledatetime of all positive PCR tests - unix timestamp int
    samplejson = sorted(get_case_samples(caseuuid), key = lambda x: x['sampleDateTime'])
    for sample in samplejson:
        if len([1 for x in get_pathogentests(sample["uuid"]) if x['testResult']=="POSITIVE" and x["testType"]=="PCR_RT_PCR"])>0:
            return sample['sampleDateTime']
    return ""
        
