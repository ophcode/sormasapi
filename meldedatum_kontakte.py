import sormasapi
from pprint import pprint
import config
import os

#Einmal-Skript zur Korrektur von Kontakten, bei denen durch einen Import alle Jahreszahlen auf das Jahr 0021 gesetzt wurden

if __name__ == "__main__":
    os.environ['HTTPS_PROXY'] = config.proxy
    jsonlist = sormasapi.get_since("contacts","31.07.2021")
    #print(sormasapi.datestring_to_int("27.07.2021"))
    #print(sormasapi.datestring_to_int("27.07.2021")+61486736400000)
    for json in jsonlist:
        print(json["uuid"])
        if "YOG" in json.get("externalID",""):
            if json.get("quarantineFrom",0)<0:
                changedict = {
                    "reportDateTime" : sormasapi.datestring_to_int("31.07.2021 00:00"),
                    "quarantineFrom" :json["quarantineFrom"]+63114073200000,
                    "quarantineTo" : json["quarantineFrom"]+63114073200000+1209600000,
                    "lastContactDate" : json["quarantineFrom"]+63114073200000,
                    "followUpUntil" : json["quarantineFrom"]+63114073200000+1209600000,
                    "quarantineOrderedOfficialDocumentDate" : sormasapi.datestring_to_int("31.07.2021 00:00")
                }
                print(sormasapi.update_contact(json["uuid"], changedict))
                j = sormasapi.query("contacts",json["uuid"])

            
    """
    changedict = {"reportDateTime" : sormasapi.datestring_to_int("31.07.2021 00:00")}
    sormasapi.update_contact(uuid, changedict)
    json = sormasapi.query("contacts",uuid)
    pprint(json)
    """
    