import sormasapi
import config
import os
from pprint import pprint


if __name__ == "__main__":
    if config.useproxy:
        os.environ['HTTPS_PROXY'] = config.proxy
    
    #caseuuid = "XXXXXX-XXXXXX-XXXXXX-XXXXXXXX" #Informationen für einen Fall anzeigen
    #case = sormasapi.query("cases",caseuuid)
    #pprint(case)
    
    #person = sormasapi.query("persons","XXXXXX-XXXXXX-XXXXXX-XXXXXXXX") #Informationen einer Person anzeigen
    #pprint(person)
    
    #contactjson = sormasapi.query("XXXXXX-XXXXXX-XXXXXX-XXXXXXXX") #Informationen eines Kontaktes anzeigen
    #pprint(contactjson)
