import sormasapi
import csv

# Read csv with changes
# Check if uuid in SORMAS
# (Check if case needs to be changed) -> Changedate / previous value
# Make a checkpoint after first case (Test case)

def csv_to_dict(filename):  #Convert csv file to nested dictionary, first order key: uuid, second order keys specified in file header
    with open(filename,encoding="UTF-8",newline="") as csvfile:
        csvreader = csv.reader(csvfile, delimiter=";", quotechar='"')
        d = {}
        header_keys = next(csvreader)
        for row in csvreader:
            dd=dict(zip(header_keys,row))
            d[row[0]]=dd
    return d 
        
def read_doubles(filename):
    d={}
    with open(filename,encoding="UTF-8",newline="") as csvfile:
        csvreader = csv.reader(csvfile, delimiter=";", quotechar='"')
        d = {}
        for row in csvreader:
            if row[1] not in d:
                d[row[1]]={}
                d[row[1]]["uuid"]=row[0]
                d[row[1]]["sex"]=row[4]
                d[row[1]]["birthdateDD"]=row[5]
                d[row[1]]["birthdateMM"]=row[6]
                d[row[1]]["birthdateYYYY"]=row[7]
            d[row[1]]["firstName"]=row[2]
            d[row[1]]["lastName"]=row[3]
            if row[4]:
                d[row[1]]["sex"]=row[4]
            if row[5]:
                d[row[1]]["birthdateDD"]=row[5]
            if row[6]:
                d[row[1]]["birthdateMM"]=row[6]
            if row[7]:
                d[row[1]]["birthdateYYYY"]=row[7]
    return d 
        
if __name__ == "__main__":
    doubles = read_doubles("dedupe.csv")
    for key in doubles:
        print(key)
        puuid=sormasapi.query("cases",doubles[key]["uuid"])["person"]["uuid"]
        print(puuid)
        changedict={}
        for x in ["firstName","lastName","sex","birthdateDD","birthdateMM","birthdateYYYY"]:
            changedict[x]=doubles[key][x]
        #print(sormasapi.query("persons",puuid))
        print(changedict)
        stat=sormasapi.update_person(puuid,changedict,prioritize_existing = False)
        if not "200" in str(stat):
            print(stat)
            input("Press Enter to continue...")

    """
    changes = csv_to_dict("qend.csv")
    for key in changes:
        changedict = changes[key]
        changedict["quarantineTo"]=sormasapi.datestring_to_int(changedict["quarantineTo"])
        print(key)
        print(changedict)
        stat = sormasapi.update_case(key, changedict)
        
        if not "200" in str(stat):
            print(stat)
            input("Press Enter to continue...")
    """
        
