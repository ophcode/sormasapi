import sormasapi
import config

import csv
import datetime
import sys
import os
import re
import shutil

from mailmerge import MailMerge #Externe Library für docx-Serienbriefe
from docx2pdf import convert    #Externe Library für pdf-Erstellung mit Word
import win32com.client          #Externe Library für Email-Versand per Outlook

import tkinter as tk
import tkinter.filedialog as fd
import tkinter.simpledialog as sd

from pprint import pprint

def fill_file(input_docx_path,output_docx_path,fielddict): #Serienbrief erstellen, keys in fielddict müssen keys in docx entsprechen
    with MailMerge(input_docx_path) as document:
        document.merge_templates([fielddict],separator="page_break")
        document.write(output_docx_path)

def tasks_by_date(startint, endint=sormasapi.now()):
    return [t for t in sormasapi.get_since("tasks", dateint=startint) if t["statusChangeDate"]<=endint]

def mark_mail_as_sent(uuid, note="", context="cases"):
    changedict={
        "quarantineOrderedOfficialDocument" : True,
        "quarantineOrderedOfficialDocumentDate" : sormasapi.now()
    }
    if context=="cases":
        sormasapi.update_case(uuid, changedict, note)
    elif context=="contacts":
        sormasapi.update_contact(uuid, changedict, note)

def mark_task_as_completed(taskuuid):
    sormasapi.update_task(taskuuid, commentprefix="✅")

def send_mail(email,subject,body,attachment_path_list):
    outlook = win32com.client.Dispatch("outlook.application")
    msg = outlook.CreateItem(0)
    msg.To = email
    msg.Subject = subject
    msg.HTMLbody = body
    msg.SentOnBehalfOfName = "coronakontakt@ba-fk.berlin.de" #Getestet, funktioniert falls Zugriff auf das Konto besteht
    for attachment_path in attachment_path_list:
        att = os.path.abspath(attachment_path)
        msg.Attachments.Add(att)
    msg.display()   #Email anzeigen, kein automatisches senden #msg.Send()  
   
def convert_date(datestring): #Return DD.MM.YYYY if String is YYYY-MM-DD (HH:MM:SS.xxxxxx), return String otherwise.
    if re.match(r"[0-9][0-9][0-9][0-9]\-[0-9][0-9]\-[0-9][0-9] *",datestring) and len(datestring)<30:
        return datestring[8:10]+"."+datestring[5:7]+"."+datestring[:4]
    return datestring

def is_adult(birthdate): #Date format: yyyy-mm-dd
    #Defaults to 'True' in case of missing or malformated data
    today = datetime.date.today()
    try:
        person_birthdate=datetime.date.fromisoformat(birthdate)
    except:
        return True
    if (today-person_birthdate).days > 6574:
        return True
    return False
 
def get_mail(PLZ):
    #Retrieve mail address of responsible health department by postal code / (place name)
    URL = "https://tools.rki.de/PLZTool/?q=" + str(PLZ)
    page = requests.get(URL)
    if 'title="E-Mail"' in page.text:
        return page.text.split('href="mailto:')[1].split(">@</a>")[0]
    else:
        return ""
    
def get_fax(PLZ):
    #Retrieve fax address of responsible health department by postal code / (place name)
    URL = "https://tools.rki.de/PLZTool/?q=" + str(PLZ)
    page = requests.get(URL)
    if 'title="COVID-19 Fax">' in page.text:
        return page.text.split('title="COVID-19 Fax">')[1].split('value="')[1].split('"')[0]
    elif 'title="Fax">' in page.text:
        return page.text.split('title="Fax">')[1].split('value="')[1].split('"')[0]
    else:
        return ""
 
class Application(tk.Frame): #GUI
    def __init__(self, master=None):
        super().__init__(master)
        self.buttonfont=('Berlin Type Office Regular',18)
        self.bw= 30 # buttwonwidth
        self.bh= 2  # buttonheight
        self.bc1= "white"  # button color (unactived)
        self.bc2= "ghost white"  # button color (activated)
        self.startint=sormasapi.now()-604800000
        self.endint=sormasapi.now()
        self.Sachbearbeiter=config.stellenzeichen
        self.lb = ""
        self.master = master
        self.pack()
        self.create_widgets()
        self.mail_win = ""
        self.list_win = ""
        self.mail_answer = ""
        self.outputfolder=datetime.datetime.today().strftime("%Y-%m-%d")
        if not os.path.exists(self.outputfolder):
            os.makedirs(self.outputfolder)
        if not os.path.exists(os.path.join(self.outputfolder,"cases")):
            os.makedirs(os.path.join(self.outputfolder,"cases"))
            os.makedirs(os.path.join(self.outputfolder,"contacts"))
            os.makedirs(os.path.join(self.outputfolder,"pdf"))
        self.tasks = []
        self.index_tasks = []
        self.contact_tasks = []
        self.inputdocxpath=os.path.join("Vorlagen")
        
    def create_widgets(self):
        self.input_startdatetime = tk.Button(self,text= " 1. Startzeitpunkt wählen", command=self.choose_startdate, anchor="w", font=self.buttonfont, width=self.bw, height=self.bh, bg=self.bc1)
        self.input_startdatetime.pack(side="top")
        self.input_enddatetime = tk.Button(self, text=" 2. Endzeitpunkt wählen (optional)", command=self.choose_enddate, anchor="w", font=self.buttonfont, width=self.bw, height=self.bh, bg=self.bc1)
        self.input_enddatetime.pack(side="top")
        self.input_sachbearbeiter = tk.Button(self, text=" 3. Sachbearbeiter*in", command=self.choose_sachbearbeiter, anchor="w", font=self.buttonfont,width=self.bw, height=self.bh, bg=self.bc1)
        self.input_sachbearbeiter.pack(side="top")
        self.run = tk.Button(self, text=" 4. Index-Mails", command=self.index_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='white')
        self.run.pack(side="top")
        self.run_contacts = tk.Button(self, text=" 5. KP-Mails", command=self.contact_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='white')
        self.run_contacts.pack(side="top")
        self.quit = tk.Button(self, text=" BEENDEN", command=self.master.destroy, anchor="w", font=self.buttonfont, width=self.bw, bg='misty rose')
        self.quit.pack(side="bottom")
    
    def mail_indices(self):
        for i in self.lb.curselection():
            casejson = sormasapi.query("cases", self.index_tasks[i]["caze"]["uuid"])
            pprint(casejson)
            personjson = sormasapi.query("persons",casejson["person"]["uuid"])
            pprint(personjson)
            taskjson=self.index_tasks[i]
            if taskjson["taskStatus"]=="DONE":
                self.send_one_mail(self.index_tasks[i]["caze"]["uuid"], taskjson)
            elif taskjson["taskStatus"]=="NOT_EXECUTABLE":
                answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Nicht-erreicht-Brief drucken?\n"+taskjson.get("creatorComment","")+"\n"+taskjson.get("assigneeReply",""))
                if answer:
                    filename = self.create_docx(casejson, personjson, "Anschreiben nicht erreichte Indices", suffix="_ne", context="cases")
                    os.startfile(filename)
                    shutil.copy(filename,os.path.join(self.outputfolder,"cases",filename.split("\\")[-1]))
                    mark_mail_as_sent(casejson["uuid"], note='"Nicht erreicht"-Anschreiben am '+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n")
                    sormasapi.update_task(taskjson["uuid"], commentprefix="✉")
                    
    def mail_contacts(self):
        for i in self.lb.curselection():
            contactjson = sormasapi.query("contacts", self.contact_tasks[i]["contact"]["uuid"])
            pprint(contactjson)
            personjson = sormasapi.query("persons",contactjson["person"]["uuid"])
            pprint(personjson)
            taskjson=self.contact_tasks[i]
            if contactjson.get("contactClassification","") in ["CONFIRMED","UNCONFIRMED"] and not contactjson.get("quarantineOrderedOfficialDocument",False):
                #mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if x["personContactDetailType"]=="EMAIL"]).pop()
                mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop() #Quack
                if taskjson["taskStatus"]=="DONE":
                    text = taskjson.get("creatorComment","")+"\n"+taskjson.get("assigneeReply","")
                    text += "\nQ: "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineFrom",""))+" bis "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineTo",""))
                    answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Mail schicken?\n"+text)
                    if answer:
                        filename = self.create_docx(contactjson, personjson, "Isolationsbescheinigung", suffix="", context="contacts")
                        self.send_contact_mail(contactjson, personjson, mail, filename)
                        mark_mail_as_sent(contactjson["uuid"], context="contacts")
                        sormasapi.update_task(taskjson["uuid"], commentprefix="✅")
                        shutil.copy(filename,os.path.join(self.outputfolder,"contacts",filename.split("\\")[-1]))
                elif taskjson["taskStatus"]=="NOT_EXECUTABLE":
                    text = taskjson.get("creatorComment","")+"\n"+taskjson.get("assigneeReply","")
                    text += "\nQ: "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineFrom",""))+" bis "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineTo",""))
                    answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Nicht-erreicht-Mail schicken?\n"+text)
                    if answer:
                        filename = self.create_docx(contactjson, personjson, "KP_ne", suffix="_ne", context="contacts")
                        self.send_contact_mail(contactjson, personjson, mail, filename)
                        mark_mail_as_sent(contactjson["uuid"], note='"Nicht erreicht"-Anschreiben am '+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n", context="contacts")
                        sormasapi.update_task(taskjson["uuid"], commentprefix="✉")
                        shutil.copy(filename,os.path.join(self.outputfolder,"contacts",filename.split("\\")[-1]))
                
    def send_one_mail(self, caseuuid, taskjson):
        casejson = sormasapi.query("cases", caseuuid)
        personjson = sormasapi.query("persons",casejson["person"]["uuid"])
        mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
        self.mail_window(casejson, personjson, taskjson)
        self.list_win.wait_window(self.mail_win)
        print(self.mail_answer)
        if self.mail_answer == 1:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            self.send_standard_mail(casejson, personjson, mail, filename)
            mark_mail_as_sent(caseuuid)
            mark_task_as_completed(taskjson["uuid"])
            #TODO abgeschlossener Q-Zeitraum?
            shutil.copy(filename,os.path.join(self.outputfolder,"cases",filename.split("\\")[-1]))
        elif self.mail_answer == 2:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            self.send_mail_no_contacts(casejson, personjson, mail, filename)
            mark_mail_as_sent(caseuuid)
            mark_task_as_completed(taskjson["uuid"])
            shutil.copy(filename,os.path.join(self.outputfolder,"cases",filename.split("\\")[-1]))
        elif self.mail_answer == 3:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            os.startfile(filename)
            mark_mail_as_sent(caseuuid, note="Bescheid per Brief am "+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n")
            mark_task_as_completed(taskjson["uuid"])
            shutil.copy(filename,os.path.join(self.outputfolder,"cases",filename.split("\\")[-1]))

    def send_standard_mail(self, c, p, mail, filename):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf"),os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        if not is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_Indices_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_Indices.txt"),encoding="utf-8") as f:
                body_template = f.read()
        Anrede="Sehr geehrte*r"
        if p.get("sex")=="MALE":
            Anrede="Sehr geehrter Herr"
        if p.get("sex")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        shutil.copy(os.path.join("Vorlagen","Kontaktpersonen_Nachname_Vorname.xlsx"),os.path.join(self.outputfolder,"pdf"))
        if not os.path.exists(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastName"]+"_"+p["firstName"]+".xlsx")):
            os.rename(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_Nachname_Vorname.xlsx"), os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastName"]+"_"+p["firstName"]+".xlsx"))
        attachment_path_list.append(filename)
        attachment_path_list.append(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastName"]+"_"+p["firstName"]+".xlsx"))
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Kontakte_ab}",sormasapi.timestamp_to_datestring(c["quarantineFrom"]-172800000)).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(mail,subject,body,attachment_path_list)  

    def send_mail_no_contacts(self, c, p, mail, filename):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        if not is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_Indices_ohne_KP_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_Indices_ohne_KP.txt"),encoding="utf-8") as f:
                body_template = f.read()
        Anrede="Sehr geehrte*r"
        if p.get("sex")=="MALE":
            Anrede="Sehr geehrter Herr"
        if p.get("sex")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        attachment_path_list.append(filename)
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Kontakte_ab}",sormasapi.timestamp_to_datestring(c["quarantineFrom"]-172800000)).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(mail,subject,body,attachment_path_list)  

    def send_contact_mail(self, c, p, mail, filename): #TODO refine
        subject="Isolationsschreiben"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf")]
        if not is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_KP_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_KP.txt"),encoding="utf-8") as f:
                body_template = f.read()
        Anrede="Sehr geehrte*r"
        if p.get("sex")=="MALE":
            Anrede="Sehr geehrter Herr"
        if p.get("sex")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        aktenzeichen=c.get("externalToken","")
        if aktenzeichen=="":
            aktenzeichen=c["uuid"].split("-")[0]
        attachment_path_list.append(filename)
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(mail,subject,body,attachment_path_list)
 
    def forward_index(self, filename): #TODO: Automatische Weiterleitung per Mail/Fax an das zuständige Amt
        pass
 
    def create_docx(self, c, p, docx_file, suffix="", context="cases"):
        d={}
        if not os.path.exists(os.path.join(self.outputfolder,"pdf")):
            os.makedirs(os.path.join(self.outputfolder,"pdf"))
        Anrede = "Sehr geehrte*r"
        if p.get("sex","")=="MALE":
            Anrede="Sehr geehrter Herr"
        elif p.get("sex","")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        d["Anrede"]=Anrede
        d["Sachbearbeiter"]=self.Sachbearbeiter
        d["externaltoken"]=c.get("externalToken",c["uuid"][:6])
        d["firstname"]=p["firstName"]
        d["lastname"]=p["lastName"]
        d["street"]=p["address"].get("street","")
        d["housenumber"]=p["address"].get("houseNumber","")
        d["postalcode"]=p["address"].get("postalCode","")
        d["quarantinefrom"]=sormasapi.timestamp_to_datestring(c.get("quarantineFrom",""))
        d["quarantineto"]=sormasapi.timestamp_to_datestring(c.get("quarantineTo",""))
        if context == "contacts" and not d["quarantinefrom"]:
            d["quarantinefrom"]==sormasapi.timestamp_to_datestring(c.get("lastContactDate",""))
        if context == "contacts" and not d["quarantineto"]:
            try:
                d["quarantineto"]==sormasapi.timestamp_to_datestring(c["lastContactDate"]+1209600000)
            except:
                pass
        d["PCRSatz"]=""
        PCR_date = sormasapi.get_earliest_positive_PCR(c["uuid"])
        if context=="cases":
            PCR_date = sormasapi.get_earliest_positive_PCR(c["uuid"])
            if PCR_date:
                d["PCRSatz"]=" Der Nachweis darüber erfolgte mit Abstrich vom "+sormasapi.timestamp_to_datestring(PCR_date)+" mittels PCR-Test."
        if not is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            if os.path.exists(os.path.join(self.inputdocxpath,docx_file+"_u18.docx")):
                fill_file(os.path.join(self.inputdocxpath,docx_file+"_u18.docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+d["externaltoken"].split("-")[-1]+suffix+".docx"),d)
            else:
                fill_file(os.path.join(self.inputdocxpath,docx_file+".docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+d["externaltoken"].split("-")[-1]+suffix+".docx"),d)
        else:
            fill_file(os.path.join(self.inputdocxpath,docx_file+".docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+d["externaltoken"].split("-")[-1]+suffix+".docx"),d)
        convert(os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+d["externaltoken"].split("-")[-1]+suffix+".docx"))
        return os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+d["externaltoken"].split("-")[-1]+suffix+".pdf")
    
    def select_all(self):
        self.lb.select_set(0, tk.END)
        
    def index_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Indexpersonen")
        self.tasks=tasks_by_date(self.startint, self.endint)
        self.index_tasks=[t for t in self.tasks 
            if t["taskStatus"]=="DONE" and t.get("creatorComment","")[:1] not in "✅❌" and t["taskType"]=="CASE_INVESTIGATION"]+[t for t in self.tasks if t["taskStatus"]=="NOT_EXECUTABLE" and t.get("creatorComment","")[:1] not in "✅❌✉" and t["taskType"]=="CASE_INVESTIGATION"]
        self.lb=tk.Listbox(self.list_win, height=min(len(self.index_tasks),28), width=30, font=font, selectmode="multiple")
        for i,t in enumerate(self.index_tasks):
            casestr = t.get("creatorComment","")[:4]+"\t"+t["caze"]["caption"]
            self.lb.insert(i,casestr)
            if t["taskStatus"]=="NOT_EXECUTABLE":
                self.lb.itemconfig(i, {'bg':'light pink'})
        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.mail_indices, font=self.buttonfont)
        b1.grid(row=1, column=1, sticky=tk.N)
        b2 = tk.Button(self.list_win, text="Alle auswählen", command=self.select_all, font=self.buttonfont)
        b2.grid(row=1, column=0, sticky=tk.N)
    
    def contact_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Kontaktpersonen")
        self.tasks=tasks_by_date(self.startint, self.endint)
        self.contact_tasks=[t for t in self.tasks if t["taskStatus"]=="DONE" and t["taskType"]=="CONTACT_INVESTIGATION"]+[t for t in self.tasks if t["taskStatus"]=="NOT_EXECUTABLE" and t["taskType"]=="CONTACT_INVESTIGATION"]
        self.lb=tk.Listbox(self.list_win, height=min(len(self.contact_tasks),28), width=30, font=font, selectmode="multiple")
        for i,t in enumerate(self.contact_tasks):
            casestr = t.get("creatorComment","")[:4]+"\t"+t["contact"]["caption"]
            self.lb.insert(i,casestr)
            if t["taskStatus"]=="NOT_EXECUTABLE":
                self.lb.itemconfig(i, {'bg':'light pink'})
        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.mail_contacts, font=self.buttonfont)
        b1.grid(row=1, column=1, sticky=tk.N)
        b2 = tk.Button(self.list_win, text="Alle auswählen", command=self.select_all, font=self.buttonfont)
        b2.grid(row=1, column=0, sticky=tk.N)
        
    def mail_window(self, casejson, personjson, taskjson):
        def return_value(value):
            self.mail_answer = value
            self.mail_win.destroy()
        self.mail_win = tk.Toplevel(self, bg = "white")
        text=""
        if casejson.get("quarantineOrderedOfficialDocument","")==True:
            text+="⚠ Anordnung wurde bereits am "+sormasapi.timestamp_to_datestring(casejson.get("quarantineOrderedOfficialDocumentDate",""))+" als verschickt markiert, bitte prüfen ob Bescheid erneut versendet werden soll.\n"
        if casejson["healthFacility"]["uuid"]!="SORMAS-CONSTID-ISNONE-FACILITY":
            text+="🚑 Quarantäneort als 'institutionell' markiert ("+casejson["healthFacility"]["caption"]+"), prüfen ob Brief/Mail verschickt werden muss.\n"
        if casejson["caseClassification"]=="NO_CASE":
            text+="❗ Akte wurde als 'KEIN FALL' markiert\n"
        if not casejson.get("quarantineTo","") or not casejson.get("quarantineFrom",""):
            text+="⚠ Quarantänezeitraum unvollständig\n"
            
        #if d["street"]=="" or d["postalcode"]=="":
        #    self.note(c_id, "⚠ Adresse unvollständig")

        self.mail_win.title(personjson["firstName"]+" "+personjson["lastName"]+" "+casejson["uuid"])
        text+= "\n Quarantänezeitraum: "+sormasapi.timestamp_to_datestring(casejson.get("quarantineFrom",""))+" bis "+sormasapi.timestamp_to_datestring(casejson.get("quarantineTo",""))+"\n"
        text+= "Symptombeginn: "+sormasapi.timestamp_to_datestring(casejson["symptoms"].get("onsetDate",""))+"\n"
        text+= casejson.get("additionalDetails","")
        label = tk.Label(self.mail_win, text=text, bg = "white", font=font, wraplength=1000).grid(row=0, column=0, columnspan=4, sticky=tk.N)
        b1 = tk.Button(self.mail_win, text="Mail + KP", command = lambda: return_value(1), font=self.buttonfont).grid(row=1, column=0, sticky=tk.N)
        b2 = tk.Button(self.mail_win, text="Nur Mail", command = lambda: return_value(2), font=self.buttonfont).grid(row=1, column=1, sticky=tk.N)
        b3 = tk.Button(self.mail_win, text="Brief", command = lambda: return_value(3), font=self.buttonfont).grid(row=1, column=2, sticky=tk.N)
        b4 = tk.Button(self.mail_win, text="Nichts", command =  lambda : return_value(0), font=self.buttonfont).grid(row=1, column=3, sticky=tk.N)
  
    def choose_startdate(self):
        self.startint = sormasapi.datestring_to_int(sd.askstring("Datum", "Ermittlungsdatum eingeben (TT.MM.JJJJ HH:MM)", parent=self.master))
        self.input_startdatetime["bg"]=self.bc2
    
    def choose_enddate(self):
        self.endint = sormasapi.datestring_to_int(sd.askstring("Datum", "Ermittlungsdatum eingeben (TT.MM.JJJJ HH:MM)", parent=self.master))
        self.input_enddatetime["bg"]=self.bc2
        
    def choose_sachbearbeiter(self):
        self.Sachbearbeiter = sd.askstring("Sachbearbeiter*in", "Stellenzeichen eingeben", parent=self.master)
        self.input_sachbearbeiter["bg"]=self.bc2
        self.input_sachbearbeiter["text"]="3. Sachbearbeiter*in\n"+self.Sachbearbeiter
        print("Sachbearbeiter: "+self.Sachbearbeiter)
        
if __name__ == "__main__":
    if config.useproxy:
        os.environ['HTTPS_PROXY'] = config.proxy
    root = tk.Tk()
    app = Application(master=root)
    icon = tk.PhotoImage(file = os.path.join("Vorlagen","icon.png"))
    root.iconphoto(False, icon)
    global font
    font=('Berlin Type Office Regular',14)
    app.mainloop()
