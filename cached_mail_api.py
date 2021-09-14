import sormasapi
import config

import csv
import datetime
import sys
import os
import re
import shutil
import string
import unicodedata
from pprint import pprint

from mailmerge import MailMerge #Externe Library für docx-Serienbriefe
from docx2pdf import convert    #Externe Library für pdf-Erstellung mit Word
import win32com.client          #Externe Library für Email-Versand per Outlook

import tkinter as tk            #Graphische Oberfläche
import tkinter.filedialog as fd
import tkinter.simpledialog as sd

def fill_file(input_docx_path,output_docx_path,fielddict):
    """Füllt die Serienbrief-Vorlage aus. Benutzte Keywords müssen denen in der Vorlage entsprechen."""
    with MailMerge(input_docx_path) as document:
        document.merge_templates([fielddict],separator="page_break")
        document.write(output_docx_path)

def copy(filename, context="cases"):
    """Kopiert versendete Anschreiben an den richtigen Ort im Gemeinschaftslaufwerk. Spezifisch für GA FK"""
    if context=="cases":
        if unicodedata.normalize("NFKD",filename.split("\\")[-1][0])[0] in string.ascii_letters:
            dest_path=os.path.join(config.indexpath,unicodedata.normalize("NFKD",filename.split("\\")[-1][0])[0].upper(),filename.split("\\")[-1])
        else:
            dest_path=os.path.join(config.indexpath,filename.split("\\")[-1])
    elif context=="contacts":
        if unicodedata.normalize("NFKD",filename.split("\\")[-1][0])[0] in string.ascii_letters:
            dest_path=os.path.join(config.contactpath,unicodedata.normalize("NFKD",filename.split("\\")[-1][0])[0].upper(),filename.split("\\")[-1])
        else:
            dest_path=os.path.join(config.contactpath,filename.split("\\")[-1])
    i=2
    while os.path.exists(dest_path):
        print("Datei bereits vorhanden: "+dest_path)
        dest_path=dest_path.rsplit(".",1)[0]+"_"+str(i)+"."+dest_path.rsplit(".",1)[1]
        i+=1
    print(filename)
    print(dest_path)
    shutil.copy(filename,dest_path)   

def mark_mail_as_sent(uuid, note="", context="cases"):
    """Markiert Bescheid in SORMAS als versendet - optionaler Kommentar für das Follow-Up-Feld"""
    changedict={
        "quarantineOrderedOfficialDocument" : True,
        "quarantineOrderedOfficialDocumentDate" : sormasapi.now()
    }
    if context=="cases":
        sormasapi.update_case(uuid, changedict, note)
    elif context=="contacts":
        sormasapi.update_contact(uuid, changedict, note)

def mark_task_as_completed(taskuuid):
    """Markiert eine Aufgabe als komplett erledigt. Wir spiegeln das durch ein ✅ am Beginn des Aufgabenkommentars wieder."""
    sormasapi.update_task(taskuuid, commentprefix="✅")

def send_mail(email,subject,body,attachment_path_list=[]):
    """Generiert eine Email in Outlook, optional mit Anhängen als Liste aus Dateipfaden."""
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

def get_mail(PLZ):
    """Retrieve mail address of responsible health department by postal code / (place name). Not yet in use."""
    URL = "https://tools.rki.de/PLZTool/?q=" + str(PLZ)
    page = requests.get(URL)
    if 'title="E-Mail"' in page.text:
        return page.text.split('href="mailto:')[1].split(">@</a>")[0]
    else:
        return ""
    
def get_fax(PLZ):
    """Retrieve fax address of responsible health department by postal code / (place name). Not yet in use."""
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
        self.startint=sormasapi.now()-604800000
        self.endint=sormasapi.now()
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
        if not os.path.exists(os.path.join(self.outputfolder,"pdf")):
            os.makedirs(os.path.join(self.outputfolder,"pdf"))
        self.tasks = []
        self.index_tasks = []
        self.contact_tasks = []
        self.ersatzschein_tasks = []
        self.gb_list = []
        self.inputdocxpath=os.path.join("Vorlagen")
        
    def create_widgets(self):
        self.run = tk.Button(self, text="Initialisieren/Aktualisieren", command=self.refresh, anchor="w",  font=self.buttonfont, width=self.bw, bg='white')
        self.run.pack(side="top")
        self.run = tk.Button(self, text="Index-Mails", command=self.index_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='light grey')
        self.run.pack(side="top")
        self.run_contacts = tk.Button(self, text="KP-Mails", command=self.contact_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='light grey')
        self.run_contacts.pack(side="top")
        self.run_ersatzscheine = tk.Button(self, text="Ersatzscheine", command=self.ersatzschein_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='light grey')
        self.run_ersatzscheine.pack(side="top")
        self.run_gb = tk.Button(self, text="Genesenenbescheide", command=self.gb_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='light grey')
        self.run_gb.pack(side="top")
        self.run_te = tk.Button(self, text="Test-Erinnerungs-Mails", command=self.te_listbox, anchor="w",  font=self.buttonfont, width=self.bw, bg='light grey')
        self.run_te.pack(side="top")
        self.quit = tk.Button(self, text=" BEENDEN", command=self.master.destroy, anchor="w", font=self.buttonfont, width=self.bw, bg='misty rose')
        self.quit.pack(side="bottom")
    
    def refresh(self):
        """Loads info from cache, loads new info from database, refreshes cache"""
        self.endint = sormasapi.now()
        self.cases = sormasapi.load_and_update_json(os.path.join(config.cachepath,"cases.json"))
        print(str(len(self.cases))+" cases loaded")
        self.persons = sormasapi.load_and_update_json(os.path.join(config.cachepath,"persons.json"),"persons")
        print(str(len(self.persons))+" persons loaded")
        self.tasks = sormasapi.load_and_update_json(os.path.join(config.cachepath,"tasks.json"),"tasks")
        print(str(len(self.tasks))+" tasks loaded")
        self.run.configure(bg = "white")
        self.run_contacts.configure(bg = "white")
        self.run_ersatzscheine.configure(bg = "white")
        self.run_gb.configure(bg = "white")
        self.run_te.configure(bg = "white")
    
    def send_all_gb(self):
        for i in self.lb.curselection():
            casejson = self.gb_list[i]
            pprint(casejson)
            p = sormasapi.query("persons",casejson["person"]["uuid"])
            mail = ([""]+[x.get("contactInformation","") for x in p.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
            text = "Mail: "+mail
            text += "\n"+casejson.get("followUpComment","")
            answer = tk.messagebox.askyesno(p["firstName"]+" "+p["lastName"],"Mail schicken?\n"+text)
            if answer:
                filename = self.create_docx(casejson, p, "GB", suffix="_gb", context="cases")
                subject = "Bestätigung über eine zurückliegende SARS-CoV-2-Infektion"
                Anrede="Sehr geehrte*r"
                if p.get("sex")=="MALE":
                    Anrede="Sehr geehrter Herr"
                if p.get("sex")=="FEMALE":
                    Anrede="Sehr geehrte Frau"
                if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
                    Anrede="Sehr geehrte gesetzliche Vertreter:innen von"
                body_template=""          
                with open(os.path.join("Vorlagen","Email_GB.html"),encoding="utf-8") as f:
                    body_template = f.read()
                body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Sachbearbeiter}",config.stellenzeichen)
                send_mail(mail,subject,body,[filename])
                copy(filename,"cases")
                sormasapi.update_case(casejson["uuid"], followupprefix=sormasapi.timestamp_to_datestring(sormasapi.now())+": 💌Genesenenbescheid verschickt - "+config.stellenzeichen+"\n")          
    
    def send_all_te(self):
        for i in self.lb.curselection():
            casejson = self.te_list[i]
            pprint(casejson)
            p = sormasapi.query("persons",casejson["person"]["uuid"])
            mail = ([""]+[x.get("contactInformation","") for x in p.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
            text = "Mail: "+mail
            text += "\n"+casejson.get("followUpComment","")
            answer = tk.messagebox.askyesno(p["firstName"]+" "+p["lastName"],"Mail schicken?\n"+text)
            if answer:
                subject = "Hinweise zum Quarantäneende"
                Anrede="Sehr geehrte*r"
                Sie_Kind = "Sie"
                if p.get("sex")=="MALE":
                    Anrede="Sehr geehrter Herr"
                if p.get("sex")=="FEMALE":
                    Anrede="Sehr geehrte Frau"
                if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
                    Anrede="Sehr geehrte gesetzliche Vertreter:innen von"
                    Sie_Kind = "Ihr Kind"
                body_template=""          
                with open(os.path.join("Vorlagen","Email_Testerinnerung.html"),encoding="utf-8") as f:
                    body_template = f.read()
                body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Sachbearbeiter}",config.stellenzeichen).replace("{Sie_Kind}",Sie_Kind).replace("{Q_Ende}",sormasapi.timestamp_to_datestring(casejson.get("quarantineTo",0)))
                send_mail(mail,subject,body)
                sormasapi.update_case(casejson["uuid"], followupprefix=sormasapi.timestamp_to_datestring(sormasapi.now())+": ⏰Testerinnerung verschickt - "+config.stellenzeichen+"\n")
             
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
                    mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
                    if mail:
                        send_mail(mail,"Positiver Covid-19 Test - bitte um Rückmeldung","Siehe Anhang",[filename])
                        #TODO mail_body etc.
                    else:
                        os.startfile(filename)
                    copy(filename,"cases")
                    mark_mail_as_sent(casejson["uuid"], note='✉"Nicht erreicht"-Anschreiben am '+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n")
                    sormasapi.update_task(taskjson["uuid"], commentprefix="✉")
                    
    def mail_contacts(self):
        for i in self.lb.curselection():
            contactjson = sormasapi.query("contacts", self.contact_tasks[i]["contact"]["uuid"])
            pprint(contactjson)
            personjson = sormasapi.query("persons",contactjson["person"]["uuid"])
            pprint(personjson)
            taskjson=self.contact_tasks[i]
            if contactjson.get("contactClassification","") in ["CONFIRMED","UNCONFIRMED"] and not contactjson.get("quarantineOrderedOfficialDocument",False):
                mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
                if taskjson["taskStatus"]=="DONE":
                    text = "Mail: "+mail
                    text += "\n"+taskjson.get("creatorComment","")+"\n"+taskjson.get("assigneeReply","")
                    text += "\nQ: "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineFrom",""))+" bis "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineTo",""))
                    answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Mail schicken?\n"+text)
                    if answer:
                        filename = self.create_docx(contactjson, personjson, "Isolationsbescheinigung", suffix="", context="contacts")
                        self.send_contact_mail(contactjson, personjson, mail, filename)
                        mark_mail_as_sent(contactjson["uuid"], context="contacts")
                        sormasapi.update_task(taskjson["uuid"], commentprefix="✅")
                        copy(filename,"contacts")
                elif taskjson["taskStatus"]=="NOT_EXECUTABLE":
                    text = taskjson.get("creatorComment","")+"\n"+taskjson.get("assigneeReply","")
                    text += "\nQ: "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineFrom",""))+" bis "+sormasapi.timestamp_to_datestring(contactjson.get("quarantineTo",""))
                    answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Nicht-erreicht-Mail schicken?\n"+text)
                    if answer:
                        filename = self.create_docx(contactjson, personjson, "KP_ne", suffix="_ne", context="contacts")
                        self.send_contact_mail(contactjson, personjson, mail, filename)
                        mark_mail_as_sent(contactjson["uuid"], note='✉"Nicht erreicht"-Anschreiben am '+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n", context="contacts")
                        sormasapi.update_task(taskjson["uuid"], commentprefix="✉")
                        copy(filename,"contacts")
                        
    def mail_ersatzscheine(self):
        for i in self.lb.curselection():
            taskjson=self.ersatzschein_tasks[i]
            personjson={}
            if taskjson["taskContext"]=="CASE":
                context="cases"
                cjson = sormasapi.query("cases", taskjson["caze"]["uuid"])
                personjson = sormasapi.query("persons", cjson["person"]["uuid"])
                pprint(personjson)
            if taskjson["taskContext"]=="CONTACT":
                context="contacts"
                cjson = sormasapi.query("contacts",taskjson["contact"]["uuid"])
                personjson = sormasapi.query("persons", cjson["person"]["uuid"])
                pprint(personjson)
            mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
            answer = tk.messagebox.askyesno(personjson["firstName"]+" "+personjson["lastName"],"Ersatzschein schicken?")
            if answer:
                filename = self.create_docx({"quarantineTo":taskjson["dueDate"]}, personjson, "Ersatzschein", suffix="_ersatzschein", context=context)
                self.send_ersatzschein_mail(personjson, mail, filename,sormasapi.timestamp_to_datestring(taskjson["dueDate"]))
                sormasapi.update_task(taskjson["uuid"], changedict = {"taskStatus":"DONE"}, commentprefix="✅")
                copy(filename,context)
                
    def send_one_mail(self, caseuuid, taskjson):
        casejson = sormasapi.query("cases", caseuuid)
        personjson = sormasapi.query("persons",casejson["person"]["uuid"])
        mail = ([""]+[x.get("contactInformation","") for x in personjson.get("personContactDetails",[]) if "@" in x.get("contactInformation")]).pop()
        self.mail_window(casejson, personjson, taskjson)
        self.list_win.wait_window(self.mail_win)
        print(self.mail_answer)
        #TODO continue here
        if self.mail_answer == 1:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            self.send_standard_mail(casejson, personjson, mail, filename)
            mark_mail_as_sent(caseuuid)
            mark_task_as_completed(taskjson["uuid"])
            #TODO abgeschlossener Q-Zeitraum?
            copy(filename,"cases")
        elif self.mail_answer == 2:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            self.send_mail_no_contacts(casejson, personjson, mail, filename)
            mark_mail_as_sent(caseuuid)
            mark_task_as_completed(taskjson["uuid"])
            copy(filename,"cases")
        elif self.mail_answer == 3:
            filename = self.create_docx(casejson, personjson, "Anschreiben Indices")
            os.startfile(filename)
            mark_mail_as_sent(caseuuid, note="Bescheid per Brief am "+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt\n")
            mark_task_as_completed(taskjson["uuid"])
            copy(filename,"cases")

    def send_standard_mail(self, c, p, mail, filename):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf"),os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_Indices_u18.html"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_Indices.html"),encoding="utf-8") as f:
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
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Kontakte_ab}",sormasapi.timestamp_to_datestring(c["quarantineFrom"]-172800000)).replace("{Sachbearbeiter}",config.stellenzeichen)
        send_mail(mail,subject,body,attachment_path_list)  

    def send_mail_no_contacts(self, c, p, mail, filename):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_Indices_ohne_KP_u18.html"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_Indices_ohne_KP.html"),encoding="utf-8") as f:
                body_template = f.read()
        Anrede="Sehr geehrte*r"
        if p.get("sex")=="MALE":
            Anrede="Sehr geehrter Herr"
        if p.get("sex")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        attachment_path_list.append(filename)
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Kontakte_ab}",sormasapi.timestamp_to_datestring(c["quarantineFrom"]-172800000)).replace("{Sachbearbeiter}",config.stellenzeichen)
        send_mail(mail,subject,body,attachment_path_list)  

    def send_contact_mail(self, c, p, mail, filename):
        subject="Isolationsschreiben"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf")]
        if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            with open(os.path.join("Vorlagen","Email_KP_u18.html"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_KP.html"),encoding="utf-8") as f:
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
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Sachbearbeiter}",config.stellenzeichen)
        send_mail(mail,subject,body,attachment_path_list)
        
    def send_ersatzschein_mail(self, p, mail, filename, datum):
        subject="Ersatzschein für Ihre PCR-Testung"
        attachment_path_list=[filename]
        Anrede="Sehr geehrte*r"
        body_template=""
        if p.get("sex")=="MALE":
            Anrede="Sehr geehrter Herr"
        if p.get("sex")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        sich_ihr_Kind = "sich"
        if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
            sich_ihr_Kind = "Ihr Kind"
            Anrede="Sehr geehrte gesetzliche Vertreter:innen von"
        with open(os.path.join("Vorlagen","Ersatzschein.html"),encoding="utf-8") as f:
            body_template = f.read()
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstName"]).replace("{Nachname}",p["lastName"]).replace("{Sachbearbeiter}",config.stellenzeichen).replace("{sich_ihr_Kind}",sich_ihr_Kind).replace("{Datum}",datum)
        send_mail(mail,subject,body,attachment_path_list)
        
    def forward_index(self, filename): #TODO
        pass
 
    def create_docx(self, c, p, docx_file, suffix="", context="cases"):
        d={}
        if not os.path.exists(os.path.join(self.outputfolder,"pdf")):
            os.makedirs(os.path.join(self.outputfolder,"pdf"))
        Anrede = "Sehr geehrte/r"
        if p.get("sex","")=="MALE":
            Anrede="Sehr geehrter Herr"
        elif p.get("sex","")=="FEMALE":
            Anrede="Sehr geehrte Frau"
        d["Anrede"]=Anrede
        d["Sachbearbeiter"]=config.stellenzeichen
        d["externaltoken"]=c.get("externalToken",c.get("uuid","")[:6])
        d["firstname"]=p["firstName"]
        d["lastname"]=p["lastName"]
        d["street"]=p["address"].get("street","")
        d["housenumber"]=p["address"].get("houseNumber","")
        d["postalcode"]=p["address"].get("postalCode","")
        d["city"]=p["address"].get("city","Berlin")
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
        PCR_date=""
        if suffix!="_ersatzschein":
            PCR_date = sormasapi.get_earliest_positive_PCR(c["uuid"])
        if context=="contacts":
            d["reason1"]="☒"
            d["reason2"]="☐"
        if context=="cases":
            d["reason1"]="☐"
            d["reason2"]="☒"
            if suffix!="_ersatzschein":
                PCR_date = sormasapi.get_earliest_positive_PCR(c["uuid"])
            d["PCR_date"] = ""
            if PCR_date:
                d["PCRSatz"]=" Der Nachweis darüber erfolgte mit Abstrich vom "+sormasapi.timestamp_to_datestring(PCR_date)+" mittels PCR-Test."
                d["PCR_date"]= sormasapi.timestamp_to_datestring(PCR_date)
        d["geborenSatz"]=""
        if p.get("birthdateDD","") and p.get("birthdateMM","") and p.get("birthdateYYYY",""):
            d["geborenSatz"]=", geboren am "+str(p.get("birthdateDD",""))+"."+str(p.get("birthdateMM","")).zfill(2)+"."+str(p.get("birthdateYYYY","")).zfill(2)+" "
        if not sormasapi.is_adult(str(p.get("birthdateYYYY",""))+"-"+str(p.get("birthdateMM","")).zfill(2)+"-"+str(p.get("birthdateDD","")).zfill(2)):
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
 
    def gb_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Genesenenbescheid")
        self.gb_list = []
        for case in self.cases.values():
            #if case.get("quarantineTo",0) < self.endint and case.get("quarantineTo",0) >= self.startint and not "💌" in case.get("followUpComment",""):
            if case.get("quarantineTo",0) < sormasapi.now()-432000000 and case.get("quarantineTo",0) >= 1630533600000 and not "💌" in case.get("followUpComment",""):
                if sormasapi.get_earliest_positive_PCR(case["uuid"]) and case["caseClassification"]!="NO_CASE": #and case.get("followUpStatus","FOLLOW_UP") != "FOLLOW_UP"
                    self.gb_list.append(case)
        self.lb = tk.Listbox(self.list_win, height=min(len(self.gb_list),28), width=30, font=font, selectmode="multiple")

        for i,casejson in enumerate(self.gb_list):
            casestr = casejson.get("person",{}).get("caption","")
            self.lb.insert(i,casestr)

        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.send_all_gb, font=self.buttonfont)
        b1.grid(row=1, column=1, sticky=tk.N)
        b2 = tk.Button(self.list_win, text="Alle auswählen", command=self.select_all, font=self.buttonfont)
        b2.grid(row=1, column=0, sticky=tk.N)

    def te_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Testerinnerungen")
        self.te_list = []
        for casejson in self.cases.values():
            #if casejson.get("quarantineTo",0) < self.endint and casejson.get("quarantineTo",0) >= self.startint and not "⏰" in case.get("followUpComment",""):
            if casejson.get("quarantineTo",0) < sormasapi.now()+259200000 and casejson.get("quarantineTo",0) >= 1631484000000 and not "⏰" in casejson.get("followUpComment",""):
                if casejson.get("followUpStatus","FOLLOW_UP") == "FOLLOW_UP" and casejson["caseClassification"]!="NO_CASE": #if sormasapi.get_earliest_positive_PCR(casejson["uuid"]) entfernt
                    if not (casejson.get("outcomeDate",0)==946681200000 and sormasapi.is_fully_vaccinated(casejson)):
                        self.te_list.append(casejson)
        self.lb = tk.Listbox(self.list_win, height=min(len(self.te_list),28), width=30, font=font, selectmode="multiple")
        for i,casejson in enumerate(self.te_list):
            casestr = casejson.get("person",{}).get("caption","")
            self.lb.insert(i,casestr)

        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.send_all_te, font=self.buttonfont)
        b1.grid(row=1, column=1, sticky=tk.N)
        b2 = tk.Button(self.list_win, text="Alle auswählen", command=self.select_all, font=self.buttonfont)
        b2.grid(row=1, column=0, sticky=tk.N)
        
 
    def index_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Indexpersonen")
        self.index_tasks=self.tasks_by_date(self.startint, self.endint)
        self.index_tasks=[t for t in self.index_tasks if t["taskStatus"]=="DONE" and t.get("creatorComment"," ")[:1] not in "✅❌" and t["taskType"]=="CASE_INVESTIGATION"]+[t for t in self.index_tasks if t["taskStatus"]=="NOT_EXECUTABLE" and t.get("creatorComment"," ")[:1] not in "✅❌✉" and t["taskType"]=="CASE_INVESTIGATION"]
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
        self.contact_tasks=self.tasks_by_date(self.startint, self.endint)
        self.contact_tasks=[t for t in self.contact_tasks if t["taskStatus"]=="DONE" and t["taskType"]=="CONTACT_INVESTIGATION"]+[t for t in self.contact_tasks if t["taskStatus"]=="NOT_EXECUTABLE" and t["taskType"]=="CONTACT_INVESTIGATION"]
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
        
    def ersatzschein_listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Ersatzscheine")
        self.ersatzschein_tasks=self.tasks_by_date(self.startint, self.endint)
        self.ersatzschein_tasks=[t for t in self.ersatzschein_tasks if t["assigneeUser"]["uuid"].startswith("U5HB2X") and t["taskStatus"]=="PENDING"]
        self.lb=tk.Listbox(self.list_win, height=min(len(self.ersatzschein_tasks),28), width=30, font=font, selectmode="multiple")
        for i,t in enumerate(self.ersatzschein_tasks):
            if "contact" in t:
                casestr = t["contact"]["caption"]
            elif "caze" in t:
                casestr = t["caze"]["caption"]
            self.lb.insert(i,casestr)
        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.mail_ersatzscheine, font=self.buttonfont)
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
        text+= casejson.get("additionalDetails","")+"\n"
        text+= "🗨"+taskjson.get("creatorComment","")+"\n"
        text+= "🗨"+taskjson.get("assigneeReply","")
        label = tk.Label(self.mail_win, text=text, bg = "white", font=font, wraplength=1000).grid(row=0, column=0, columnspan=4, sticky=tk.N)
        b1 = tk.Button(self.mail_win, text="Mail + KP", command = lambda: return_value(1), font=self.buttonfont).grid(row=1, column=0, sticky=tk.N)
        b2 = tk.Button(self.mail_win, text="Nur Mail", command = lambda: return_value(2), font=self.buttonfont).grid(row=1, column=1, sticky=tk.N)
        b3 = tk.Button(self.mail_win, text="Brief", command = lambda: return_value(3), font=self.buttonfont).grid(row=1, column=2, sticky=tk.N)
        b4 = tk.Button(self.mail_win, text="Nichts", command =  lambda : return_value(0), font=self.buttonfont).grid(row=1, column=3, sticky=tk.N)

    def tasks_by_date(self, startint, endint=sormasapi.now()):
        """Gibt alle in einem bestimmten Zeitraum als geändert markierten Aufgaben zurück"""
        return [t for t in self.tasks.values() if t["statusChangeDate"]<=endint]
        #return [t for t in sormasapi.get_since("tasks", dateint=startint) if t["statusChangeDate"]<=endint] #API

if __name__ == "__main__":
    os.environ['HTTPS_PROXY'] = config.proxy
    root = tk.Tk()
    app = Application(master=root)
    icon = tk.PhotoImage(file = os.path.join("Vorlagen","icon.png"))
    root.iconphoto(False, icon)
    global font
    font=('Berlin Type Office Regular',14)
    app.mainloop()
