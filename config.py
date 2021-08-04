# Intended use of these scripts: Trying out API functionalities on a local SORMAS instance populated with test data
# Take care when handling passwords to productive instances
sormasurl = "https://bundesland-landkreis.sormas.bund.de/sormas-rest/" # URL der SORMAS-Instanz - docker-test nur lokal
restuser = "restapi" # Benutzername eines Users mit Nutzerrolle "Restuser"
restpw = "xxxxxxx" # Dazugehöriges Passwort
stellenzeichen = "PaKo 07" # Wird in den erstellten Dokumenten/Mails/Bescheiden verwendet

useruuid = "XXXXXX-XXXXXX-XXXXXX-XXXXXXXX" # uuid of user script uses to create things
verify = True # SSL certification (requests module) - False iff local docker instance

# Für Kontaktimporte - mit Standardwerten befüllte Akten
person_template_uuid = "XXXXXX-XXXXXX-XXXXXX-XXXXXXXX"       # UUID von mit Standardwerten befüllter Person
contact_template_uuid = "XXXXXX-XXXXXX-XXXXXX-XXXXXXXX"      # UUID von mit Standard/Defaultwerten befülltem Kontakt

useproxy = False
proxy = "http://user:pw@proxyurl.de:8080"