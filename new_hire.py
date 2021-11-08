from graph_email import  GraphEmail

class NewHire:
    def __init__(self, fullname, startdate, location, email):
        self._fullname = fullname
        self._startdate = startdate
        self._location = location
        self._email = email
        self._graphemail = GraphEmail(self.get_firstname(), self.get_email())

    def get_firstname(self):
        namesplit = self._fullname.split()
        firstname = namesplit[0]
        return firstname

    def get_lastname(self):
        namesplit = self._fullname.split()
        lastname = namesplit[-1]
        return lastname

    def get_fullname(self):
        return self.get_firstname() + ' ' + self.get_lastname()

    def get_startdate(self):
        return self._startdate

    def get_location(self):
        return self._location

    def get_email(self):
        return self._email

    def getpayload(self):
        return self._graphemail.get_payload()