from msg_template import GraphMessage

class GraphEmail:
    def __init__(self, firstname, recipient):
        self._firstname = firstname
        self._recipient = recipient
        self._subject = 'Welcome to EA!!!'
        self._content_type = 'HTML'
        self._content = GraphMessage(self._firstname)

    def get_payload(self):

        payload = {
            "message": {
                "subject": self._subject,
                "body": {
                    "contentType": self._content_type,
                    "content": self._content.get_content(),
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": self._recipient
                        }
                    }
                ]
            }
        }
        return payload

