DJANGO = 'http://192.168.2.211/acme3'
TEXT_RE = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)

class email_import(Dispatching):
    def GetAllEmails(self):
        TheOneAndHolyOneComma = ','
        AllEmails = []
        AllEmails = TheOneAndHolyOneComma.join(self.get_cust_options_emails())
        return AllEmails

    def GetCSREmails(self):
        TheOneAndHolyOneComma = ','
        AllEmails = []
        AllEmails = TheOneAndHolyOneComma.join(self.get_cust_options_csrEmails())
        return AllEmails
    
    def GetCUSEmails(self):
        TheOneAndHolyOneComma = ','
        AllEmails = []
        AllEmails = TheOneAndHolyOneComma.join(self.get_cust_options_custEmails())
        return AllEmails

    def SendCompletionEmail(self, to=None, inserts=None, custtotal=None, attachments=None, inputfiles=True):
        Options = ''

        if to == "CSR":
            Options += (' -t ' + self.GetCSREmails())
        elif to == "CUS":
            Options += (' -t ' + self.GetCUSEmails())
        elif to: 
            Options += (' -t ' + to)
        elif not to:
            Options += (' -t ' + self.GetAllEmails())

        if inserts is True:
            Options += ' -i YES'

        if custtotal is True:
            Options += ' -c YES'

        if attachments:
            Options += (' -a ' + attachments)

        if inputfiles is True:
            Options += (' -f ' + 'YES')

        command = "python /opt/adf/producer/scripts/despatch_scripts/send_completion_email_dispatch.py -j " + self._dir + Options
        self.WriteStdout(command, "SendCompletionEmail")
        self.ExecuteSubProcess(command, 'sendCompletionEmail', string_fmt=True)
