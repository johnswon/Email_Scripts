def SendCompletionEmail(self):

        inputFile = open(self._dir + self._pathSep + self._job + '.bil',"r")
        stuff = ET.parse(inputFile).getroot()

        jobnum = str(stuff.find("JOBNUM").text)
        appcode = str(stuff.find("APPCODE").text)
        workorder = str(stuff.find('WOID').text)
        input_file_name = str(stuff.find('DISPFILENAME').text)

        customer =  open(self._dir + self._pathSep + self._job + '.par',"r")
        customer_par = ET.parse(customer).getroot()
        customer_name = customer_par.find("Content").find('JOB').find('CUSTOMER').text

        jobdict = {}
        jobentry = []
        total_mailpieces = 0
        total_documents  = 0
        total_pages      = 0
        total_sheets     = 0
        total_insert1    = 0
        total_insert2    = 0
        total_insert3    = 0
        total_insert4    = 0
        total_insert5    = 0
        total_insert6    = 0

        THE_SPLITTER = ""
        THE_CONVERGENCE = ""
        CompletionEmailBody = ""

        #for new email process

        for stuffGrp in stuff.getiterator('SPLIT') :

            stfGrp = str(stuffGrp.find("STUFFGROUP").text)
            splittype = str(stuffGrp.find('SPLITTYPE').text)
            for pkg in stuffGrp.getiterator('PACKAGE') :

                file_name   = str(pkg.find('FILENAME').text)
                description = str(pkg.find('DESCRIPTION').text)
                documents   = str(pkg.find('TOTALDOCUMENTS').text)
                pages       = str(pkg.find('TOTALIMPRESSIONS').text)
                sheets      = str(pkg.find('TOTALPHYSICALPAGES').text)
                HHmerges    = str(pkg.find('TOTALHHMERGES').text)
                insert1     = str(pkg.find('INSERT1CNT').text)
                insert2     = str(pkg.find('INSERT2CNT').text)
                insert3     = str(pkg.find('INSERT3CNT').text)
                insert4     = str(pkg.find('INSERT4CNT').text)
                insert5     = str(pkg.find('INSERT5CNT').text)
                insert6     = str(pkg.find('INSERT6CNT').text)

                mailpieces  = str(int(documents) - int(HHmerges))

                jobentry = []
                jobentry.append(input_file_name)
                jobentry.append(description)
                jobentry.append(mailpieces)
                jobentry.append(documents)
                jobentry.append(pages)
                jobentry.append(sheets)
                jobentry.append(splittype)

                if insert1 > '':
                    insert1count = insert1
                else:
                    insert1count = 0
                if insert2 > '':
                    insert2count = insert2
                else:
                    insert2count = 0
                if insert3 > '':
                    insert3count = insert3
                else:
                    insert3count = 0
                if insert4 > '':
                    insert4count = insert4
                else:
                    insert4count = 0
                if insert5 > '':
                    insert5count = insert5
                else:
                    insert5count = 0
                if insert6 > '':
                    insert6count = insert6
                else:
                    insert6count = 0

                jobentry.append(insert1count)
                jobentry.append(insert2count)
                jobentry.append(insert3count)
                jobentry.append(insert4count)
                jobentry.append(insert5count)
                jobentry.append(insert6count)

                jobdict[file_name] = jobentry

                if splittype == 'DOCUMENTS':
                    total_mailpieces += int(mailpieces)
                    total_documents  += int(documents)
                    total_pages      += int(pages)
                    total_sheets     += int(sheets)
                    total_insert1    += int(insert1count)
                    total_insert2    += int(insert2count)
                    total_insert3    += int(insert3count)
                    total_insert4    += int(insert4count)
                    total_insert5    += int(insert5count)
                    total_insert6    += int(insert6count)

        template = open("/opt/adf/app/common/email_template/emailtemplate", "rb")
        for line in template:
            if "<!--JOBNUM-->" in line:
                line = line.replace("<!--JOBNUM-->", jobnum)
            if "<!--APPLICATION-->" in line:
                line = line.replace("<!--APPLICATION-->", appcode)
            if "<!--INPUTFILE-->" in line:
                line = line.replace("<!--INPUTFILE-->", input_file_name.split("-")[-1])
            if "<!--WORKORDER-->" in line:
                line = line.replace("<!--WORKORDER-->", workorder)
            CompletionEmailBody += line
        template.close()


        for key in sorted(jobdict.iterkeys()):
            if jobdict[key][6] == 'DOCUMENTS':
                CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->", "<tr>\n" +
                                                                "<td>" + jobdict[key][1] + "</td>\n" +
                                                                "<td>"  + jobdict[key][2]  + "</td>\n" +
                                                                "<td>"  + jobdict[key][3]  + "</td>\n" +
                                                                "<td>"  + jobdict[key][4]  + "</td>\n" +
                                                                "<td>"  + jobdict[key][5]  + "</td>\n" +
                                                                #'<td style="color: rgb(0, 131, 173);">'  + str((int(jobdict[key][2])* 100)/total_mailpieces)   + "%</td>\n" +
                                                                "</tr>\n" + 
                                                                "<!--SPLIT_INFO-->\n")
                #CompletionEmailBody += line
                #template.close()


        #TOTAL Counts
        THE_SPLITTER = CompletionEmailBody.split("\n")
        CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->", "<tr>\n" +
                                                        '<td style="font-weight: bold;">Final Total</td>\n' +
                                                        "<td>"  + str(total_mailpieces)  + "</td>\n" +
                                                        "<td>"  + str(total_documents)  + "</td>\n" +
                                                        "<td>"  + str(total_pages)  + "</td>\n" +
                                                        "<td>"  + str(total_sheets)  + "</td>\n" +
                                                        "</tr>\n" +
                                                        "<!--SPLIT_INFO-->\n")
        #template.close()
        OtherFilesFirstFlag = True

# show any other output files produced

        for key in sorted(jobdict.iterkeys()):
            if jobdict[key][6] != 'DOCUMENTS':

                if OtherFilesFirstFlag == True:
                    CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->", '<tr><tr/>' + '<tr>\n' +
                    '<td style="color: rgb(0, 131, 173); font-weight: bold; text-decoration: underline;">Other Files</td>\n' +
                    "</tr>\n" +
                    "<!--SPLIT_INFO-->\n")
                    OtherFilesFirstFlag = False

                CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->","<tr>\n" +
                                        "<td>" + jobdict[key][1] + "</td>\n" +
                                        "<td>" + jobdict[key][2] + "</td>\n" +
                                        "<td>" + jobdict[key][3] + "</td>\n" +
                                        "<td>" + jobdict[key][4] + "</td>\n" +
                                        "<td>" + jobdict[key][5] + "</td>\n" +
                                        "</tr>\n" +
                                        "<!--SPLIT_INFO-->\n")

# show insert counts if requested

        if self._ShowInserts is True:
            #CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->", '<tr>\n' +
            #'<td style="color: rgb(0, 131, 173); font-weight: bold; text-decoration: underline;">Analysis of Inserts Used</td>\n' +
            #"</tr>\n" +
            #"<!--SPLIT_INFO-->\n")

            #CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_INFO-->",

            #CompletionEmailBody += '<PRE>' + 'JobID'.ljust(16) + 'Insert#1'.ljust(10) + 'Insert#2'.rjust(10) + 'Insert#3'.rjust(10) + 'Insert#4'.rjust(10) + 'Insert#5'.rjust(10) + 'Insert#6'.rjust(10) + '<PRE/><br />'

            for key in sorted(jobdict.iterkeys()):
                if jobdict[key][6] == 'DOCUMENTS':

                    CompletionEmailBody = CompletionEmailBody.replace("<!--EXTRAINFO-->", '<table class="Split-Table">\n' +
                                                                    '<tr>\n' +
                                                                    '<th style="color: rgb(0, 131, 173);">Analysis of Inserts Used</th>\n' +
                                                                    '</tr>\n' +
                                                                    '<tr style="color: rgb(0, 131, 173); font-size: 11; font-weight: bold;">\n' +
                                                                    '<td>Job ID</th>\n' +
                                                                    '<td>Insert 1</td>\n' +
                                                                    '<td>Insert 2</td>\n' +
                                                                    '<td>Insert 3</td>\n' +
                                                                    '<td>Insert 4</td>\n' +
                                                                    '<td>Insert 5</td>\n' +
                                                                    '<td>Insert 6</td>\n' +
                                                                    '</tr>\n' +
                                                                    '<!--SPLIT_AMT-->\n' +
                                                                    '</table>\n' +
                                                                    "<!--EXTRAINFO-->\n\n")


                    CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_AMT-->",
                                                                "<td>" + key + "</td>\n" +
                                                                "<td>" + jobdict[key][7] + "</td>\n" +
                                                                "<td>" + jobdict[key][8] + "</td>\n" +
                                                                "<td>" + jobdict[key][9] + "</td>\n" +
                                                                "<td>" + jobdict[key][10] + "</td>\n" +
                                                                "<td>" + jobdict[key][11] + "</td>\n" +
                                                                "<td>" + jobdict[key][12] + "</td>\n" +
                                                                "</tr>\n" +
                                                                "<!--SPLIT_AMT-->\n")

                #detail_line = (key.ljust(13) + ' ' + jobdict[key][7].rjust(10) +
                #                                          jobdict[key][8].rjust(10) +
                #                                          jobdict[key][9].rjust(10) +
                #                                          jobdict[key][10].rjust(10) +
                #                                          jobdict[key][11].rjust(10) +
                #                                          jobdict[key][12].rjust(10) + "\n")

                #CompletionEmailBody += '<PRE>' + detail_line + '<PRE/>'

            #CompletionEmailBody += '<PRE>' + dashes2 + '<PRE/>'
            CompletionEmailBody = CompletionEmailBody.replace("<!--SPLIT_AMT-->",
                                                        '<td style="font-weight: bold;">FINAL TOTAL</td>\n' +
                                                        "<td>" + str(total_insert1) + "</td>\n" +
                                                        "<td>" + str(total_insert2) + "</td>\n" +
                                                        "<td>" + str(total_insert3) + "</td>\n" +
                                                        "<td>" + str(total_insert4) + "</td>\n" +
                                                        "<td>" + str(total_insert5) + "</td>\n" +
                                                        "<td>" + str(total_insert6) + "</td>\n" +
                                                        "</tr>\n" +
                                                        "<!--SPLIT_AMT-->\n")


            #total_line = ' '.ljust(14) + str(total_insert1).rjust(10) + str(total_insert2).rjust(10) + str(total_insert3).rjust(10) + str(total_insert4).rjust(10) + str(total_insert5).rjust(10)+ str(total_insert6).rjust(10)
            #CompletionEmailBody += '<PRE>' + total_line + '<PRE/><br /><br />'

# show Breakdown of multiple input files if requested

        if self._ShowInputFiles is True:

            if os.path.exists(self._dir + self._pathSep + self._job + '.input_files.csv'):

                IFileCustTotal = 0
                IFilePagesTotal = 0
                IFileSheetsTotal = 0
                CompletionEmailBody += '<br /><br /><p>Summary of Input Files in Job:<p/><br />'
                CompletionEmailBody += '<p>' + heading2 + '<p/><br />'

                for line in csv.reader(open(self._dir + self._pathSep + self._job + '.input_files.csv','r')):

                    (BaseName,ext) = os.path.splitext(line[0])
                    if ext.upper() not in ['.AFP', '.TXT', '.PCL', '.PDF', '.PS', '.XML']:
                        BaseName = line[0]

                    IFileCustTotal += int(line[1])
                    IFilePagesTotal += int(line[2])
                    IFileSheetsTotal += int(line[3])
                    CompletionEmailBody += '<p>' + BaseName[:54].ljust(53) + str(line[1]).rjust(11) + str(line[2]).rjust(9) +str(line[3]).rjust(9) + '\n' + '<p/>'

                CompletionEmailBody += '<p>' + dashes4 + '<p/>'
                total_line = ' '.ljust(53) + str(IFileCustTotal).rjust(11) + str(IFilePagesTotal).rjust(9) + str(IFileSheetsTotal).rjust(9)
                CompletionEmailBody += '<p>' + total_line + '<p/><br /><br />'

# show Customer supplied counts if requested

        if self._ShowCustCounts is True:

            if os.path.exists(self._dir + self._pathSep + 'customer_counts.txt'):
                CustTotal = 0
                CompletionEmailBody += '<br /><br /><p>Customer Supplied Counts:<p/><br />'

                for line in csv.reader(open(self._dir + self._pathSep + 'customer_counts.txt','r')):

                    CustTotal += int(line[1])
                    CompletionEmailBody += '<p>' + line[0].ljust(44) + str(line[1]).rjust(10) + '\n' + '<p/>'

                CompletionEmailBody += '<p>' + dashes3 + '<p/>'
                total_line = ' '.ljust(44) + str(CustTotal).rjust(10)
                CompletionEmailBody += '<p>' + total_line + '<p/><br /><br />'

        emailFile = open(self._dir + self._pathSep + self._job + '_completion_email',"w")
        emailFile.write(CompletionEmailBody)
        emailFile.close()

        Msg = email.mime.Multipart.MIMEMultipart()
        Msg['Subject'] = customer_name + ' Job completed for JOBID:' + jobnum + ", WO#" + workorder
        Msg["From"]    = self._PDSEnvironment._ProducerEmailAddr

        if len(self._OverrideToAddr) > 0:
            Msg["To"]      = ", ".join(self._OverrideToAddr)
        else:
            Msg["To"]      = ", ".join(self._PDSEnvironment._ProdEmailAddr)

        body = email.mime.Text.MIMEText(CompletionEmailBody, "html")
        Msg.attach(body)

        # if os.path.exists(self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job + '_201.pdf'):
        #
        #     Report = open((self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job  + "_201.pdf"), 'rb')
        #     attachment = email.mime.application.MIMEApplication(Report.read(), _subtype="pdf")
        #     attachmentName = self._job + "_201.pdf"
        #     attachment.add_header('Content-Disposition', 'attachment',filename=attachmentName)
        #     Msg.attach(attachment)

        Att_image = open("/opt/adf/app/common/email_template/logo.png", "rb").read()
        img = MIMEImage(Att_image, 'png')
        img.add_header('Content-Id', '<logo>')
        img.add_header("Content-Disposition", "attachment", filename="logo")
        Msg.attach(img)

        #if os.path.exists(self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job + "_mm.pdf"):

        #    Report = open((self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job + "_mm.pdf"), 'rb')
        #    attachment = email.mime.application.MIMEApplication(Report.read(), _subtype="pdf")
        #    attachmentName = self._job + "_mm.pdf"
        #    attachment.add_header('Content-Disposition', 'attachment',filename=attachmentName)
        #    Msg.attach(attachment)

        #    Report = open((self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job + "_mmDetail.pdf"), 'rb')
        #    attachment = email.mime.application.MIMEApplication(Report.read(), _subtype="pdf")
        #    attachmentName = self._job + "_mmDetail.pdf"
        #    attachment.add_header('Content-Disposition', 'attachment',filename=attachmentName)
        #    Msg.attach(attachment)

        #if os.path.exists(self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job + "_stuffrpt.pdf"):

        #    Report = open((self._dir + self._pathSep + self._ReportsDir + self._pathSep + self._job  + "_stuffrpt.pdf"), 'rb')
        #    attachment = email.mime.application.MIMEApplication(Report.read(), _subtype="pdf")
        #    attachmentName = self._job + "_stuffrpt.pdf"
        #    attachment.add_header('Content-Disposition', 'attachment',filename=attachmentName)
        #    Msg.attach(attachment)

        if len(self._AttachList) > 0:

            for FileAttachment in self._AttachList:

                if os.path.exists(FileAttachment):
                    (_,FileName) = os.path.split(FileAttachment)
                    (_,FileType) = os.path.splitext(FileAttachment)
                    FileType = str(FileType[1:]).lower()
                    Report = open(FileAttachment, 'rb')
                    attachment = email.mime.application.MIMEApplication(Report.read(), _subtype=FileType)
                    attachmentName = str(FileName)
                    attachment.add_header('Content-Disposition', 'attachment',filename=attachmentName)
                    Msg.attach(attachment)

        if len(self._OverrideToAddr) > 0:
            self._PDSEnvironment._EmailServer.sendmail(self._PDSEnvironment._ProducerEmailAddr, self._OverrideToAddr, Msg.as_string())
        else:
            self._PDSEnvironment._EmailServer.sendmail(self._PDSEnvironment._ProducerEmailAddr, self._PDSEnvironment._ProdEmailAddr, Msg.as_string())

#  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
#
# Mainline code
#
#  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

processing = SendCompletionEmail(sys.argv)

#try:

processing.SendCompletionEmail()

#except Exception, ex:
#    print 'Unexpected Exception Caught during SendCompletionEmail: ' + str(ex)
#    raise ex

sys.exit(0)