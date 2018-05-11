#!/usr/bin/env python2
#
# This is a helper script for generating a list of transactions for the tax
# declaration.
#
# Author: Martin Schobert <martin@weltregierung.de>
# License: I don't care.
#

import sys
import gzip
import xml.etree.ElementTree as ET
import operator
from decimal import *
import re
from prettytable import *
import functools
import datetime
from optparse import OptionParser

from xml.sax.saxutils import escape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Frame, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A3, A4, landscape, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfgen import canvas


lang = { 'de' : { 'date' : 'Datum',
                  'amount' : 'Betrag',
                  'description' : 'Beschreibung',
                  'document_avail' : 'Beleg'},
         'en' : { 'date' : 'Date',
                  'amount' : 'Amount',
                  'description' : 'Description',
                  'document_avail' : 'Doc'}
         }

class Transaction:
    def __init__(self, acc, date, value, descr):
        self.acc = acc
        self.date = date
        self.value = value
        self.descr = descr

    def get_date(self):
        return self.date

    def get_value(self):
        return self.value

    def get_descr(self):
        return self.descr

class Account:
    def __init__(self, acc_id, acc_name):
        self.accounts = []
        self.id = acc_id
        self.name = acc_name
        self.tax = False
        self.root = False
        self.transactions = []

    def add_transaction(self, t):
        self.transactions.append(t)

    def has_transactions(self):
        if self.transactions:
            return True
        return False

    def print_transactions(self, transaction_visitor_fnk):
        sum = Decimal(0)

        x = PrettyTable(["Date", "Amount", "Description"])
        x.align["Amount"] = "r"
        x.align["Description"] = "l"

        for t in sorted(self.transactions, key=lambda t: t.get_date(), reverse=True):

            x.add_row([t.get_date(), t.get_value(), t.get_descr()])
            if transaction_visitor_fnk:
                transaction_visitor_fnk(t)

#        for t in sorted(self.transactions, key=lambda t: t.get_date(), reverse=True):
#            print "%10s %10.2f %s" % (t.get_date(), t.get_value(), t.get_descr())
            sum += t.get_value()

        x.add_row(['', sum, ''])
        print x

#        print "%10s %10.2f" % ("", sum)

    def add_sub_account(self, acc):
        if self.tax:
            acc.set_tax_relevant()

#        acc.set_name(self.name + " / " + acc.get_name())
        self.accounts.append(acc)

    def get_id(self):
        return self.id

    def get_name(self):
        return self.name

    def set_name(self, name):
        self.name = name

    def set_tax_relevant(self):
        self.tax = True
        for sub in self.accounts:
            sub.set_tax_relevant()

    def is_tax_relevant(self):
        return self.tax

    def reset_name(self, prefix = ''):
        if prefix != '':
            self.set_name(prefix + " / " + self.name)

        for sub in self.accounts:
            sub.reset_name(self.get_name())
            

    def set_root(self):
        self.root = True

    def is_root(self):
        return self.root

    def is_expense(self):
        return 'Ausgabe' in self.name

class AccountSet:

    def __init__(self):
        self.accounts = {}

    def add(self, acc):
        self.accounts[acc.get_id()] = acc

    def has(self, acc):
        return acc in self.accounts

    def get(self, acc):
        return self.accounts[acc]

    def print_tax_relevant_accounts(self, account_visitor_fnk, after_account_visitor_fnk, transaction_visitor_fnk, year, print_empty_categories=False):
#        print "tax relevant accounts"
#        print "-----------------------"

        for a in (sorted(self.accounts, key = lambda aid: self.accounts[aid].name)):
            acc = self.accounts[a]
            if acc.is_tax_relevant() and (print_empty_categories or acc.has_transactions()):
                print acc.get_id() + " " + acc.get_name().encode('utf-8')
                account_visitor_fnk(acc)
 #               print "--------------------------------------------------------"
                acc.print_transactions(transaction_visitor_fnk)

                print "\n\n"
                if after_account_visitor_fnk:
                    after_account_visitor_fnk(acc)



    def reset_names(self):
        for a in self.accounts:
            acc = self.accounts[a]
            if acc.is_root():
                acc.reset_name()

    

class Report:
    def __init__(self, filename, lang):
        self.lang = lang
        self.elements = []
        self.summary_elements = []
        self.doc = SimpleDocTemplate(filename, pagesize=A4)
        self.styles = getSampleStyleSheet()
        # container for the "Flowable" objects
        #styleN = styles["Normal"]

        self.row_sizes = [2.5 * cm, 2 * cm, 9 *cm, 1.5 * cm]


    def report_account(self, acc):
        self.report_account_title(acc)
        return self.report_account_table_head(acc)

    def report_account_title(self, acc):

        # make category box
        tableHeading = [[Paragraph("<para align=center>" + acc.get_name().encode('utf-8') + "</para>",self.styles['Normal'])]]
        tH = Table(tableHeading, [15 *cm])
        tH.hAlign = 'LEFT'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                               ('VALIGN',(0,0),(-1,-1),'TOP'),
                               ('BOX',(0,0),(-1,-1),1,colors.black),
                               ('INNERGRID',(0,0),(-1,-1),1,colors.black)] )
        tblStyle.add('BACKGROUND',(0,0),(-1,-1),colors.lightblue)
        tH.setStyle(tblStyle)
        self.elements.append(tH)
        self.elements.append(Spacer(1, 0.3 * cm))

    def report_account_table_head(self, acc):
        # Make heading for each column
        column1Heading = Paragraph("<para align=center>" + lang[self.lang]['date'] + "</para>",self.styles['Normal'])
        column2Heading = Paragraph("<para align=center>" + lang[self.lang]['amount'] + "</para>",self.styles['Normal'])
        column3Heading = Paragraph("<para align=center>" + lang[self.lang]['description'] + "</para>",self.styles['Normal'])
        column4Heading = Paragraph("<para align=center>" + lang[self.lang]['document_avail'] + "</para>",self.styles['Normal'])

        row_array = [column1Heading, column2Heading, column3Heading, column4Heading]
        tableHeading = [row_array]
        tH = Table(tableHeading, self.row_sizes)   # These are the column widths for the headings on the table
        tH.hAlign = 'LEFT'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                               ('VALIGN',(0,0),(-1,-1),'TOP'),
                               ('BOX',(0,0),(-1,-1),1,colors.black),
                               ('INNERGRID',(0,0),(-1,-1),1,colors.black)] )
        tblStyle.add('BACKGROUND',(0,0),(-1,-1),colors.lightblue)
        tH.setStyle(tblStyle)
        self.elements.append(tH)

        self.sum = Decimal(0)

    def add_row(self, row_data):
        # Assemble rows of data for each column
        column1Data = Paragraph("<para align=center> " + row_data[0] + "</para>", self.styles['Normal'])
        column2Data = Paragraph("<para align=right> " + row_data[1] + "</para>", self.styles['Normal'])
        column3Data = Paragraph("<para align=left> " + escape(row_data[2]) + "</para>", self.styles['Normal'])
        column4Data = Paragraph("<para align=left> </para>", self.styles['Normal'])
        tableRow = [[column1Data, column2Data, column3Data, column4Data]]
            
        tR=Table(tableRow, self.row_sizes)   
        tR.hAlign = 'LEFT'
        tR.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.white),
                                ('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                                ('VALIGN',(0,0),(-1,-1),'TOP'),
                                ('BOX',(0,0),(-1,-1),1,colors.black),
                                ('INNERGRID',(0,0),(-1,-1),1,colors.black)]))
        self.elements.append(tR)
        del tR

    
    def report_transaction(self, t):
        self.add_row([t.get_date(), '{0:.2f}'.format(t.get_value()), t.get_descr()])
        self.sum += t.get_value();

    def after_account(self, acc):
        self.add_row(['', '{0:.2f}'.format(self.sum), ''])
        #self.elements.append(Spacer(1, 1 * cm))
        self._newpage()
        
    def render_transactions(self):
        self.doc.build(self.elements)

    def _folding_line(self, acc = None):
        self.elements.append(Spacer(1, 22.5 * cm))
        self.elements.append(Paragraph("<para align=left><u>" + ("&nbsp;" * 3) + "_______________________________________________________________________________</u></para>", self.styles['Normal']))
        self._newpage()

    def _newpage(self, acc = None):
        self.elements.append(PageBreak())
        
    def render_seperators(self, accounts, year, print_empty_categories):
        accounts.print_tax_relevant_accounts(self.report_account_title, self._folding_line, None, year, print_empty_categories)
        


def get_xml(zf):
    f = gzip.open(zf, 'rb')
    file_content = f.read()
    f.close()

    try:
        parsed_doc = ET.fromstring(file_content)
        return parsed_doc
        
    except ET.ParseError as e:
        pos = e.position
        print(e)
        print("+ XML parse error in line %d, column %d:" % (pos[0], pos[1]))
        print(file_content.splitlines()[pos[0]])

    return None

def get_tax_accounts(xml_root):

    accounts = AccountSet()

    for a in xml_root.findall(".//ACCOUNTS/ACCOUNT"):

        aid = a.attrib['id']
        if 'name' in a.attrib:
            aname = a.attrib['name']
        else:
            aname = '------------'

        # create an account, even if not tax relevant
        acc = Account(aid, aname)

        accounts.add(acc)

        if len(a.attrib['parentaccount']) == 0:
            acc.set_root()

        tax_relevant = a.findall(".//PAIR[@key='Tax']")
        if tax_relevant and tax_relevant[0].attrib['value'] == 'Yes':            
#            print a.attrib['id'] + " " + a.attrib['name'].encode('utf-8')
            acc.set_tax_relevant()
#            print "   TAX"

    check_sub_accounts(xml_root, accounts)

    return accounts

#def check_sub_accounts(xml_root, a, accounts, acc):
def check_sub_accounts(xml_root, accounts):

    for a in xml_root.findall("./ACCOUNTS/ACCOUNT"):

        acc = accounts.get(a.attrib['id'])

#        print "+ acc: " + acc.get_id() + " " + acc.get_name().encode('utf-8')
#        if acc.is_tax_relevant():
#            print "  TAX"

        for sub in a.findall("./SUBACCOUNTS/SUBACCOUNT"):
                
            sub_acc_id = sub.attrib['id']
            sub_acc = accounts.get(sub_acc_id)
            sub_acc_name = sub_acc.get_name()

            acc.add_sub_account(sub_acc)

#            print "\t+ subs: " + sub_acc_id + " " + sub_acc_name.encode('utf-8')
#            if acc.is_tax_relevant():
#                print "\t  TAX"


def remove_newlines(str):
    return str.replace('\n', '')

def lookup_payee(xml_root, payee_id):
    e = xml_root.find("./PAYEES/PAYEE[@id=\"%s\"]" % (payee_id))
    if e == None:
        return None
    else:
        return e.attrib['name']
    
        
    
def get_transactions(xml_root, accounts, year):
    for t in xml_root.findall("./TRANSACTIONS/TRANSACTION"):
        postdate = t.attrib['postdate']
        memo = remove_newlines(t.attrib['memo'])

        postdate_dt = datetime.datetime.strptime(postdate, "%Y-%m-%d").date()

        if postdate_dt.year != year:
            continue

        payee = ""
        
        # split bookings
        for s in t.findall(".//SPLIT"):
            val = s.attrib['value']
            acc_id = s.attrib['account']
            acc = accounts.get(acc_id)
            acc_name = acc.get_name().encode('utf-8')
            if payee == "":
                payee = lookup_payee(xml_root, s.attrib['payee'])
                if payee:
                    payee += ":\n"
                else:
                    payee = ""
            
            if s.attrib['memo']:
                mem = payee + remove_newlines(s.attrib['memo'])
            else:
                mem = payee + memo

            
            
            # is it an expese or income account
            if acc.is_tax_relevant():

                print "\n" + postdate + ": " + memo
                print "+ %10s : %s | %s" % (val, acc_name, mem.encode('utf-8'))
                
                matchObj = re.match(r'^([\+\-]?)(\d+)\/(\d+)', val)
                if matchObj:
                    val = Decimal(matchObj.group(2)) / Decimal(matchObj.group(3))
                    if matchObj.group(1) == "-":
                        val *= Decimal(-1)

                    print  str(val) + "-> " + acc_name

                    trans = Transaction(acc, postdate, Decimal(-1) * val, mem)
                    acc.add_transaction(trans)


def main(filename, year, outfile,lang, print_empty_categories):
    xml_root = get_xml(filename)
    if xml_root == None:
        print "+ Error: can't parse XML"
        return
    
    accounts = get_tax_accounts(xml_root)
    accounts.reset_names()
    get_transactions(xml_root, accounts, year)

    report = Report(outfile, lang)

    report.render_seperators(accounts, year, print_empty_categories)
    accounts.print_tax_relevant_accounts(report.report_account, report.after_account, report.report_transaction, year, print_empty_categories)
    report.render_transactions()

if __name__ == "__main__":
    getcontext().prec = 6
    reload(sys)
    sys.setdefaultencoding('utf-8')

    parser = OptionParser()
    parser.add_option("--file", dest="file", help="The KMyMoney file to use")
    parser.add_option("--year", dest="year", help="Render transactions for a year", type='int', default = datetime.date.today().year - 1)
    parser.add_option("--out", dest="outfile", help="Report file name")
    parser.add_option("--lang", dest="language", help="Language (%s)" % (",".join(lang.keys())), default='de')
    parser.add_option("--print-empty-categories", dest="print_empty_categories", action='store_true', help="print categories without transactions")

    (options, args) = parser.parse_args()


    options.language = options.language.lower()

    if not options.outfile:
        options.outfile = "%d_transactions_report.pdf" % (options.year)

    if options.language not in lang.keys():
        print "Language not supported."
        sys.exit()

    if options.file:
        main(options.file, options.year, options.outfile, options.language, options.print_empty_categories)
    else:
        print "Please, specify a KMyMoney file."

