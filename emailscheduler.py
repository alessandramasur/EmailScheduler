"""
@author: Alessandra Masur
created on: 12.03.2023
"""
import win32com.client
import logging
import csv
import argparse
import datetime
import math
import yaml
from yaml.loader import SafeLoader

logger = logging.getLogger(__name__)
FORMAT = "[%(levelname)s] || %(message)s || file:%(name)s || line:%(lineno)d"
logging.basicConfig(level=logging.INFO, format=FORMAT)

def parse_yaml(pathtofile: str):
    """
    Parses the given YAML file.
    :pathtofile: path to YAML file as str
    :return: subject, science_text, admin_text all as string
    """
    with open(pathtofile) as f:
        data = yaml.load(f, Loader=SafeLoader)
        subject = data["subject"]
        body = data["body"]
        return subject, body
    
def parse_csv(pathtofile: str, delimiter: str):
    """
    Parses the given CSV file.
    :pathtofile: path to CSV file as str
    :delimiter: delimiter in CSV file as str
    :return: list with CSV rows
    """
    csvlist = []
    with open(pathtofile) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=delimiter)
        for row in csv_reader:
            csvlist.append(row)
    return csvlist

def write_emails(inputfile: str, delimiter: str, subject: str, body: str, preview: bool, mailsPerTenMin: int):
    """
    Schedule E-Mails from given CSV file with given text.

    :inputfile: path to CSV file
    :delimiter: delimiter in CSV file
    :subject: subject of the E-Mails as str
    :scientific_text: scientific body as str
    :admin_text: administrative body as str
    :preview: bool for turning preview mode on or off
    :mailsPerTenMin: how many E-Mails are sent every 10 minutes as int
    """
    
    outlook = win32com.client.Dispatch("Outlook.Application")

    f = open("scheduled_mails.CSV", "w", newline="")
    writer=csv.writer(f, delimiter=delimiter)

    inputlist = parse_csv(inputfile, delimiter)
    outputlist = []

    currentDay = datetime.datetime.now().day
    currentMonth = datetime.datetime.now().month
    currentYear = datetime.datetime.now().year

    mails = 0
    packets = 0
    for day in range(1, math.ceil(len(inputlist)/(mailsPerTenMin*72))+1):
        for hour in range(7, 19):
            for minute in range(0, 51, 10):
                deliverytime = datetime.datetime(currentYear, currentMonth, currentDay+day, hour, minute, 0, tzinfo=datetime.timezone.utc)
                # iterate inputlist
                if mails >= len(inputlist):
                    f.close()
                    logging.info("Excecuted successfully")
                    return
                for rec in inputlist[mailsPerTenMin*packets : (packets+1)*mailsPerTenMin]:
                    # check for duplicates
                    if rec[0] not in outputlist:
                        newmail = outlook.CreateItem(0)
                        newmail.Subject = subject
                        newmail.To = rec[0]
                        # insert name in mail body
                        newmail.HTMLBody = body.format(name=rec[1])
                        newmail.DeferredDeliveryTime = deliverytime
                        # ---preview mode--------------
                        if preview == True:
                            newmail.Display()
                            return
                        # -----------------------------
                        newmail.Send()
                        # log recipients
                        outputlist.append(rec[0])     
                        writer.writerow([rec[0], deliverytime])           
                    mails += 1
                packets += 1

    f.close()
    logging.info("Excecuted successfully")
    return


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--inputCSV", help=("Path to the input CSV file."), type=str, required=True)
    parser.add_argument("--delimiter", help=("Delimiter string/symbol in CSV file."), type=str, default=";", required=False)
    parser.add_argument("--inputYAML", help=("Path to the input YAML file."), default="input/mailcontent.yaml", required=False)
    parser.add_argument("--mailsPerTenMin", help=("How many mails should be sent every ten minutes."), default="35", required=False)
    parser.add_argument("--previewMode", help=("Set 'True' or 'False' for turning preview mode on or off."), type=str, default="True", required=False)
    args = parser.parse_args()

    mailsPerTenMin = int(args.mailsPerTenMin)
    if args.previewMode == "False":
        preview = False
    else:
        preview = True
    subject, body = parse_yaml(args.inputYAML)

    write_emails(args.inputCSV, args.delimiter, subject, body, preview, mailsPerTenMin)


if __name__ == "__main__":
    main()