import random
import string
import uuid
import pandas as pd
from datetime import datetime

def generate_random_alphanumeric(length=16):
  characters = string.ascii_letters + string.digits
  return ''.join(random.choice(characters) for _ in range(length))


df = pd.read_excel("New Microsoft Excel Worksheet.xlsx")
df = df.dropna()

xml_content = f"""
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">"""
         
xml_end = """</TALLYMESSAGE>
        </REQUESTDATA>
        </IMPORTDATA>
        </BODY>
    </ENVELOPE>
    """

vourcher_count = 0
for index, row in df.iterrows():
    vourcher_count += 1
    # Generate a random alphanumeric string
    random_string = generate_random_alphanumeric()
    
    # Generate a UUID
    unique_id = str(uuid.uuid4())
    
    voucher_date = row['Voucher dt'].strftime('%Y%m%d')
    amount = row['amt']
    instrument_no = row['Inst N']
    naration = row['Narration']

    xml_content += f"""
        <VOUCHER VCHTYPE="Receipt" ACTION="Create" OBJVIEW="Accounting Voucher View">
            <DATE>{voucher_date}</DATE>
            <NARRATION>{naration}</NARRATION>
            <VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>
            <PARTYLEDGERNAME>ICICI-HUDCO Br. A/c No-350401000372</PARTYLEDGERNAME>
            <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
            <VOUCHERNUMBER>{vourcher_count}</VOUCHERNUMBER>
            <EFFECTIVEDATE>{voucher_date}</EFFECTIVEDATE>
            <ISINVOICE>No</ISINVOICE>
            <ISDELETED>No</ISDELETED>

            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>ICICI-HUDCO Br. A/c No-350401000372</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <LEDGERFROMITEM>No</LEDGERFROMITEM>
              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
              <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
              <AMOUNT>-{amount}</AMOUNT>

              <BANKALLOCATIONS.LIST>
                <DATE>{voucher_date}</DATE>
                <INSTRUMENTDATE>{voucher_date}</INSTRUMENTDATE>
                <NAME>{unique_id}</NAME>
                <TRANSACTIONTYPE>Others</TRANSACTIONTYPE>
                <PAYMENTFAVOURING>General Donation (Non 80-G)</PAYMENTFAVOURING>
                <INSTRUMENTNUMBER>{instrument_no}</INSTRUMENTNUMBER>
                <UNIQUEREFERENCENUMBER>{random_string}</UNIQUEREFERENCENUMBER>
                <PAYMENTMODE>Transacted</PAYMENTMODE>
                <BANKPARTYNAME>General Donation (Non 80-G)</BANKPARTYNAME>
                <AMOUNT>{amount}</AMOUNT>
            </BANKALLOCATIONS.LIST>
            </ALLLEDGERENTRIES.LIST>

            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>General Donation (Non 80-G)</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <LEDGERFROMITEM>No</LEDGERFROMITEM>
              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
              <ISPARTYLEDGER>No</ISPARTYLEDGER>
              <AMOUNT>{amount}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>

        </VOUCHER>"""

xml_content += xml_end

# Save the XML content to a file
with open("try4.xml", "w") as file:
    file.write(xml_content)
print("XML file generated successfully.")
