import sys
from audioop import add
from inspect import istraceback
import xlsxwriter

def main():
    if len(sys.argv) > 1:
        args = sys.argv[1:]
    else:
        print("Please enter the path where .txt and .xlsx file are located.")
        exit()
    workbook = xlsxwriter.Workbook(args[1])
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', "Agenting No")
    worksheet.write('B1', "Process Date")
    worksheet.write('C1', "Process Time")
    worksheet.write('D1', "Settlement Date")
    worksheet.write('E1', "Address")
    worksheet.write('F1', "Email")
    worksheet.write('G1', "Tran Date")
    worksheet.write('H1', "Tran Time")
    worksheet.write('I1', "Vault ID")
    worksheet.write('J1', "Checking No")
    worksheet.write('K1', "Code No")
    worksheet.write('L1', "Type")
    worksheet.write('M1', "Msq Ref No")
    worksheet.write('N1', "Veri Code")
    worksheet.write('O1', "Trxn Amount")
    worksheet.write('P1', "Converted Amount")
    worksheet.write('Q1', "Flat Type")

    paragrah_count = 0
    line_num = 2
    agent_pos  = 0
    start_pos = 0
    end_pos = 0
    isTransaction = 0
    isAddress = 0
    
    agent_no = ""
    proc_date = ""
    proc_time = ""
    settle_date = ""
    address = ""
    address_res = ""
    email = ""
    
    with open(args[0]) as f:
        for line in f:
            if line.strip() == "END OF REPORT":
                paragrah_count += 1
                
            # Agent No
            agent_pos = line.find('AGENTING NO')
            if agent_pos > 0 and line.find(':') > 0:
                start_pos  = line.find(':') + 1
                end_pos = line.rfind('AGENT CODE')
                agent_no = line[start_pos: end_pos].strip()
                worksheet.write('A'+str(line_num), agent_no)
            
            # Process Date
            processDate_pos = line.find('PROC DATE')    
            if processDate_pos > 0:
                start_pos  = processDate_pos + len('PROC DATE');
                end_pos = line.rfind('TIME')
                proc_date = line[start_pos: end_pos].strip()
                proc_time = line[end_pos + len('TIME'): ].strip()
                worksheet.write('B'+str(line_num), proc_date)
                worksheet.write('C'+str(line_num), proc_time)
            
            # Settlement Date
            settle_pos = line.find('AGENTING TRANSACTIONS SETTLED ON')    
            if settle_pos > 0:
                start_pos  = settle_pos + len('AGENTING TRANSACTIONS SETTLED ON')
                settle_date = line[start_pos: start_pos + 12].strip()
                worksheet.write('D'+str(line_num), settle_date)
                
            # Address
            start_pos = line.find("ADDRESS")
            if start_pos >= 0 and line.find(':') > 0:
                isAddress = 1
            if isAddress == 1:
                end_pos = line.find('PHONE')
                if end_pos > 0:
                    address += line[line.find(':') + 1: end_pos].strip()
                    address += " "
                end_pos = line.find('FAX')
                if end_pos > 0:
                    address += line[0: end_pos].strip()
                    address += " "
                if line.find('PHONE') < 0 and line.find('FAX') < 0:
                    address += line.strip()
                    address += " "
            # Email
            email_pos = line.find('EMAIL')
            if email_pos > 0 and line.find(':') > 0:
                address_res = address[0: address.find('EMAIL')]
                worksheet.write('E'+str(line_num), address_res)
                address = ""
                isAddress = 0
                start_pos  = line.find(':') + 1;
                email = line[start_pos: ].strip()
                worksheet.write('F'+str(line_num), email)
            
            #Transaction
            if isTransaction == 2 and line.find('----------------------') < 0:
                    worksheet.write('A'+str(line_num), agent_no)
                    worksheet.write('B'+str(line_num), proc_date)
                    worksheet.write('C'+str(line_num), proc_time)
                    worksheet.write('D'+str(line_num), settle_date)
                    worksheet.write('E'+str(line_num), address_res)
                    worksheet.write('F'+str(line_num), email)
                    worksheet.write('G'+str(line_num), line[0: 5].strip())
                    worksheet.write('H'+str(line_num), line[6: 14].strip())
                    worksheet.write('I'+str(line_num), line[15: 23].strip())
                    worksheet.write('J'+str(line_num), line[24: 36].strip())
                    worksheet.write('K'+str(line_num), line[37: 50].strip())
                    worksheet.write('L'+str(line_num), line[51: 66].strip())
                    worksheet.write('M'+str(line_num), line[67: 90].strip())
                    worksheet.write('N'+str(line_num), line[91: 98].strip())
                    worksheet.write('O'+str(line_num), line[99: 110].strip())
                    worksheet.write('Q'+str(line_num), line[111: 118].strip())
                    line_num += 1
            if isTransaction == 2 and line.find('----------------------') >= 0:
                    isTransaction = 0
            if line.find('DTOF') >= 0 and line.find('TRXN') > 0 and line.find('VAULT') > 0:
                isTransaction = 1
            if isTransaction == 1 and line.find('----------------------') >= 0:
                isTransaction += 1
            
    f.close();

    workbook.close()
    print("Successfully Converted ", paragrah_count, "parts of ", args[0])
    
if __name__ == '__main__':
    main()