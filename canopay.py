import sys
import os
import csv
import pandas as pd
import subprocess



if __name__ == '__main__':


    try:

        ### Variables ###
        file_name = sys.argv[1]
        #file_name = "canopay.pdf"
        file_name = file_name.replace(".pdf","")
        txt = ""
        ls = []
        del_list = []


        print("Pdf to Excel conversion started ..")
        print("input file name : " + file_name)

        ## pdftotext ##

        p = subprocess.Popen('pdftotext -f 1 -l 1 -r 300 ' + file_name + '.pdf' + '   -layout', shell=True, stdout=subprocess.PIPE)
        stdout, stderr = p.communicate()
        print("error in pdftotext:" + str(stderr))

        ## csv file ##

        print("Creating csv file based on column positions")
        f = open(file_name + ".txt")
        lines = [line for line in f if line.strip().replace("  ",",")]
        with open(file_name + '.csv', mode='w+') as pdf_file:
            pdf_writer = csv.writer(pdf_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            for line_no,i in enumerate(lines):
               if ( "remarks" in i[0:20].strip().lower()):
                   break;
               if(len(i.strip()) > 6 and line_no > 3 ):
                   pdf_writer.writerow([i[0:20].strip(),i[21:37].strip(),i[38:125].strip(),i[127:140].strip(),i[141:165].strip(),i[166:190].strip(),i[191:250].strip()])

        ## Pandas ##

        print("Loading data in panda lib and making changes for booking text columns")

        df = pd.read_csv(file_name +'.csv')
        for i in range(len(df)):

            if pd.isna(df['Booking Date'][i]) and i != len(df)-1:
                txt =  ls[-1] + "\015" + df['Booking Text'][i]
                ls.pop()
                ls.append(txt)
                del_list.append(i)
            else:
                ls.append(df['Booking Text'][i])


        print("date and number columns standardization")
        df = df.drop(df.index[del_list])
        df['Booking Text'] = ls
        df['Booking Date'] = pd.to_datetime(df['Booking Date'], errors='coerce')
        df['Txn Date'] = pd.to_datetime(df['Txn Date'], errors='coerce')
        df['Value Date'] = pd.to_datetime(df['Value Date'], errors='coerce')
        df['Booking Date'] = df['Booking Date'].dt.strftime('%Y/%m/%d')
        df['Txn Date'] = df['Txn Date'].dt.strftime('%Y/%m/%d')
        df['Value Date'] = df['Value Date'].dt.strftime('%Y/%m/%d')
        df['Debit'] = df['Debit'].str.replace(',', '')
        df['Credit'] = df['Credit'].str.replace(',', '')
        df['Balance'] = df['Balance'].str.replace(',', '')

        df.to_excel(file_name + ".xlsx", index=False)

        print("Excel Conversion done.. file_name :" + file_name + ".xlsx")

    except OSError as err:
        print("OS error: {0}".format(err))
    except ValueError:
        print("Could not convert data to an integer.")
    except:
        print("Unexpected error:", sys.exc_info()[0])
        raise


