import argparse
import csv
from fuzzywuzzy import fuzz

CBX_FIRSTNAME = 0
CBX_LASTNAME = 1
CBX_ID = 2
CBX_BIRTHDATE = 3
CBX_COMPANY = 4

HC_FIRSTNAME = 0
HC_LASTNAME = 1
HC_COMPANY = 2


# define commandline parser
parser = argparse.ArgumentParser(description='Tool to match employees without birthday to employees ID in CBX, all input/output files must be in the current directory')
parser.add_argument('cbx_list',
                    help='UTF-8 csv DB export of employees with the following columns: ID, firstname, lastname, birthday, company')
parser.add_argument('hc_list',
                    help='Windows 1252 csv file with the following columns: firstname, lastname, company')
parser.add_argument('output',
                    help='csv file with the following columns: firstname, lastname, company')
args = parser.parse_args()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    data_path = './data/'
    cbx_file = data_path + args.cbx_list
    hc_file = data_path + args.hc_list
    output_file = data_path + args.output

    # output parameters used
    print(f'Reading CBX list: {args.cbx_list}')
    print(f'Reading HC list: {args.hc_list}')
    print(f'Outputing results in: {args.output}')

    # read data
    cbx_data = []
    hc_data = []
    with open(cbx_file, 'r', encoding="utf-8") as cbx:
        for row in csv.reader(cbx):
            cbx_data.append(row)
    with open(hc_file, 'r', encoding="cp1252") as hc:
        for row in csv.reader(hc):
            hc_data.append(row)

    # match
    for hc_row in hc_data:
        matches = {}
        hc_firstname = hc_row[HC_FIRSTNAME]
        hc_lastname = hc_row[HC_LASTNAME]
        hc_company = hc_row[HC_COMPANY]
        clean_hc_firstname = hc_firstname.lower().strip()
        clean_hc_lastname = hc_firstname.lower().strip()
        clean_hc_company = hc_company.lower().replace('.', '').replace(',', '').strip()

        for cbx_row in cbx_data:
            cbx_firstname = cbx_row[CBX_FIRSTNAME]
            cbx_lastname = cbx_row[CBX_LASTNAME]
            cbx_company = cbx_row[CBX_COMPANY]
            clean_cbx_firstname = cbx_firstname.lower().strip()
            clean_cbx_lastname = cbx_firstname.lower().strip()
            clean_cbx_company = cbx_company.lower().replace('.', '').replace(',', '').strip()

            ratio_firstname = fuzz.ratio(clean_cbx_firstname, clean_hc_firstname)
            ratio_lastname = fuzz.ratio(clean_cbx_lastname, clean_hc_lastname)
            ratio_company = fuzz.ratio(clean_cbx_company, clean_hc_company)
            if ratio_firstname > 80 and ratio_lastname > 80 and ratio_company > 60:
                overall_ratio = ratio_company * ratio_lastname * ratio_firstname / 10000
                if cbx_row[CBX_ID] in matches:
                    matches[cbx_row[CBX_ID]].append({'firstname':cbx_firstname,
                                                     'lastname': cbx_lastname,
                                                     'birthdate': cbx_row[CBX_BIRTHDATE],
                                                     'company': cbx_company,
                                                     'ratio':overall_ratio})
                else:
                    matches[cbx_row[CBX_ID]] = [{'firstname':cbx_firstname,
                                                 'lastname': cbx_lastname,
                                                 'birthdate': cbx_row[CBX_BIRTHDATE],
                                                 'company': cbx_company,
                                                 'ratio': overall_ratio}]
        ids = []
        for key, value in matches.items():
            companies = []
            value.sort(key=lambda x: x['ratio'], reverse=True)
            for item in value[0:5]:
                companies.append(f'{item["company"]}: {item["ratio"]}')
            ids.append(f'{key}, {item["firstname"]} {item["lastname"]}, {item["birthdate"]} --> {", ".join(companies)}')
        hc_row.append('\n'.join(ids))
    with open(output_file, 'w', newline='', encoding='cp1252') as resultfile:
        writer = csv.writer(resultfile)
        total = len(hc_data)
        index = 1
        for row in hc_data:
            print(f'{index} of {total}')
            writer.writerow(row)
            index += 1
