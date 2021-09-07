import argparse
import csv
from fuzzywuzzy import fuzz

CBX_FIRSTNAME = 0
CBX_LASTNAME = 1
CBX_ID = 2
CBX_BIRTHDATE = 3
CBX_COMPANY = 4

HC_COMPANY = 0
HC_FIRSTNAME = 1
HC_LASTNAME = 2

# define commandline parser
parser = argparse.ArgumentParser(description='Tool to match employees without birthday to employees ID in CBX, all input/output files must be in the current directory', formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('cbx_list',
                    help='''csv DB export file of employees with the following columns: 
                        Cognibox ID, firstname, lastname, birthdate, contractor''')

parser.add_argument('hc_list',
                    help='''csv file with the following columns:
    contractor, firstname, lastname, any other columns...'''
                    )
parser.add_argument('output',
                    help='''csv file with the following columns: 
    contractor, firstname, lastname, any other columns..., Cognibox ID, birthdate,  matching information  
Matching information format:
    Cognibox ID, firstname lastname, birthdate --> Contractor 1, match ratio 1,
    Contractor 2, match ratio 2, etc...
The matching ratio is a value betwween 0 and 100, where 100 is a perfect best match.
Please note the Cognibox ID is set ONLY if a single match his found. If no match
or multiple matches are found it is left empty.''')

parser.add_argument('--cbx_list_encoding', dest='cbx_encoding', action='store',
                    default='utf-8',
                    help='Encoding for the cbx list (default: utf-8)')

parser.add_argument('--hc_list_encoding', dest='hc_encoding', action='store',
                    default='cp1252',
                    help='Encoding for the hc list (default: cp1252)')

parser.add_argument('--output_encoding', dest='output_encoding', action='store',
                    default='cp1252',
                    help='Encoding for the hc list (default: cp1252)')


parser.add_argument('--min_company_match_ratio', dest='ratio', action='store',
                    default=60,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 60)')

args = parser.parse_args()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    data_path = './data/'
    cbx_file = data_path + args.cbx_list
    hc_file = data_path + args.hc_list
    output_file = data_path + args.output

    # output parameters used
    print(f'Reading CBX list: {args.cbx_list} [{args.cbx_encoding}]')
    print(f'Reading HC list: {args.hc_list} [{args.hc_encoding}]')
    print(f'Outputing results in: {args.output} [{args.output_encoding}]')
    print(f'contractor match ratio: {args.ratio}')
    # read data
    cbx_data = []
    hc_data = []
    with open(cbx_file, 'r', encoding=args.cbx_encoding) as cbx:
        for row in csv.reader(cbx):
            cbx_data.append(row)
    with open(hc_file, 'r', encoding=args.hc_encoding) as hc:
        for row in csv.reader(hc):
            hc_data.append(row)

    with open(output_file, 'w', newline='', encoding=args.output_encoding) as resultfile:
        writer = csv.writer(resultfile)

        # match
        total = len(hc_data)
        index = 1
        for hc_row in hc_data:
            matches = {}
            hc_firstname = hc_row[HC_FIRSTNAME]
            hc_lastname = hc_row[HC_LASTNAME]
            hc_company = hc_row[HC_COMPANY]
            for cbx_row in cbx_data:
                cbx_firstname = cbx_row[CBX_FIRSTNAME]
                cbx_lastname = cbx_row[CBX_LASTNAME]
                cbx_company = cbx_row[CBX_COMPANY]
                clean_cbx_firstname = cbx_firstname.lower().strip()
                clean_cbx_lastname = cbx_firstname.lower().strip()
                clean_cbx_company = cbx_company.lower().replace('.', '').replace(',', '').strip()

                ratio_firstname = fuzz.ratio(cbx_firstname.lower().strip(),hc_firstname.lower().strip())
                ratio_lastname = fuzz.ratio(cbx_lastname.lower().strip(), hc_lastname.lower().strip())
                ratio_company = fuzz.ratio(cbx_company.lower().replace('.', '').replace(',', '').strip(),
                                           hc_company.lower().replace('.', '').replace(',', '').strip())
                if ratio_firstname >= 90 and ratio_lastname >= 90 and ratio_company >= float(args.ratio):
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

            # append matching results to the hc_list
            if len(matches) == 1:
                key = list(matches.keys())[0]
                hc_row.append(key)
                hc_row.append(matches[key][0]["birthdate"])
            else:
                hc_row.append('')
                hc_row.append('')
            hc_row.append('\n'.join(ids))
            writer.writerow(hc_row)
            print(f'{index} of {total} [{len(matches)} found]')
            index += 1
