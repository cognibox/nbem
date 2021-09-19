import argparse
import csv
from fuzzywuzzy import fuzz

CBX_FIRSTNAME = 0
CBX_LASTNAME = 1
CBX_ID = 2
CBX_BIRTHDATE = 3
CBX_COMPANY = 4
CBX_PARENTS = 5
CBX_PREVIOUS = 6

HC_COMPANY = 0
HC_FIRSTNAME = 1
HC_LASTNAME = 2

# define commandline parser
parser = argparse.ArgumentParser(description='Tool to match employees without birthday to employees ID in CBX, '
                                             'all input/output files must be in the current directory',
                                 formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('cbx_list',
                    help='''csv DB export file (no header) of employees with the following columns: 
    Cognibox ID, firstname, lastname, birthdate, contractor, 
    contractor parent list,
    employee previous employer list''')

parser.add_argument('hc_list',
                    help='''csv file (with header) and the following columns:
    contractor, firstname, lastname, any other columns...'''
                    )
parser.add_argument('output',
                    help='''csv file with the following columns: 
    contractor, firstname, lastname, any other columns..., Cognibox ID, birthdate,  best matching score, 
    partial name matching, matching information  
Matching information format:
    Cognibox ID, firstname lastname, birthdate --> Contractor 1 [parents: C1 parent1;C1 parent2;etc..] 
    [previous: Empl. Previous1;Empl. Previous2], match ratio 1,
    Contractor 2 [C2 parent1;C2 parent2;etc..], match ratio 2, etc...
The matching ratio is a value betwween 0 and 100, where 100 is a perfect match.
Please note the Cognibox ID and birthdate is set ONLY if a single match his found. If no match
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


parser.add_argument('--min_company_match_ratio', dest='ratio_company', action='store',
                    default=60,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 60)')

parser.add_argument('--list_seperator', dest='list_seperator', action='store',
                    default=';',
                    help='string seperator used for lists (default: ;)')

parser.add_argument('--min_name_match_ratio', dest='ratio_name', action='store',
                    default=90,
                    help='Minimum match ratio for contractors, between 0 and 100 (default 90)')

args = parser.parse_args()

if __name__ == '__main__':
    data_path = './data/'
    cbx_file = data_path + args.cbx_list
    hc_file = data_path + args.hc_list
    output_file = data_path + args.output

    # output parameters used
    print(f'Reading CBX list: {args.cbx_list} [{args.cbx_encoding}]')
    print(f'Reading HC list: {args.hc_list} [{args.hc_encoding}]')
    print(f'Outputing results in: {args.output} [{args.output_encoding}]')
    print(f'contractor match ratio: {args.ratio_company}')
    print(f'employee match ratio: {args.ratio_name}')
    # read data
    cbx_data = []
    hc_data = []
    print('Reading Cognibox data file...')
    with open(cbx_file, 'r', encoding=args.cbx_encoding) as cbx:
        for row in csv.reader(cbx):
            cbx_data.append(row)
    print(f'Completed reading {len(cbx_data)} employees.')
    print('Reading hiring client data file...')
    with open(hc_file, 'r', encoding=args.hc_encoding) as hc:
        for row in csv.reader(hc):
            hc_data.append(row)
    print(f'Completed reading {len(hc_data)} employees.')
    with open(output_file, 'w', newline='', encoding=args.output_encoding) as resultfile:
        writer = csv.writer(resultfile)
        headers = hc_data.pop(0)
        headers.extend(['Cognibox ID', 'birthdate', 'best matching score', 'is partial name match',
                       'matching information'])
        writer.writerow(headers)

        # match
        total = len(hc_data)
        index = 1
        for hc_row in hc_data:
            matches = []
            hc_firstname = hc_row[HC_FIRSTNAME]
            hc_lastname = hc_row[HC_LASTNAME]
            hc_company = hc_row[HC_COMPANY]
            clean_hc_company = hc_company.lower().replace('.', '').replace(',', '').strip()
            for cbx_row in cbx_data:
                cbx_firstname = cbx_row[CBX_FIRSTNAME]
                cbx_lastname = cbx_row[CBX_LASTNAME]
                cbx_company = cbx_row[CBX_COMPANY]
                cbx_parents = cbx_row[CBX_PARENTS]
                cbx_previous = cbx_row[CBX_PREVIOUS]
                cbx_name = f'{cbx_firstname.lower().strip()} {cbx_lastname.lower().strip()}'.strip()
                hc_name = f'{hc_firstname.lower().strip()} {hc_lastname.lower().strip()}'.strip()
                ratio_name = fuzz.token_set_ratio(cbx_name, hc_name)
                ratio_name_exact = fuzz.token_sort_ratio(cbx_name, hc_name)

                if ratio_name >= float(args.ratio_name):
                    partial = True if ratio_name_exact < float(args.ratio_name) else False
                    ratio_company = fuzz.token_sort_ratio(cbx_company.lower().replace('.', '').replace(',', '').strip(),
                                                          clean_hc_company)
                    ratio_parent = 0
                    cbx_parent_list = cbx_parents.split(args.list_seperator)
                    for item in cbx_parent_list:
                        if item == cbx_company:
                            continue
                        ratio = fuzz.token_sort_ratio(item.lower().replace('.', '').replace(',', '').strip(),
                                                      clean_hc_company)
                        ratio_parent = ratio if ratio > ratio_parent else ratio_parent
                    ratio_previous = 0
                    for item in cbx_previous.split(args.list_seperator):
                        if item == cbx_company or item in cbx_parent_list:
                            continue
                        ratio = fuzz.token_sort_ratio(item.lower().replace('.', '').replace(',', '').strip(),
                                                      clean_hc_company)
                        ratio_previous = ratio if ratio > ratio_previous else ratio_previous
                    if (ratio_company >= float(args.ratio_company) or
                            ratio_parent >= float(args.ratio_company) or
                            ratio_previous >= float(args.ratio_company)):
                        print('   --> ',cbx_firstname, cbx_lastname, cbx_company, cbx_row[CBX_ID], ratio_company,
                              ratio_parent, ratio_previous)
                        ratio_company = ratio_parent if ratio_parent > ratio_company else ratio_company
                        ratio_company = ratio_previous if ratio_previous > ratio_company else ratio_company
                        overall_ratio = ratio_company * ratio_name / 100
                        parent_str = f'[parent: {cbx_parents}]' if cbx_parents else None
                        previous_str = f'[previous: {cbx_previous}]' if cbx_previous else None
                        display = [cbx_company]
                        if parent_str:
                            display.append(parent_str)
                        if previous_str:
                            display.append(previous_str)
                        cbx_company = ' '.join(display)
                        matches.append({'cbx_id': cbx_row[CBX_ID],
                                        'firstname': cbx_firstname,
                                        'lastname': cbx_lastname,
                                        'birthdate': cbx_row[CBX_BIRTHDATE],
                                        'company': cbx_company,
                                        'ratio': overall_ratio,
                                        'partial': partial, })
            ids = []
            best_match = 0
            matches.sort(key=lambda x: x['ratio'], reverse=True)
            companies = []
            if matches:
                best_match = matches[0]['ratio'] if matches[0]['ratio'] > best_match else best_match
            for item in matches[0:5]:
                companies.append(f'{item["company"]}: {item["ratio"]}')
                ids.append(f'{item["cbx_id"]}, {item["firstname"]} {item["lastname"]},'
                           f'{item["birthdate"]} --> {", ".join(companies)}')

            # append matching results to the hc_list
            uniques_cbx_id = set(item['cbx_id'] for item in matches)
            hc_row.append(matches[0]["cbx_id"] if len(uniques_cbx_id) == 1 else '?' if len(uniques_cbx_id) > 1 else '')
            hc_row.append(matches[0]["birthdate"] if len(uniques_cbx_id) == 1 else '')
            hc_row.append(best_match if len(uniques_cbx_id) == 1 else '')
            hc_row.append(partial if len(uniques_cbx_id) == 1 else '')
            hc_row.append('|'.join(ids))
            writer.writerow(hc_row)
            print(f'{index} of {total} [{len(uniques_cbx_id)} found]')
            index += 1
