#!/usr/bin/python3
import os
import sys
import argparse
import pandas as pd


def detect_version(file_name):
	h = "".join([hex(ord(c))[2:].zfill(2) for c in file_name.read(2)])
	file_name.seek(0)
	if h == '7141':
		return 180  # '4.6c'
	elif h == '3241':
		return 200  # 'non-unicode'
	elif h == '0032':
		return 400  # 'unicode'
	elif h == '3200':
		return 400  # 'unicode'
	elif h == '0478':
		return 400  # 'unicode'
	else:
		print("Failed to detect block size: ", h)
		sys.exit()


def take_block(block):
	return [block[1:4],                                                 # (1) (?) EventId
			block[4:8] + '.' + block[8:10] + '.' + block[10:12],        # (2)  Date
			block[12:14] + ':' + block[14:16] + ':' + block[16:18],     # (3)  Time
			block[18:25],                                               # (4) (?) OS Process ID
			block[25:30],                                               # (5) (?) SAP Process ID
			block[30],                                                  # (6)  Type of connection (Dialog, RFC, etc.)
			block[31],                                                  # (7) (?) SAP Process ID in hex
			block[32:40].strip(),                                       # (8)  Terminal
			block[40:52].strip(),                                       # (9)  User LOGIN
			block[52:72].strip(),                                       # (10) Transaction
			block[72:112].strip(),                                      # (11) Report
			block[112:115],                                             # (12) Client
			block[115],                                                 # (13) (?) SessionID
			block[116:180].strip(),                                     # (14) Transaction parameters
			block[180:].strip()]                                        # (15) IP/FQDN


def parsing_for_its_in_args(its, args):
	if args:
		for arg in args:
			for it in its:
				if arg in it:
					return True
		return False
	else:
		return True


def remove_extra_cols(element):
	cols_dict = {'eventid': 0, 'date': 1, 'time': 2, 'osid': 3, 'sapid': 4, 'typecon': 5, 'sapidhex': 6, 'terminal': 7,
				'login': 8, 'tcode': 9, 'report': 10, 'client': 11, 'sessionid': 12, 'param': 13, 'fqdn': 14}
	if parsed_args.remove:
		cols_for_remove = []
		for col in parsed_args.remove:
			cols_for_remove.append(cols_dict[col])
		cols_for_remove.sort(reverse=True)
		for col in cols_for_remove:
			del element[col]
	return element


def parse_file(filename):
	res = []
	with open(filename, encoding='latin-1') as f:
		block_size = detect_version(f)
		block = f.read(block_size)
		while block != '':
			item = take_block(block[0:len(block):2])
			if all([parsing_for_its_in_args([item[7], item[14]], parsed_args.terminal),
					parsing_for_its_in_args([item[8]], parsed_args.login),
					parsing_for_its_in_args([item[9], item[13]], parsed_args.tcode),
					parsing_for_its_in_args([item[10]], parsed_args.report),
					parsing_for_its_in_args([item[11]], parsed_args.client)]):
				res.append(remove_extra_cols(item))
			block = f.read(block_size)
		f.close()
	return res


def print_results(res):
	res.insert(0, remove_extra_cols(
		["EventID", "Date", "Time", "OSProcID", "SAPProcID", "TypeConn", "SAPIDHEX", "Terminal",
		"Login", "T-Code", "Report", "Client", "SessionId", "Parameters", "FQDN"]))
	for row in res:
		for col in row:
			print(f"{col}", end="\t")
		print()


def excel_export(res, filename):
	df = pd.DataFrame(res)
	writer = pd.ExcelWriter(f'{filename}.xlsx', engine='xlsxwriter')
	df.to_excel(writer, sheet_name=f'{filename}', index=False)
	writer.close()


def main():
	input_files = set()
	for name in parsed_args.source_names:
		if os.path.isfile(name):
			input_files.add(os.path.abspath(name))
		elif os.path.isdir(name):
			for path, dirs, files in os.walk(name):
				for file in files:
					current_file = os.path.join(path, file)
					if '.AUD' in current_file:
						input_files.add(os.path.abspath(current_file))
	
	print("We start the parsing with: ", input_files)
	if parsed_args.remove:
		print("Cols for remove: ", parsed_args.remove)
	if parsed_args.terminal:
		print("Filter by terminal(in fields terminal and fqdn): ", parsed_args.terminal)
	if parsed_args.login:
		print("Filter by login: ", parsed_args.login)
	if parsed_args.tcode:
		print("Filter by T-Code(in fields tcode and param): ", parsed_args.tcode)
	if parsed_args.report:
		print("Filter by report: ", parsed_args.report)
	if parsed_args.client:
		print("Filter by client: ", parsed_args.client)
	result = []
	for file_name in input_files:
		result = result + parse_file(file_name)
	
	print("Collected rows:", len(result))
	if parsed_args.print:
		print_results(result)
	else:
		print("For printing on display plz use -print option")
	
	if parsed_args.excel:
		print("Export to excel is started...")
		excel_export(result, parsed_args.excel_file_name)
		print("and was finished!")
	else:
		print("For exporting to excel plz use -excel option")


if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument('-remove', metavar='eventid osid sapid sapidhex sessionid', nargs='*',
						help='You can specify cols for remove. If you use this option, default removing will be '
							'overwritten. Other cols: date time typecon terminal login tcode report client param '
							'fqdn',
						default=['eventid', 'osid', 'sapid', 'sapidhex', 'sessionid'])
	parser.add_argument('-terminal', metavar='31709', nargs='*',
						help='It tries to search this in fields terminal and fqdn, as substring')
	parser.add_argument('-login', metavar='SAPROOT CUASM7', nargs='*',
						help='It tries to search this in field login, as substring')
	parser.add_argument('-tcode', metavar='SU01 PFCG', nargs='*',
						help='It tries to search this in fields tcode and param, as substring')
	parser.add_argument('-report', metavar='ZBC', nargs='*',
						help='It tries to search this in field report, as substring')
	parser.add_argument('-client', metavar='000 300', nargs='*',
						help='It tries to search this in field client, as substring')
	parser.add_argument('-print', help='Print all', action='store_true')
	parser.add_argument('-excel', help='Enable export to excel with default name results.xlsx', action='store_true')
	parser.add_argument('-excel_file_name', metavar='results', help='use this name for excel file', default="results")
	parser.add_argument('source_names', nargs='*', help='parse all *.AUD from this directory or file, "./" by default',
						default=".")
	parsed_args = parser.parse_args()
	main()
