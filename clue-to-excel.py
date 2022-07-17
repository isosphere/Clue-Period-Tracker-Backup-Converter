import argparse
import json
import logging
import sys
import xlsxwriter

parser = argparse.ArgumentParser(description="Generates an Excel-compatible file from a  .cluedata (JSON) file.")
parser.add_argument("input", type=str, help="Path to the .cluedata file")
# Forcing explicit output filenames (no default). I don't want to touch the filesystem without permission.
parser.add_argument("output", type=str, help="Path for the output file including the file name and extension ('.xlsx')")
parser.add_argument("--verbosity", type=str, default="INFO", help="The Python logging log level")

def parse_cluedata(input_path: str):
	"""Returns a Python structure representing the contents of the .cluedata file given.

	Args:
		input_path (str): the path to the .cluedata (JSON) input file

	Returns:
		dict: the data structure as parsed
	"""

	with open(input_path, 'r') as clue_backup:
		json_structure = json.loads(clue_backup.read())

	assert 'data' in json_structure, ".cluedata file does not have the expected structure"

	return json_structure

def identify_columns(structure: dict):
	"""Identifies the appropriate columns for the grid output by examining the input structure.

	Args:
		structure (dict): The data structure representing the .cluedata file

	Returns:
		list: The sequence of columns as they will be represented in the grid output
	"""
	known_columns = list()

	# collect columns
	for day in structure['data']:
		for key in day.keys():
			if key not in known_columns:
				known_columns.append(key)

	return known_columns

def generate_excel(structure:dict, output:str):
	"""Given a data structure representing a .cluedata file, generates an Excel-compatible file to `output`.

	Args:
		structure (dict): The data structure representing the .cluedata file
		output (str): The path that the output file will be written to, including file name.
	"""	

	structure_columns = identify_columns(structure)

	workbook = xlsxwriter.Workbook(output)
	worksheet = workbook.add_worksheet()

	col = 0
	for column in structure_columns:
		worksheet.write(0, col, column)
		col += 1

	row = 1
	for day in structure['data']:
		for key in day.keys():
			if isinstance(day[key], list):
				worksheet.write(row, structure_columns.index(key), ', '.join(day[key]))
			elif isinstance(day[key], dict):
				worksheet.write(row, structure_columns.index(key), str(day[key]))
			else:
				worksheet.write(row, structure_columns.index(key), day[key])
		row += 1
	
	worksheet.freeze_panes(1, 1)
	workbook.close()
	

if __name__ == '__main__':
	args = parser.parse_args()
	logging.basicConfig(stream=sys.stdout, level=args.verbosity)
	logger = logging.getLogger('cli')

	logger.info("Arguments accepted, attempting to parse file.")
	parsed_structure = parse_cluedata(input_path=args.input)
	logger.info("File parsed successfully, outputting file.")
	generate_excel(parsed_structure, args.output)
	logger.info("Complete.")