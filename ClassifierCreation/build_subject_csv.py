'''
This file builds a subject .csv file of subject text and its classification.
'''
__author__ = 'Kurtis Kuszmaul'

import email.parser
import csv
import argparse

def extract_subject(filename):
	'''
	Extract the subject from the .eml file.
	'''
	fp = open(filename)
	msg = email.message_from_file(fp)
	sub = msg.get('subject')
	sub = str(sub)
	sub = ' '.join(sub.split())
	sub = 'no subject' if len(sub) == 0 else sub
	return sub

parser = argparse.ArgumentParser()
parser.add_argument('--emails', required=True,
    	 help='Directory that contains the emails to process')
parser.add_argument('--label', required=True,
		 help='Label file for classification')
parser.add_argument('--output', required=True,
    	 help='Name of the output csv file')
args = parser.parse_args()

print(args.emails)
email_dir = args.emails if args.emails[-1] == '/' else args.emails + '/'
label = args.label
output = args.output if args.output[-4:] == '.csv' else args.output + '.csv'
classification_map = {'0': 'spam', '1': 'not spam'}

with open(output, 'w') as classifier:
	csv_writer = csv.writer(classifier, dialect='excel')

	with open(label, 'r') as label:
		for line in label:
			parts = line.split(' ')
			classification = parts[0]
			filename = parts[1].strip('\n')
			filename = email_dir + filename
			subject = extract_subject(filename)
			csv_writer.writerow((subject, classification_map[classification]))
			print('row written')
