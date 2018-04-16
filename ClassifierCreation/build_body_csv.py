'''
This file builds a .csv file of body text and its classification.
'''
__author__ = 'Kurtis Kuszmaul'

import email.parser
import csv
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--bodies', required=True,
            help='Directory that contains the bodies to process')
parser.add_argument('--label', required=True,
            help='Label file for classification')
parser.add_argument('--output', required=True,
            help='Name of the output csv file')
args = parser.parse_args()

print(args.bodies)
body_dir = args.bodies if args.bodies[-1] == '/' else args.bodies + '/'
label = args.label
output = args.output if args.output[-4:] == '.csv' else args.output + '.csv'
classification_map = {'0': 'spam', '1': 'not spam'}

with open(output, 'w') as classifier:
    csv_writer = csv.writer(classifier, dialect='excel')

    with open(label, 'r') as label:
        row_nbr = 0
        for line in label:
            parts = line.split(' ')
            classification = parts[0]
            filename = parts[1].strip('\n')
            filename = body_dir + filename
            with open(filename, 'r') as body_file:
                try:
                    for chunk in body_file:
                        try:
                            row_nbr = row_nbr + 1
                            chunk = chunk.replace('\n', '')
                            csv_writer.writerow((chunk, classification_map[classification]))
                            print('row #' + str(row_nbr) + ' written')
                        except:
                            pass
                except:
                    pass

