'''
This file tests Watson classifiers and reports its accuracy
and average confidence.
'''
__author__ = 'Kurtis Kuszmaul'

import json
import csv
import argparse
from watson_developer_cloud import NaturalLanguageClassifierV1

subject_nlc = NaturalLanguageClassifierV1(
  username='a477516a-4cdf-4080-93bc-064265ec1643',
  password='4JnCcEcxFDjM')
subject_classifier_id = '2fc15ax329-nlc-819'

body_nlc = NaturalLanguageClassifierV1(
         username='cd32418e-01b1-478e-9c24-a46a0767a0c7',
         password='AXISL3obSiSo'
)
body_classifier_id = 'ab2c7bx342-nlc-368'

parser = argparse.ArgumentParser()
parser.add_argument('--input', required=True,
    	 help='csv file used for testing')
parser.add_argument('--classifier', required=True,
		 help='classifier to use')
args = parser.parse_args()

# Assign proper classifier
if args.classifier == 'subject':
    classifier = subject_nlc
    classifier_id = subject_classifier_id
else:
    classifier = body_nlc
    classifier_id = body_classifier_id

with open(args.input, 'r') as subjects:
    csv_reader = csv.reader(subjects, dialect='excel')
    num_correct = 0.0
    conf_sum = 0.0
    total = 0.0

    try:
        for row in csv_reader:
            chunk = row[0]
            classification = row[1]

            # classify subject
            response = classifier.classify(classifier_id, chunk)
            top_class = response['top_class']
            confidence = response['classes'][0]['confidence']

            if top_class == classification:
                num_correct += 1
            conf_sum += confidence
            total += 1
            print('email ' + str(int(total)) + ' classified')
    except:
        pass

    accuracy = num_correct/total * 100
    avg_confidence = conf_sum/total * 100
    print('*********************************************')
    print('Total subjects processed: ' + str(int(total)))
    print('Total correctly classified: ' + str(int(num_correct)))
    print('Accuracy: ' + str(accuracy))
    print('Average confidence: ' + str(avg_confidence))
    print('*********************************************')
