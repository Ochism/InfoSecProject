import json
import csv
from watson_developer_cloud import NaturalLanguageClassifierV1

subject_nlc = NaturalLanguageClassifierV1(
  username='a477516a-4cdf-4080-93bc-064265ec1643',
  password='4JnCcEcxFDjM')
subject_classifier_id = '2fc15ax329-nlc-819'

with open('testemails.csv', 'r') as subjects:
    csv_reader = csv.reader(subjects, dialect='excel')
    num_correct = 0.0
    conf_sum = 0.0
    total = 0.0

    for row in csv_reader:
        subject = row[0]
        classification = row[1]

        # classify subject
        response = subject_nlc.classify(subject_classifier_id, subject)
        top_class = response['top_class']
        confidence = response['classes'][0]['confidence']

        if top_class == classification:
            num_correct += 1
        conf_sum += confidence
        total += 1
        print('email ' + str(int(total)) + ' classified')

    accuracy = num_correct/total * 100
    avg_confidence = conf_sum/total * 100
    print('*********************************************')
    print('Total subjects processed: ' + str(int(total)))
    print('Total correctly classified: ' + str(int(num_correct)))
    print('Accuracy: ' + str(accuracy))
    print('Average confidence: ' + str(avg_confidence))
    print('*********************************************')
