SPAM CLASSIFIER OUTLOOK ADD-IN
================================================
FULL PROJECT CODE AND INSTALLATION INSTRUCTIONS CAN BE FOUND ON GITHUB AT https://github.com/Ochism/InfoSecProject

This software application is a spam email classifier that protects its users from potentially harmful
phishing or spam emails. It is implemented as a Microsoft Outlook Add-in that gets called whenever the user
receives a new mail. The Add-in will classify the mail as either SPAM or NOT SPAM, prepend its classification
and confidence to the email's subject, then move that email into its appropriate folder (Inbox or WatsonSpam).

COMPONENTS
================================================
Classifiers

  Two classifiers were trained using IBM Watson's Natural Language Classifier service. The classifiers were
  trained using an online corpus of 4327 emails that were split into 80% training data and 20% testing data.
  One classifier was responsible for classifying the subjects of emails and the other was used for the email
  bodies.
  
  Subject Classifier
    - 92.96% Accuracy
    - 97.79% Average Confidence
  Body Classifier
    - 94.77% Accuracy
    - 95.55% Average Confidence
    
  The creation, training and testing of these classifiers was done by Kurtis Kuszmaul. Code for these processes
  can be found in the ClassifierCreation directory.
  
  The email corpus can be found at http://www.csmining.org/index.php/spam-email-datasets-.html

Outlook Add-in

  The Outlook Add-in runs in the background of Outlook and fires whenever new mail is received. It locates the
  new mail item, extracts the subject and body from it, then classifies those two text fields using the
  classifiers explained above. The confidence of the classifications is weighted and compared to determine a
  final classification and confidence level. This classification and confidence percentage is prepended to the
  subject, then the appropriate action is taken on the email.
  
  Subject Class = Body Class
    - Classification done based on weighted sum of subject and body classifier confidence
    - Requires 85% confidence to keep classification
  Subject Class != Body Class
    - Classification of the higher of the two weighted confidences taken
    - Requires 95% confidence to keep classification
    
  The design and development of this Outlook Add-in was done by Gregory Ochs, Ethan Knez, and Kurtis Kuszmaul.
  Code for these the add-in can be found in the SpamClassifier directory.
  
EXTERNAL COMPONENTS (NOT DEVELOPED BY TEAM)
================================================
IBM Watson Natural Language Classifier - https://github.com/watson-developer-cloud/dotnet-standard-sdk/tree/development/src/IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1
  - Used for custom classifications of text

Visual Studio Tools for Office - https://docs.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins
  - Used for integrating custom Add-in functionality with Microsoft Outlook

CONTRIBUTORS
================================================
- Ethan Knez
- Kurtis Kuszmaul
- Gregory Ochs
