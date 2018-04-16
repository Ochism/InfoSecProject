# Spam Classifier Outlook Add-in
This software application is a spam email classifier that protects its users from potentially harmful phishing or spam emails. It is implemented as a Microsoft Outlook Add-in that gets called whenever the user receives a new mail. The Add-in will classify the mail as either SPAM or NOT SPAM, prepend its classification and confidence to the email's subject, then move that email into its appropriate folder (Inbox or WatsonSpam).

## Requirements
* Visual Studio 2017
* Microsoft Office 365 account
* Microsoft Outlook
* .NET Framework 4.6.*
* [IBM Watson Natural Language Classifier](https://github.com/watson-developer-cloud/dotnet-standard-sdk/tree/development/src/IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1)

## Getting Started
Open up the __SpamClassifier.sln__ Visual Studio solution file on a Windows computer. This will open Visual Studio and build the project. All Add-in code written by the team is located in __SpamClassifier/ThisAddIn.cs__.

_**NOTE:** The solution file will not automatically install the IBM Watson Natural Language Classifier package. This can be done with the following command issued in the NuGet console:_
```

PM > Install-Package IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1

```
## Components

### Classifiers
Two classifiers were trained using IBM Watson's Natural Language Classifier service. The classifiers were trained using an [online corpus](http://www.csmining.org/index.php/spam-email-datasets-.html) of 4327 emails that were split into 80% training data and 20% testing data. One classifier was responsible for classifying the subjects of emails and the other was used for the email bodies.

#### Subject Classifier
* 92.96% accuracy
* 97.79% average confidence

#### Body Classifier
* 94.77% accuracy
* 95.55% average confidence

The creation, training and testing of these classifiers was done by Kurtis Kuszmaul. Code for these processes can be found in the __ClassifierCreation__ directory.

### Outlook Add-in
The Outlook Add-in runs in the background of Outlook and fires whenever new mail is received. It locates the new mail item, extracts the subject and body from it, then classifies those two text fields using the classifiers explained above. The confidence of the classifications is weighted and compared to determine a final classification and confidence level. This classification and confidence percentage is prepended to the subject, then the appropriate action is taken on the email.

#### Subject Class = Body Class
* Classification done based on weighted sum of subject and body classifier confidence
* Requires 85% confidence to keep classification

#### Subject Class != Body Class
* Classification of the higher of the two weighted confidences taken
* Requires 95% confidence to keep classification

## External Components (Not Developed by Team)
* [IBM Watson Natural Language Classifier](https://github.com/watson-developer-cloud/dotnet-standard-sdk/tree/development/src/IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1)
  * Used for custom classifications of text
* [Visual Studio Tools for Office](https://docs.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins)
  * Used for integrating custom Add-in functionality with Microsoft Outlook

## Contributors
* Ethan Knez
* Kurtis Kuszmaul
* Greg Ochs
