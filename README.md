<img src="https://github.com/MatthewHana/OutlookDSD/assets/1935851/d1abf019-e0b8-4499-800d-4a59afce4410" width="450px">

![C#](https://img.shields.io/badge/Language-C%23-green) ![Status](https://img.shields.io/badge/Status-Release-red) ![Ver](https://img.shields.io/badge/Version-1.0.1-blue)

OutlookDSD - A simple add-in for **Outlook** that shows the **D**KIM, **S**PF and **D**MARC status of an email message in the ribbon.

## About & Why?
Outlook doesn't alert you when an email has failed DKIM, SPF or DMARC validation. This could create security risks by allowing users to fall victim to phishing or other attacks.

OutlookDSD adds visual indicators displaying the validation results of the DKIM, SPF and DMARC status of received emails, and warning users if a message fails any security validation.


<img src="https://github.com/MatthewHana/OutlookDSD/assets/1935851/8674eead-722b-4219-97d8-a4754e24fe89" width="800px">

*Valid Email*

<img src="https://github.com/MatthewHana/OutlookDSD/assets/1935851/41327802-4d8c-4505-b737-b994cfa01082" width="800px">

*Phishing Email*


## Install
To install the Add-in Simply download the latest package Zip file and run the setup.exe file.

### Requirements 
OutlookDSD requies the following:
* [Outlook for Desktop](https://support.microsoft.com/en-us/office/download-and-install-or-reinstall-microsoft-365-or-office-2021-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)
* [.NET Framework 4.7.2 (or higher)](https://support.microsoft.com/en-us/topic/microsoft-net-framework-4-7-2-offline-installer-for-windows-05a72734-2127-a15d-50cf-daf56d5faec2)
* [Visual Studio 2010 Tools for Office Runtime](http://go.microsoft.com/fwlink/?linkid=140384)

## Uninstall
To uninstall OutlookDSD simply use the windows Add or Remove Programs function.

## Download
[Download the latest package v1.0.1.zip from here](https://).

## Configuration
OutlookDSD can be configured from the 'Add-in Options dialog'  in Outlook.

To open the dialog click **File** on the ribbon and select **Options**.
<img src="https://github.com/MatthewHana/OutlookDSD/assets/1935851/5c8988eb-084d-4fe5-8997-1674202c2187" width="800px">

Select **Add-ins** from the left hand, and click the **Add-in Options...** button.
<img src="https://github.com/MatthewHana/OutlookDSD/assets/1935851/f2f683d0-1b9e-46ae-a024-21a71ad23a60" width="800px">
