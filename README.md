# ⭐ Why 

I volunteered through [Catchafire](https://www.catchafire.org/) on a [Technical Specification Development project with a total impact of $7,802](https://www.catchafire.org/impact/match/2614761/the-women-s-center-of-southeastern-michigan--technical-specification-development/). 

The project goal was to optimize the accounting workflow to reduce the level of effort needed to support the organization's growth of therapists and clients. The practice management software used, Kareo facilitates the data needed to perform accounting within Quickbooks. The ideal solution would export needed data to Excel with the least amount of effort and cost to automate and maintain.

Due to the old-school nature of Kareo’s API, it is not feasible to directly connect to Kareo's API from within Excel using Power Query. This challenge led to the need to build this custom tool to connect to Kareo’s API and download the desired data into local Excel workbooks to use.  This makes the solution not ideal due to requiring expertise to construct the tool and support it long term.  As I was building this tool, I learned about Kareo’s Custom Report option using a Microsoft Excel Add-In to hook into Power Query to access the data - which was the major roadblock I experienced when first trying to implement my solution. Check out the outlined [Preferred Solution](https://github.com/CarrieKroutil/KareoTool/blob/master/README.md#-preferred-solution) below over this Kareo Tool.

I learned about [Power Query](https://support.microsoft.com/en-us/office/about-power-query-in-excel-7104fbee-9e62-4cb9-a02e-5bfb1a6c536a) and [using it in Excel](https://support.microsoft.com/en-us/office/about-power-query-in-excel-7104fbee-9e62-4cb9-a02e-5bfb1a6c536a) through this process. 

> **Note**
> I highly recommend the book [M Is for (Data) Monkey: A Guide to the M Language in Excel Power Query](https://www.amazon.com/Data-Monkey-Guide-Language-Excel/dp/1615470344) if you are interested in learning more. The M language is small fraction of the book and it also explains how to import excel workbooks into the master workbook and automate the desired output. 

I have captured my lessons learned in this README to help others. Please ⭐ star my repo to let me know you found my sharing useful.

# 🎯 Kareo Overview

[Kareo](http://kareo.com) is an Electronic Health Record and Practice Management Software Application.

Kareo is now part of Tebra [announced in this article](https://www.tebra.com/tebra-story):
> The world of healthcare is evolving, and so are we. Kareo and PatientPop have joined forces as Tebra to support the connected practice of the future and modernize every step of the patient journey.

## 📊 Kareo's Data Structure

In order to automate the process and export the data, it is important to understand Kareo's data structure.

![image](https://user-images.githubusercontent.com/11277317/227247680-9438ac0e-bdc3-49a8-a731-2ebd2d5660d7.png)

The Kareo API Technical Guide is available for download from here: https://helpme.kareo.com/01_Kareo_PM/12_API_and_Integration.

In the API, encounter details tie all the APIs together with EncounterID and contain: 
- PatientID
- AppointmentID
- RenderingProviderID.

Then charges & payments are related to the encounter's "BatchNumber" value.

Transactions do not have any key to tie back to encounters, other than the PatientID and service date.    

## 📡 Kareo SOAP API

Kareo’s SOAP API uses the Web Services Description Language (WSDL) which is an XML-based interface description language used to describe the functionality offered by a web service. It is available via the following endpoint: https://webservice.kareo.com/services/soap/2.1/KareoServices.svc?singleWsdl. 

### 🔬 Wizdler Chrome Exension

If you install a Chrome extension called Wizdler, you can browse the API’s available requests and responses: https://chrome.google.com/webstore/detail/wizdler/oebpmncolmhiapingjaagmapififiakb?hl=en 

### 🧰 Calling API Example

If you want to try out the API, download [Postman](https://www.postman.com) or preferred tool and follow these steps:

1. Create a “POST” action pointing to https://webservice.kareo.com/services/soap/2.1/KareoServices.svc?singleWsdl
1. Set Headers:
    - Content-Type = text/xml; charset=utf-8
    - SOAPAction = http://www.kareo.com/api/schemas/KareoServices/GetPractices
1. Set Body:
    - Select radio buttom, “raw”
    - Choose “XML” from  dropdown describing type
    - Paste in the appropriate body based on the desired endpoint, in this example showing GetPractices, the body should contain the following, and **updating the three {...} values with your credentials (`<CustomerKey>{...}</CustomerKey><Password>{...}</Password><User>{...}</User>`) as well as pratice name**:

```xml
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetPractices xmlns="http://www.kareo.com/api/schemas/"><request xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><RequestHeader><ClientVersion>v1</ClientVersion><CustomerKey>{...}</CustomerKey><Password>{...}</Password><User>{...}</User></RequestHeader><Fields><Active>false</Active><AdministratorAddressLine1>false</AdministratorAddressLine1><AdministratorAddressLine2>false</AdministratorAddressLine2><AdministratorCity>false</AdministratorCity><AdministratorCountry>false</AdministratorCountry><AdministratorEmail>false</AdministratorEmail><AdministratorFax>false</AdministratorFax><AdministratorFaxExt>false</AdministratorFaxExt><AdministratorFullName>false</AdministratorFullName><AdministratorPhone>false</AdministratorPhone><AdministratorPhoneExt>false</AdministratorPhoneExt><AdministratorState>false</AdministratorState><AdministratorZipCode>false</AdministratorZipCode><BillingContactAddressLine1>false</BillingContactAddressLine1><BillingContactAddressLine2>false</BillingContactAddressLine2><BillingContactCity>false</BillingContactCity><BillingContactCountry>false</BillingContactCountry><BillingContactEmail>false</BillingContactEmail><BillingContactFax>false</BillingContactFax><BillingContactFaxExt>false</BillingContactFaxExt><BillingContactFullName>false</BillingContactFullName><BillingContactPhone>false</BillingContactPhone><BillingContactPhoneExt>false</BillingContactPhoneExt><BillingContactState>false</BillingContactState><BillingContactZipCode>false</BillingContactZipCode><CreatedDate>false</CreatedDate><Email>false</Email><Fax>false</Fax><FaxExt>false</FaxExt><ID>false</ID><LastModifiedDate>false</LastModifiedDate><NPI>false</NPI><Notes>false</Notes><Phone>false</Phone><PhoneExt>false</PhoneExt><PracticeAddressLine1>false</PracticeAddressLine1><PracticeAddressLine2>false</PracticeAddressLine2><PracticeCity>false</PracticeCity><PracticeCountry>false</PracticeCountry><PracticeName>false</PracticeName><PracticeState>false</PracticeState><PracticeZipCode>false</PracticeZipCode><SubscriptionEdition>false</SubscriptionEdition><TaxID>false</TaxID><WebSite>false</WebSite><kFaxNumber>false</kFaxNumber></Fields><Filter><Active i:nil="true"/><FromCreatedDate i:nil="true"/><FromLastModifiedDate i:nil="true"/><ID i:nil="true"/><NPI i:nil="true"/><PracticeName>{...}</PracticeName><TaxID i:nil="true"/><ToCreatedDate i:nil="true"/><ToLastModifiedDate i:nil="true"/></Filter></request></GetPractices></s:Body></s:Envelope>
```

# ⚔️ Kareo Tool Overview

This Kareo Tool was built to allow for data to be retrieved from Kareo's API to create for a more fluid accounting workflow.

Additional steps can be taken once the data has been exported into Excel spreadsheets via Power Query. 

![image](https://user-images.githubusercontent.com/11277317/227238460-88cd3a88-be82-4c32-b5dc-1cc4f1cd5aa9.png)

Proposed Steps Explained:
1. The accountant’s personal Windows computer with access to Kareo’s Billing desktop application has another process set up to run on a scheduled daily basis.
1. The process runs this custom application that calls Kareo’s API. 
1. Requested data is received from Kareo’s API.
1. The custom application then downloads the data into individual Excel spreadsheets containing the desired data in a simple format.
1. On the accountant’s PC, is their special Excel workbook that uses Power Query to import the data from step 4 and then automates the desired steps to achieve a reusable process. 

The power of exporting data as the solution relies first on a consistent output of data from Kareo that is automated to avoid consuming time to be manually performed. Second, and most critical is the ability to use Excel’s Power Query capabilities to automate the manual steps performed today.  This will reduce workload and also provide peace of mind that the process is sustainable long term.

## ⚙️ Generate Executable

Build the solution in release mode to generate the executable file, found in `\bin\Release`.

## 📌 How to Export API Data

In order to use this tool, you must first have a Customer Key, an API username and API password.  These values are settings configured in the App.config file:

``` xml
    <add key="CustomerKey" value="..." />
    <add key="ApiUser" value="..." />
    <add key="ApiPassword" value="..." />
```

Then configure the desired endpoints to run via more settings in the App.config aka `PraticeManagementExporter.exe.config`:

``` xml
    <add key="EnableProviders" value="true" />
    <add key="EnablePatients" value="false" />
```

Once the settings are in place, the tool is free to run the `PraticeManagementExporter.exe` file alongside the `PraticeManagementExporter.exe.config` from any Windows machine.

## 📂 Output

The results of running the tool create Excel spreadsheets inside the "Output" folder for each enabled endpoint located in the directory the executable lives.

# 🚀 Preferred Solution

Instead of using this tool, I suggest [Kareo's Custom Reports](https://helpme.kareo.com/01_Kareo_PM/11_Run_Reports_and_Analytics/10_Custom_Reports/About_Custom_Reports) are used.

- Check out this how to [Create Custom Reports](https://helpme.kareo.com/01_Kareo_PM/11_Run_Reports_and_Analytics/10_Custom_Reports/Create_Custom_Reports) example from Kareo.

This is the ideal option to consider, given it can achieve the following outcome:

![image](https://user-images.githubusercontent.com/11277317/227254921-3e4b75e7-76d7-4af6-8c0e-725279e8788c.png)

Proposed Custom Report Steps Explained:
1. The accountant’s personal Windows computer has admin Kareo credentials.
1. Excel’s Power Query is used to pull data from Kareo.
1. The data can be manipulated for the desired outcome in a new workbook.
