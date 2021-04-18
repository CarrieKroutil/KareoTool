# Kareo Tool

[Kareo](http://kareo.com) is an Electronic Health Record and Practice Management Software Application.

They provide an application programming interface via a WSDL, and documentation can be found: https://helpme.kareo.com/01_Kareo_PM/12_API_and_Integration.

## Overview

The Kare Tool was built to allow for data to be retrieved from Kareo's API to create for a more fluid accounting workflow.

Additional steps can be taken once the data has been exported into Excel spreadsheets via Power Query.

## Generate Executable

Build the solution in release mode to generate the executable file, found in `\bin\Release`.

## How to Export API Data

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

## Output

The results of running the tool create Excel spreadsheets inside the "Output" folder for each enabled endpoint located in the directory the executable lives.
