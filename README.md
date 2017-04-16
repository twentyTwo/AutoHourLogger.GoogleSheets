## Googlesheet Auto Hour Logger
This application automatically log hours in everyday basis to a google sheet.



## Step 1: Turn on the Google Sheets API

1. Use  [this wizard](https://console.developers.google.com/start/api?id=sheets.googleapis.com) to create or select a project in the Google Developers Console and automatically turn on the API. Click  **Continue** , then  **Go to credentials**.
2. On the  **Add credentials to your project**  page, click the  **Cancel**  button.
3. At the top of the page, select the  **OAuth consent screen**  tab. Select an  **Email address** , enter a  **Product name** if not already set, and click the  **Save**  button.
4. Select the  **Credentials**  tab, click the  **Create credentials**  button and select  **OAuth client ID**.
5. Select the application type  **Other** , enter the name &quot;Google Sheets API Quickstart&quot;, and click the  **Create**  button.
6. Click  **OK**  to dismiss the resulting dialog.
7. Click the file\_download (Download JSON) button to the right of the client ID.
8. Move this file to your working directory and rename it client\_secret.json.

## Step 2: Prepare the project

1. Create a new Visual C# Console Application project in Visual Studio.
2. Open the NuGet Package Manager Console, select the package source  **nuget.org** , and run the following command:

Install-PackageGoogle.Apis.Sheets.v4

## Step 3: Set up the sample

1. Drag client\_secret.json (downloaded in Step 1) into your Visual Studio Solution Explorer.
2. Select client\_secret.json, and then go to the Properties window and set the  **Copy to Output Directory**  field to  **Copy always**.
