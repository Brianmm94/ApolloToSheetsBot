# ApolloToSheetsBot
 Parses Discord Apollo bot events and places data from them into a Google Spreadsheet

Setup:
1. In your Google API Console under Credentials, create an "OAuth client ID" and select "Desktop app" for the application type.
2. Click on your new credential and select "DOWNLOAD JSON" at the top.
3. Rename the file to "credentials.json" and place it into the repository with the bot script.
4. Replace all of the empty strings in the ".env" file with valid values.
5. Run the script, and it should require you to perform OAuth2 authentication in your default browser before it starts placing values into the spreadsheet.
