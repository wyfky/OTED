# OTED
Output Tools for Search Results of Erudition（愛如生）Database(OTED)

The OTED is developed for helping researcher get the search results and its full text of Erudition Database. It consists of three parts that is Web page for user data file upload, google sheets and Python processing program.

A google form is used for submitting the query keywords, Database name and email address to google sheets. [google form](https://goo.gl/forms/EmSCeso2QbZ8oCS02)

Html file and JavaScript file in folder of google sheets upload page is to upload the search results data, email address and Database name to google sheets. JS file is [jquery-csv library created by evanplaice](https://github.com/evanplaice/jquery-csv). In google sheets a google script for receiving and writing post data from html page is needed.

The file of erus-email-form.py is to read the user’s query keywords from google sheets and return the search results to user by email.

The file of readGazetteers_Feng.py is to read the search result record that user upload and return the full text by email.

**Note: When use the OTED, Please confirm to your library if have the data mining agreement with Erudition company. Otherwise it is probably illegal.

