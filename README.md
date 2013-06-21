Automated-Data-Entry
====================

Automated data entry into Excel files using curl-PHP for downloading the data from the web.



I have used it for collecting a data arranged into tabular form on webpage and the webpages could be retrieved by iterating with a simple GET parameter. :)

CURL: Used to download the HTML source data of webpages given the GET parameters.
Simple_HTML_DOM: PHP library used for parsing the html data.
PHPExcel: PHP Library used for exporting the data into excel sheets.


The curl configrations used are for a HTTPS based site. You can modify it according to your needs.
For additional information on curl and its parameters, google 'CURL'.

I have used simple_html_dom library for html parsing. I have looped through the tables on the page to extract the data accordingly 
and transfer it to appropriate excel cells.


[Note: This will only work for data sources on the web with 'GET' method.]
[Note: The code has lot of commented code. Please ignore or use it for reference.]



