# EPA Grant Scraper
This tool was made to quickly parse through all the downloadable pdfs from an annual excel file that the EPA produces. The documents are both the recipients and non-recipients of grants for that year by the EPA. This program parses, uses timed http requests (so as not to DDOS, etc) to download the files from the EPA, and renames and organizes all of the files. There is some basic validations for paths and warnings/logs for errors that could occur due to bad responses during the http cycle. The console has had some color treatment so as to identify successful and non-successful outputs accordingly.

The next step is to convert this into an executable rather than having to run the program from an IDE.
