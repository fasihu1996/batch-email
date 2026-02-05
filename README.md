# Batch email sender

This project allows to send batch emails from a specified email server with personalised greetings, the same body and a 
possible attachement. Keeping possible rate limits in mind, a delay of 1 second between each sent email has been added.

## How to use
+ Clone the repository and add the required values to the `.env` file in the project directory
+ Make sure that the PDF folder contains the matching PDF files with the pattern of `lastname_firstname.pdf`
+ Customise the Excel spreadsheet path to point to your desired sheet
  + The expected column headers are "Last name", "First name" and "email"
+ Use `connection.py` to test the connetion to the SMTP server
+ Simply run `python ./send.py` to run the batch sending script.

