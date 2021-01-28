# ScholarshipRequestAutomation

The purpose of this project is to automate the applications for merit scholarships in a romanian university.

As Is: students are filling merit scholarship templates in person with standard information such as: overrall grade in the last semester, name, student id, 
identity card information which is then handed to the university secretary signed, in person.

To be: students are given access to a given word document template which needs to be filled online. The templates are sent by the students to a given gmail address
and the python script reads all "New" mails, checks the attachments, downloads them and verifies that all information is filled. It then validates the information and
inserts it into a PostgreSQL database.

Several exceptions are handled:
1. There are no attachments in the mail: the script sends back an email to the student saying that the mail sent does not have an attachment
2. The attachment is not a word document type: the script sends back an email to the student saying that the document format is not correct and the given template should be used

If the information was inserted successfully the script sends back an email saying that the application has been successfully received.
