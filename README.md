# ScholarshipRequestAutomation

The purpose of this project is to automate the applications for merit scholarships in a romanian university.

As Is: students are filling merit scholarship templates in person with standard information such as: overrall grade in the last semester, name, student id, 
identity card information.

To be: students are given access to a given word document template which needs to be filled online. The templates are sent by the students to a given gmail address
and the python script reads all "New" mails, checks the attachments, downloads them and verifies that all information is filled. It then validates the information and
inserts it into a PostgreSQL database.
