import docx
import os
import datetime
import psycopg2
import exceptions

os.chdir('C:\\Users\\nemes\\OneDrive\\Desktop\\TestDirectory\\attachments')
curr_date = datetime.date.today().strftime("%d %b %Y")

searched_key_words = ['Prenume:', 'Nume:', 'Telefon:', 'E-mail:', 'Sectia:', 'Profil:', 'Media semestrul curent:', 'Media semestrul precedent (daca este cazul):']

# Read file and extract necessary information that follows key words from list, remove whitespace
def getText(filename):
    doc = docx.Document(filename)
    extracted_info = []
    for para in doc.paragraphs:
        for key_word in searched_key_words:
            if key_word in para.text:
                extracted_info.append(para.text.split(key_word, 1)[1].strip())

    return extracted_info

def InsertData(extracted_info):
    # connect to db
    conn = psycopg2.connect(f'dbname=IngineriaSistemelor user=postgres password={os.environ["SQLPASS"]}')
    cursor = conn.cursor()
    # remove null values
    extracted_info = list(filter(None, extracted_info))
    # insert data without previous_gpa in case of students that are in the first semester
    if len(extracted_info) == 7:
        sql_query = f"INSERT INTO student_scholarships(first_name, last_name, phone_number, email, segment, profile, " \
                f"current_gpa) VALUES (%s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(sql_query, extracted_info)
        conn.commit()
    # insert data with previous_gpa in case of students that are in semesters > 1
    elif len(extracted_info) == 8:
        sql_query = f"INSERT INTO student_scholarships(first_name, last_name, phone_number, email, segment, profile, current_gpa, previous_gpa) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(sql_query, extracted_info)
        conn.commit()
    else:
        print("Incorrect number of fields")
        raise exceptions.MissingValues


def GetAndInsertData(filePath):
    fileName, file_extension = os.path.splitext(filePath)
    if file_extension == '.docx' or file_extension == '.doc':
        necessary_info = getText(filePath)
        InsertData(necessary_info)
    else:
        print("The only accepted types of documents are .docx and .doc")
        raise exceptions.WrongTypeOfDocument






