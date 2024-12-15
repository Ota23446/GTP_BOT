import pythoncom
import win32com.client
import requests
import os

pythoncom.CoInitialize()
url = "http://confluence.jira.lan:8090/exportword?pageId=24577712"
response = requests.get(url, verify=False)
response.raise_for_status()

with open("test.doc", "wb") as f:
    f.write(response.content)

word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open("test.doc")
tables = doc.Tables
print(f"Number of tables: {tables.Count}")
doc.Close()
word.Quit()
pythoncom.CoUninitialize()