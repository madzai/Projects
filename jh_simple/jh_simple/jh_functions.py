import os
import shutil
from shutil import copy
import webbrowser
import pandas as pd
from openpyxl import Workbook, load_workbook
# import pdfkit
from bs4 import BeautifulSoup
import requests

def open_file(path):
    # open file if file exists
    if os.path.exists(path):
        path = os.path.realpath(path)
        os.startfile(path)
    
def copy_rename_file(src_path, dest_path):
    copy(src_path, dest_path)
    
def open_url(link):
    webbrowser.open(link)
    
# def print_pdf_url(url, pdf_path):
#     #Define path to wkhtmltopdf.exe
#     path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

#     #Point pdfkit configuration to wkhtmltopdf.exe
#     config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

#     #Convert Webpage to PDF
#     pdfkit.from_url(url, output_path=pdf_path, configuration=config)

def jobsDB(url):
    # html_content = "webpage.html"
    # soup = BeautifulSoup(html_content, 'html.parser')
    
    response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
    soup = BeautifulSoup(response.content, 'html.parser')
    job_title = soup.find('h1', {'data-automation': 'job-detail-title'}).text.strip()
    company = soup.find('span', {'data-automation': 'advertiser-name'}).text.strip()
    
    # company = soup.find('span', {'data-automation': 'advertiser-name'})
    # company = soup.find('button', {'class': 'a3yfdf0 a3yfdf7 _5xlhblh _4t0edd0'}).text
    
    return job_title, company

def job_and_company():
    wb = load_workbook("jobs - new.xlsx")
    # active worksheet
    ws = wb.active
    df = pd.read_excel("jobs - new.xlsx")
    # get urls from column with header URL
    list_url = df['URL'].tolist()
    list_company = []
    list_title = []
    
    # n = len(list_url)
    # for i in range(n):
    #     url = list_url[i]
    #     job_title, company = jobsDB(url)
    #     # paste job title and company on columns B & C

    #     # folder_name = company + " - " + job_title
    #     # # ws.cell(row=i+1, column=4).value = folder_name
    #     # print(folder_name)
    
    for i in range(len(list_url)):
        url = list_url[i]
        job_title, company = jobsDB(url)
        
        list_company.append(company)
        list_title.append(job_title)
        # paste job title and company on columns B & C

        # folder_name = company + " - " + job_title
        # # ws.cell(row=i+1, column=4).value = folder_name
        # print(folder_name)
        
    start_row = 2
    for index, (title, comp) in enumerate(zip(list_title, list_company), start=start_row):
        ws[f'B{index}'] = title
        ws[f'C{index}'] = comp
        # ws[f'H{index}'] = item
    
    # save the Excel file
    wb.save('jobs - new.xlsx')
####################################################################
def job_folders():
    df = pd.read_excel("jobs.xlsx")
    list_url = df['URL'].tolist()
    # list_company = df['Company'].tolist()
    # list_post = df['Post'].tolist() 
    
    n = len(list_url)
    # folders
    for i in range(n):
        # folders
        directory_name = "E:/Projects/jh_simple/job_apply/"
        url = list_url[i]
        job_title, company = jobsDB(url)
        folder_name = company + " - " + job_title
        path_name = directory_name + folder_name
        print(folder_name)
        # create folder
        os.mkdir(path_name)
        # copy template files (CV, cover letter) to folder (job_apply)
        template_files = ["Cover Letter (Edwin Ma).docx", "CV (Edwin Ma).docx"]
        for template_file in template_files:
            src_path = "E:/Projects/jh_simple/job_template/" + template_file
            dest_path = path_name
            copy_rename_file(src_path, dest_path)
        # cover letter
        header = "RE: Application for the position of " + job_title
        intro = "I am writing to apply for the position of " + job_title + ", as advertised on jobsdb.com."
        print(header)
        print(intro)
        print()
        # open job page URL
        open_url(list_url[i])
        
        # print pdf - https://pyshark.com/convert-html-to-pdf-using-python/#convert-webpage-to-pdf-using-python
        # pdf_filename = list_company[i] + " - " + list_post[i] + ".pdf"
        # pdf_path = "C:/Users/edwma/Python_learning/Projects/jh_simple/job_apply/" + folder_name + "/" + pdf_filename
        # print_pdf_url(list_url[i], pdf_path)
        
# b4 web srape job title and company name
def job_details(incl_letter=False, open_web=False):
    df = pd.read_excel("jobs.xlsx")

    list_company = df['Company'].tolist()
    list_post = df['Adj'].tolist()
    list_url = df['URL'].tolist()
    
    n = len(list_company)
    # folders
    for i in range(n):
        # create folders
        directory_name = "E:/Projects/jh_simple/job_apply/"
        folder_name = list_company[i] + " - " + list_post[i]
        path_name = directory_name + folder_name
        print(folder_name)
        # create folder
        os.mkdir(path_name)
        # copy template files (CV, cover letter) to folder (job_apply)
        # template_files = ["Cover Letter (Edwin Ma).docx", "CV (Edwin Ma).docx"]
        template_files = ["CV (Edwin Ma).docx"]
        if incl_letter:
            template_files.append("Cover Letter (Edwin Ma).docx");
        for template_file in template_files:
            src_path = "E:/Projects/jh_simple/job_template/" + template_file
            dest_path = path_name
            copy_rename_file(src_path, dest_path)
        # cover letter
        # header = "RE: Application for the position of " + list_post[i]
        # intro = "I am writing to apply for the position of " + list_post[i] + ", as advertised on jobsdb.com."
        # print(header)
        # print(intro)
        # print()
        
        # open job page URL
        if open_web:
            open_url(list_url[i])

def job_urls():
    df = pd.read_excel("jobs.xlsx")
    # df = pd.read_excel("excel.xlsx", sheet_name='Top Scorers')

    # list_company = df['Company'].tolist()
    # list_post = df['Adj'].tolist()
    list_url = df['URL'].tolist()
    
    for url in list_url:
        open_url(url)
       
# Test web scrapping job title & company from jobsDB 
# url = "https://hk.jobsdb.com/job/79371701?token=1~77d06a72-db3c-4675-8343-3416c29e784a&ref=post-apply-recs" 
# url = "https://hk.jobsdb.com/job/79641491?ref=search-standalone&type=standout&origin=showNewTab#sol=d798dbc25aa87903fdf2ade216dfad993432be52"
# # job_title = jobsDB(url)
# job_title, company = jobsDB(url)
# print(job_title)
# print(company)

# job_and_company()

# job_details()







