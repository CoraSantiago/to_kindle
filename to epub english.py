#Importing libraries for web scraping.
import requests
from bs4 import BeautifulSoup

#Library for selecting locally stored file.
import os

#Library for generating epub.
from ebooklib import epub

#Library for sending via email
import win32com.client as win32

# Data constantly editable
base_url = "https://site-srapy-to-be-chapter-"  #Default URL of the website you want to scrape before the page number.
start_index = 1
end_index = 75

#Loops through the scraping of content and stores it in a variable if the request doesn't error out.
def scrape_wordpress_multiple(base_url, start_index, end_index):
    def scrape_wordpress(url):
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            #Selecting all the text found within paragraphs
            paragraphs = soup.find_all('p')
            text_list = [paragraph.get_text() for paragraph in paragraphs]
            
            # Adding a line break after each paragraph.
            for i in range(len(text_list)):
                text_list[i] += '\n'
            return ''.join(text_list)
        
        #If the request fails, look for the error using the status code.
        else:
            print("Failed to make the request. Status code:", response.status_code)
            return None

    #Loop so that the same procedure is performed on all pages established in the code
    combined_text = ""
    for i in range(start_index, end_index + 1):
        url = f"{base_url}{i}/"
        print(f"Scraping {url}")
        page_text = scrape_wordpress(url)
        if page_text:
            combined_text += page_text + '\n\n'
    return combined_text

# Function to create a .epub file based on the combined text.
def create_epub_file(text, filename='Any name you like.epub'):
    book = epub.EpubBook()
    book.set_identifier('Author')
    book.set_title('Title')
    book.set_language('en')
    
    # Splitting the text into paragraphs.
    paragraphs = text.split('\n')
    
    # Creating XHTML content with paragraphs.
    content = ''
    for paragraph in paragraphs:
        content += f'<p>{paragraph}</p>'
    
    # Adding the content to the .epub book.
    book.add_item(epub.EpubHtml(title='Title', file_name='index.xhtml', content=content, lang='en'))
    book.toc = [epub.Link('index.xhtml', 'Navigation Text', 'ID')]
    #book.toc = [epub.Link('index.xhtml')]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    # Writing the .epub file.
    epub.write_epub(filename, book, {})

combined_text = scrape_wordpress_multiple(base_url, start_index, end_index)

# Creating an .epub file based on the combined text.
create_epub_file(combined_text, 'Title that will appear on your kindle.epub')    

#Option to directly send to your Kindle through your Kindle email using Outlook (only works if your Office package is genuine)
'''outlook = win32.Dispatch('Outlook.Application')
olNS = outlook.GetNameSpace('MAPI')

mail = outlook.CreateItem(0)
mail.to = 'you_email@kindle.com'
mail.Subject = 'Assunto'
mail.BodyFormat = 1
mail.Body = 'Opcional.'
mail.Attachments.Add(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'file_name.epub'))

mail._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('your_personal_email@outlook.com')))

# Uncomment the line below if you want to preview the body of the email being sent.
#mail.Display()

mail.Send()'''
