#Importando bibliotecas para web scrapy
import requests
from bs4 import BeautifulSoup

#Biblioteca para seleção de arquivo armazenado localmente
import os

#Biblioteca para gerar epub
from ebooklib import epub

#Biblioteca para enviar por email
import win32com.client as win32

# Dados constantemente editáveis
base_url = "https://site-srapy-to-be-chapter-"  #url padrão do site que quer fazer o scrapy antes do número da página
start_index = 1
end_index = 75

#Faz o loop do scraping de conteúdo e armazena em uma varíavel se a requisição não der erro
def scrape_wordpress_multiple(base_url, start_index, end_index):
    def scrape_wordpress(url):
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            #Selecionando todo o texto que se encontra em paragrafos
            paragraphs = soup.find_all('p')
            text_list = [paragraph.get_text() for paragraph in paragraphs]
            
            # Adicionando quebra de linha após cada parágrafo
            for i in range(len(text_list)):
                text_list[i] += '\n'
            return ''.join(text_list)
        
        #Se a requisição não der certo, procure o erro através do código de Status
        else:
            print("Falha ao fazer a solicitação. Código de status:", response.status_code)
            return None

    #Loop para que o mesmo procedimento seja feito em todas as páginas estabelecidas no código
    combined_text = ""
    for i in range(start_index, end_index + 1):
        url = f"{base_url}{i}/"
        print(f"Scraping {url}")
        page_text = scrape_wordpress(url)
        if page_text:
            combined_text += page_text + '\n\n'
    return combined_text

# Função para criar um arquivo .epub com base no texto combinado
def create_epub_file(text, filename='Nome que quiser.epub'):
    book = epub.EpubBook()
    book.set_identifier('Autor')
    book.set_title('Título')
    book.set_language('en')
    
    # Dividindo o texto em parágrafos
    paragraphs = text.split('\n')
    
    # Criando o conteúdo XHTML com parágrafos
    content = ''
    for paragraph in paragraphs:
        content += f'<p>{paragraph}</p>'
    
    # Adicionando o conteúdo ao livro .epub
    book.add_item(epub.EpubHtml(title='Título', file_name='index.xhtml', content=content, lang='en'))
    book.toc = [epub.Link('index.xhtml', 'Texto de Navegação', 'ID')] 
    #book.toc = [epub.Link('index.xhtml')]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    # Escrevendo o arquivo .epub
    epub.write_epub(filename, book, {})

combined_text = scrape_wordpress_multiple(base_url, start_index, end_index)

# Criando um arquivo .epub com base no texto combinado
create_epub_file(combined_text, 'Title that will appear on your kindle.epub')    

#Opção para já mandar direto pro seu kindle através do do seu e-mail kindle usando o outlook (só funciona se seu pacote office for original)
'''outlook = win32.Dispatch('Outlook.Application')
olNS = outlook.GetNameSpace('MAPI')

mail = outlook.CreateItem(0)
mail.to = 'seu_email@kindle.com'
mail.Subject = 'Assunto'
mail.BodyFormat = 1
mail.Body = 'Opcional.'
mail.Attachments.Add(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'nome_do_arquivo.extensão'))

mail._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('seu_email_pessoal@outlook.com')))

#Descomente a linha abaixo se quiser visualizar o corpo do e-mail sendo enviado
#mail.Display()

mail.Send()'''
