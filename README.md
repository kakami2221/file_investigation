import os 
import docx
import pdfplumber as ppdf
import warnings
import shutil

warnings.filterwarnings("ignore", category=UserWarning)


dir_path = input('捜査するディレクトリのパス : ')[1:-1]

n = 1
words = []
while n == 1:
    word = input('検索するワード : ')
    if word == '':
        break
    if word == 'BS':
        if words == []:
            print('選択された文字がありません')
            continue
        words.pop()

    words.append(word)

collect_file = []
collect_path = []

if os.path.exists(dir_path):
    for root,dirs,files in os.walk(dir_path):
        for file in files:
            file_path = os.path.join(root,file)
            name,ex = os.path.splitext(file)

            if ex == '.pdf':
                with ppdf.open(file_path) as pdf:
                    textp = ''
                    for page in pdf.pages:
                        textp += page.extract_text()

                    if not textp:
                        continue

                    np = 0
                    for word1 in words:
                        if word1 in textp:
                            np += 1
                            if np == len(words):
                                soutai_path = os.path.relpath(file_path,dir_path)
                                anp = [file,'<-' + soutai_path]
                                collect_file.append(anp)
                                collect_path.append(file_path)

            if ex == '.docx':
                document = docx.Document(file_path)
                
                for pages in document.paragraphs:
                    textw = pages.text
                
                for table in document.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            textw += cell.text

                nw = 0
                for word1 in words:
                    if word1 in textw:
                        nw += 1
                        if nw == len(words):
                            soutai_path = os.path.relpath(file_path,dir_path)
                            anw = [file,'<-'+ soutai_path]
                            collect_file.append(anw)
                            collect_path.append(file_path)



for keka in collect_file:
    print(keka)

print('ファイルをまとめたディレクトリを作りますか？')
endplace = input('y/n : ')

if endplace == 'y':
    makedir = input('ディレクトリをつくる場所のパス : ')[1:-1]
    dirname1 = '捜査結果'
    n1 = 1
    while n1 >= 1:
        dirname = dirname1 + f'{n1}'
        makedirplace = os.path.join(makedir,dirname)
        if os.path.isdir(makedirplace):
            n1 += 1
            continue
        else:
            os.makedirs(makedirplace)
            break

    for cfile in collect_path:
        shutil.copy(cfile,makedirplace)
