from itertools import takewhile
import urllib.request
import xlrd
import sys
import time

def col_len(sheet, index):
    col_values = sheet.col_values(index)
    col_len = len(col_values)
    for _ in takewhile(lambda x: not x, reversed(col_values)):
        col_len -= 1
    return col_len

def reporthook(count, block_size, total_size):
    global start_time
    if count == 0:
        start_time = time.time()
        return
    duration = time.time() - start_time
    progress_size = int(count * block_size)
    speed = int(progress_size / (1024 * duration))
    percent = int(count * block_size * 100 / total_size)
    sys.stdout.write("\r...%d%%, %d MB, %d KB/s, %d seconds passed" %
                    (percent, progress_size / (1024 * 1024), speed, duration))
    sys.stdout.flush()

workbook = xlrd.open_workbook('springer.xlsx')
worksheet = workbook.sheet_by_name('eBook list')
length = col_len(worksheet,17)

for r in range(1,length):
    doi = worksheet.cell(r,17).value
    isbn = doi[23:]
    num = doi[15:22]
    url = 'https://link.springer.com/content/pdf/' + num + '%2F' + isbn + '.pdf'
    name = worksheet.cell(r,0).value
    author =  worksheet.cell(r,1).value
    filename = isbn+'_'+name+'_'+author
    filename = filename.replace(' ','').replace('.','').replace(',','').replace('/','')+'.pdf'
    try:
        f = open(filename)
        print(filename + " already downloaded")
    except IOError:
        print('Downloading: ' + name + ' by ' + author + ' from ' + url)
        try:
            urllib.request.urlretrieve(url, filename,reporthook)
            print()
            print(name + ' by ' + author + ' downloaded')
        except Exception as e:
            print('Error downloading ' + name + ' by ' + author + ' from ' + url)
            print(e)
print('Finished!')
