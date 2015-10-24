#!/usr/bin/env python3

import urllib.parse
import requests
from lxml import etree

def download_file(url, chunk_size=1024):
    local_filename = url.split('/')[-1]
    resp = requests.get(url, stream=True)
    with open(local_filename, 'wb') as f:
        for chunk in resp.iter_content(chunk_size): 
            if chunk: # filter out keep-alive new chunks
                f.write(chunk)
    return local_filename

if __name__ == '__main__':
    url = 'http://unctad.org/en/Pages/DIAE/FDI%20Statistics/FDI-Statistics-Bilateral.aspx'
    resp = requests.get(url)

    html_tree = etree.HTML(resp.content)

    xpath_expr = '//select[@id="FDIcountriesxls"]/option[@value and string-length(@value)!=0]/@value'
    xls_paths = html_tree.xpath(xpath_expr)
    print('Found', len(xls_paths), 'workbooks')

    for rel_path in xls_paths:
        print('Downloading', rel_path)
        full_url = urllib.parse.urljoin(url, rel_path)
        download_file(full_url)
