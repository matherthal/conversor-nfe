#!/usr/bin/env python
import xml.etree.ElementTree as ET
import os
import csv
import sys, traceback
import sys, traceback

def process_nfes(local):
    '''
    Fetch xml NFe files recursively in current directory, parse the files
    extracting the relevant information and write the output to a csv file
    '''
    parsed = [_parse_xml(f) for f in _fetch_xml_files(local)]

    with open('output.tsv', mode='w') as csv_file:
        fieldnames = ['caminho', 'xNome', 'cProd', 'cEAN', 'xProd',
                      'NCM', 'cEANTrib', 'CEST', 'orig', 'cProdANVISA']
        
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter='\t')
        writer.writeheader()
        
        for nfes in parsed:
            for nfe in nfes:
                if nfe is not None:
                    writer.writerow(nfe)
                    print(nfe['caminho'])
    
    print('Aplicação executou com sucesso!')

def _fetch_xml_files(path):
    '''
    Recursive search for .xml files
    '''
    paths = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        for f in filenames:
            if f.endswith('.xml'):
                fpath = os.path.join(dirpath, f)
                print(fpath)
                paths.append(fpath)
    return paths

def _parse_xml(path):
    '''
    Receives a path for a xml NFe file and parses the relevant data
    '''
    try:
        root = ET.parse(path)

        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'} 
        inf = root.find('ns:NFe/ns:infNFe', ns)              
        # ide = inf.find('ns:ide', ns)  
        emit = inf.find('ns:emit', ns)
        # dest = inf.find('ns:dest', ns)

        if inf is None:
            return None

        nfes = []
            
        for det in inf.findall('ns:det', ns):
            xNome = emit.find('ns:xNome', ns).text
            
            #Atribubtos de prod
            for prod in det.findall('ns:prod', ns):
                cProd = prod.find('ns:cProd', ns).text
                cEAN = prod.find('ns:cEAN', ns)
                if cEAN is not None:
                    cEAN = cEAN.text
                xProd = prod.find('ns:xProd', ns).text
                NCM = prod.find('ns:NCM', ns).text
                cEANTrib = prod.find('ns:cEANTrib', ns).text
                CEST = prod.find('ns:CEST', ns)
                if CEST is not None:
                    CEST = CEST.text
                cProdANVISA = prod.find('ns:med/ns:cProdANVISA', ns)
                if cProdANVISA:
                    cProdANVISA = cProdANVISA.text

                for node in det.find('ns:imposto/ns:ICMS', ns).findall('*'):
                    orig = node.find('ns:orig', ns).text
                
                nfe = {
                    'caminho': path,
                    'xNome': xNome,
                    'cProd': cProd,
                    'cEAN': cEAN,
                    'xProd': xProd,
                    'NCM': NCM,
                    'cEANTrib': cEANTrib,
                    'CEST': CEST,
                    'orig': orig,
                    'cProdANVISA': cProdANVISA
                }
                
                nfes.append(nfe)
                
        return nfes
    except:
        print("ERRO! Não consegui ler o arquivo " + path)
        traceback.print_exc(file=sys.stdout)

if __name__ == '__main__':
    try:
        local = None
        if len(sys.argv) > 2:
            raise Exception('Unknown parameters')
        elif len(sys.argv) == 2:
            local = os.path.abspath(sys.argv[1])
        else:
            local = os.path.dirname(os.path.realpath(__file__))
        
        print('Lendo arquivos de: ' + local)
        
        process_nfes(local)
    except:
        traceback.print_exc(file=sys.stdout)
        
    sys.stdin.readline()
