#!/usr/bin/env python
# -*- coding=utf-8 -*-
import os
import csv
import sys, traceback
import xml.etree.ElementTree as ET

DIGITS = 2

def process_nfes(local):
    '''
    Fetch xml NFe files recursively in current directory, parse the files
    extracting the relevant information and write the output to a csv file
    '''
    parsed = [_parse_xml(f) for f in _fetch_xml_files(local)]

    with open('output.csv', mode='w') as csv_file:
        fieldnames = [
            'xNome', 'nNF', 'cProd', 'cEAN', 'xProd', 'NCM', 'cEANTrib', 
            'CEST', 'cProdANVISA', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd', 'vTotTrib', 
            'icms_orig', 'icms_CST', 'icms_vBCSTRet', 'icms_pST', 'icms_vICMSSTRet', 
            'icms_modBC', 'icms_pRedBC', 'icms_vBC', 'icms_pICMS', 'icms_vICMS', 
            'icms_vBCFCPP', 'icms_pFCP', 'icms_vFCP', 
            'pis_CST', 'pis_vBC', 'pis_pPIS', 'pis_vPIS', 
            'pis_qBCProd', 'pis_vAliqProd', 
            'cofins_CST', 'cofins_vBC', 'cofins_pCOFINS', 'cofins_vCOFINS', 
            'cofins_qBCProd', 'cofins_vAliqProd', 
            'ICMSUFDest_vBCUFDest', 'ICMSUFDest_pFCPUFDest', 'ICMSUFDest_pICMSUFDest', 
            'ICMSUFDest_pICMSInter', 'ICMSUFDest_pICMSInterPart', 'ICMSUFDest_vFCPUFDest', 
            'ICMSUFDest_vICMSUFDest', 'ICMSUFDest_vICMSUFRemet', 
            'total_vBC', 'total_vICMS', 'total_vICMSDeson', 'total_vFCPUFDest', 
            'total_vICMSUFDest', 'total_vICMSUFRemet', 'total_vFCP', 'total_vBCST', 
            'total_vST', 'total_vFCPST', 'total_vFCPSTRet', 'total_vProd', 
            'total_vFrete', 'total_vSeg', 'total_vDesc', 'total_vII', 'total_vIPI', 
            'total_vIPIDevol', 'total_vPIS', 'total_vCOFINS', 'total_vOutro', 
            'total_vNF', 'total_vTotTrib', 'caminho']
        
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        
        count = 0
        for nfes in parsed:
            if nfes is None:
                continue
            
            count += 1
            for nfe in nfes:
                if nfe is not None:
                    writer.writerow(nfe)
    
    print('Aplicação executou com sucesso!')
    print(f'FORAM PROCESSADOS {count} DOCUMENTOS')

def _fetch_xml_files(path):
    '''
    Recursive search for .xml files
    ''' 
    paths = []

    for (dirpath, dirnames, filenames) in os.walk(path):
        for f in filenames:
            if f.lower().endswith('.xml'):
                fpath = os.path.join(dirpath, f)
                print(fpath)
                paths.append(fpath)
            else:
                print(f"O arquivo {f} não é um XML")

    if paths == [] and os.path.isfile(path) and path.lower().endswith('.xml'):
        paths.append(path)
    
    return paths

def _parse_xml(path):
    get_optional = lambda v: v.text if v is not None else None
    round_optional = lambda v: str(round(float(v.text), DIGITS)) if v is not None else None

    '''
    Receives a path for a xml NFe file and parses the relevant data
    '''
    try:
        root = ET.parse(path)

        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'} 
        inf = root.find('ns:NFe/ns:infNFe', ns)        
        if inf is None:
            print(f"ATENÇÃO! O documento '{path}' não é válido e será ignorado!")
            return None

        emit = inf.find('ns:emit', ns)
        nNF = inf.find('ns:ide/ns:nNF', ns).text

        ICMSTot = inf.find('ns:total/ns:ICMSTot', ns)
        vBC = ICMSTot.find('ns:vBC', ns).text
        vICMS = ICMSTot.find('ns:vICMS', ns).text
        vICMSDeson = ICMSTot.find('ns:vICMSDeson', ns).text
        vFCPUFDest = get_optional(ICMSTot.find('ns:vFCPUFDest', ns))
        vICMSUFDest = get_optional(ICMSTot.find('ns:vICMSUFDest', ns))
        vICMSUFRemet = get_optional(ICMSTot.find('ns:vICMSUFRemet', ns))
        vFCP = ICMSTot.find('ns:vFCP', ns).text
        vBCST = ICMSTot.find('ns:vBCST', ns).text
        vST = ICMSTot.find('ns:vST', ns).text
        vFCPST = ICMSTot.find('ns:vFCPST', ns).text
        vFCPSTRet = ICMSTot.find('ns:vFCPSTRet', ns).text
        vProd = ICMSTot.find('ns:vProd', ns).text
        vFrete = ICMSTot.find('ns:vFrete', ns).text
        vSeg = ICMSTot.find('ns:vSeg', ns).text
        vDesc = ICMSTot.find('ns:vDesc', ns).text
        vII = ICMSTot.find('ns:vII', ns).text
        vIPI = ICMSTot.find('ns:vIPI', ns).text
        vIPIDevol = ICMSTot.find('ns:vIPIDevol', ns).text
        vPIS = ICMSTot.find('ns:vPIS', ns).text
        vCOFINS = ICMSTot.find('ns:vCOFINS', ns).text
        vOutro = ICMSTot.find('ns:vOutro', ns).text
        vNF = ICMSTot.find('ns:vNF', ns).text
        vTotTrib = get_optional(ICMSTot.find('ns:vTotTrib', ns))

        nfes = []
        
        for det in inf.findall('ns:det', ns):
            nfe = {
                'caminho': path
            }
            
            nfe['xNome'] = emit.find('ns:xNome', ns).text
            
            # Processing attributes of prod
            nfe['prod'] = det.find('ns:prod', ns)

            nfe['cProd'] = prod.find('ns:cProd', ns).text
            nfe['cEAN'] = get_optional(prod.find('ns:cEAN', ns))
            nfe['xProd'] = prod.find('ns:xProd', ns).text
            nfe['NCM'] = prod.find('ns:NCM', ns).text
            nfe['CEST'] = get_optional(prod.find('ns:CEST', ns))
            nfe['CFOP'] = prod.find('ns:CFOP', ns).text
            nfe['uCom'] = prod.find('ns:uCom', ns).text
            nfe['qCom'] = prod.find('ns:qCom', ns).text
            nfe['vUnCom'] = prod.find('ns:vUnCom', ns).text
            nfe['vProd'] = prod.find('ns:vProd', ns).text
            nfe['cEANTrib'] = get_optional(prod.find('ns:cEANTrib', ns))
            nfe['cProdANVISA'] = get_optional(prod.find('ns:med/ns:cProdANVISA', ns))

            nfe['qCom'] = str(round(float(qCom), DIGITS))
            nfe['vUnCom'] = str(round(float(vUnCom), DIGITS))
            nfe['vProd'] = str(round(float(vProd), DIGITS))

            # Processing attributes of imposto
            nfe['imposto'] = det.find('ns:imposto', ns)
            nfe['vTotTrib'] = get_optional(imposto.find('ns:vTotTrib', ns))
            
            # Imposto ICMS
            inner_icms = imposto.find('ns:ICMS', ns).findall('*')
            if len(inner_icms) > 1: raise Exception('Múltiplos campos dentro de ICMS')
            nfe['inner_icms'] = inner_icms[0]

            nfe['icms_orig'] = inner_icms.find('ns:orig', ns).text
            nfe['icms_CST'] = get_optional(inner_icms.find('ns:CST', ns))
            nfe['icms_vBCSTRet'] = get_optional(inner_icms.find('ns:vBCSTRet', ns))
            nfe['icms_pST'] = get_optional(inner_icms.find('ns:pST', ns))
            nfe['icms_vICMSSTRet'] = get_optional(inner_icms.find('ns:vICMSSTRet', ns))
            
            nfe['icms_modBC'] = get_optional(inner_icms.find('ns:modBC', ns))
            nfe['icms_pRedBC'] = round_optional(inner_icms.find('ns:pRedBC', ns))
            nfe['icms_vBC'] = round_optional(inner_icms.find('ns:vBC', ns))
            nfe['icms_pICMS'] = round_optional(inner_icms.find('ns:pICMS', ns))
            nfe['icms_vICMS'] = round_optional(inner_icms.find('ns:vICMS', ns))
            nfe['icms_vBCFCPP'] = round_optional(inner_icms.find('ns:vBCFCPP', ns))
            nfe['icms_pFCP'] = round_optional(inner_icms.find('ns:pFCP', ns))
            nfe['icms_vFCP'] = round_optional(inner_icms.find('ns:vFCP', ns))
            
            # Imposto PIS
            pis = imposto.find('ns:PIS', ns).findall('*')
            if len(pis) > 1: raise Exception('Múltiplos campos dentro de PIS')
            nfe['pis'] = pis[0]
            
            tags = ['CST', 'vBC', 'pPIS', 'vPIS', 'qBCProd', 'vAliqProd']
            for c in pis.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em PIS: ' + tag)
            
            nfe['pis_CST'] = pis.find('ns:CST', ns).text
            nfe['pis_vBC'] = round_optional(pis.find('ns:vBC', ns))
            nfe['pis_pPIS'] = round_optional(pis.find('ns:pPIS', ns))
            nfe['pis_vPIS'] = round_optional(pis.find('ns:vPIS', ns))
            nfe['pis_qBCProd'] = round_optional(pis.find('ns:qBCProd', ns))
            nfe['pis_vAliqProd'] = round_optional(pis.find('ns:vAliqProd', ns))

            # Imposto COFINS
            cofins = imposto.find('ns:COFINS', ns).findall('*')
            if len(cofins) > 1: raise Exception('Múltiplos campos dentro de COFINS')
            nfe['cofins'] = cofins[0]

            tags = ['CST', 'vBC', 'pCOFINS', 'vCOFINS', 'qBCProd', 'vAliqProd']
            for c in cofins.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em COFINS: ' + tag)

            nfe['cofins_CST'] = cofins.find('ns:CST', ns).text
            nfe['cofins_vBC'] = round_optional(cofins.find('ns:vBC', ns))
            nfe['cofins_pCOFINS'] = round_optional(cofins.find('ns:pCOFINS', ns))
            nfe['cofins_vCOFINS'] = round_optional(cofins.find('ns:vCOFINS', ns))
            nfe['cofins_qBCProd'] = round_optional(cofins.find('ns:qBCProd', ns))
            nfe['cofins_vAliqProd'] = round_optional(cofins.find('ns:vAliqProd', ns))

            # Imposto ICMSUFDest
            icmsufdest = imposto.find('ns:ICMSUFDest', ns)
            nfe['ICMSUFDest_vBCUFDest'] = round_optional(icmsufdest.find('ns:vBCUFDest', ns))
            nfe['ICMSUFDest_pFCPUFDest'] = round_optional(icmsufdest.find('ns:pFCPUFDest', ns))
            nfe['ICMSUFDest_pICMSUFDest'] = round_optional(icmsufdest.find('ns:pICMSUFDest', ns))
            nfe['ICMSUFDest_pICMSInter'] = round_optional(icmsufdest.find('ns:pICMSInter', ns))
            nfe['ICMSUFDest_pICMSInterPart'] = round_optional(icmsufdest.find('ns:pICMSInterPart', ns))
            nfe['ICMSUFDest_vFCPUFDest'] = round_optional(icmsufdest.find('ns:vFCPUFDest', ns))
            nfe['ICMSUFDest_vICMSUFDest'] = round_optional(icmsufdest.find('ns:vICMSUFDest', ns))
            nfe['ICMSUFDest_vICMSUFRemet'] = round_optional(icmsufdest.find('ns:vICMSUFRemet', ns))
            
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
