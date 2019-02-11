#!/usr/bin/env python
# -*- coding=utf-8 -*-
import os
import csv
import sys, traceback
import xml.etree.ElementTree as ET

DECIMAL_DIGITS = 2

def process_nfes(local):
    '''
    Fetch xml NFe files recursively in current directory, parse the files
    extracting the relevant information and write the output to a csv file
    '''
    parsed = [_parse_xml(f) for f in _fetch_xml_files(local)]

    with open('output.tsv', mode='w') as csv_file:
        fieldnames = ['xNome', 'nNF', 'cProd', 'cEAN', 'xProd', 'NCM', 'cEANTrib', 
                      'CEST', 'cProdANVISA', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd', 'vTotTrib', 
                      'icms_orig', 'icms_CST', 'icms_vBCSTRet', 'icms_pST', 'icms_vICMSSTRet', 
                      'pis_CST', 'pis_vBC', 'pis_pPIS', 'pis_vPIS', 
                      'cofins_CST', 'cofins_vBC', 'cofins_pCOFINS', 'cofins_vCOFINS', 
                      'total_vBC', 'total_vICMS', 'total_vICMSDeson', 'total_vFCPUFDest', 
                      'total_vICMSUFDest', 'total_vICMSUFRemet', 'total_vFCP', 'total_vBCST', 
                      'total_vST', 'total_vFCPST', 'total_vFCPSTRet', 'total_vProd', 
                      'total_vFrete', 'total_vSeg', 'total_vDesc', 'total_vII', 'total_vIPI', 
                      'total_vIPIDevol', 'total_vPIS', 'total_vCOFINS', 'total_vOutro', 
                      'total_vNF', 'total_vTotTrib', 'caminho']
        
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter='\t')
        writer.writeheader()
        
        for nfes in parsed:
            for nfe in nfes:
                if nfe is not None:
                    writer.writerow(nfe)
    
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
        if inf is None:
            return None

        emit = inf.find('ns:emit', ns)
        nNF = inf.find('ns:ide/ns:nNF', ns).text

        ICMSTot = inf.find('ns:total/ns:ICMSTot', ns)
        vBC = ICMSTot.find('ns:vBC', ns).text
        vICMS = ICMSTot.find('ns:vICMS', ns).text
        vICMSDeson = ICMSTot.find('ns:vICMSDeson', ns).text
        vFCPUFDest = ICMSTot.find('ns:vFCPUFDest', ns).text
        vICMSUFDest = ICMSTot.find('ns:vICMSUFDest', ns).text
        vICMSUFRemet = ICMSTot.find('ns:vICMSUFRemet', ns).text
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
        vTotTrib = ICMSTot.find('ns:vTotTrib', ns).text

        get_optional = lambda v: v.text if v is not None else None

        nfes = []
            
        for det in inf.findall('ns:det', ns):
            xNome = emit.find('ns:xNome', ns).text
            
            # Processing attributes of prod
            prod = det.find('ns:prod', ns)

            cProd = prod.find('ns:cProd', ns).text
            cEAN = get_optional(prod.find('ns:cEAN', ns))
            xProd = prod.find('ns:xProd', ns).text
            NCM = prod.find('ns:NCM', ns).text
            CEST = get_optional(prod.find('ns:CEST', ns))
            CFOP = prod.find('ns:CFOP', ns).text
            uCom = prod.find('ns:uCom', ns).text
            qCom = prod.find('ns:qCom', ns).text
            vUnCom = prod.find('ns:vUnCom', ns).text
            vProd = prod.find('ns:vProd', ns).text
            cEANTrib = get_optional(prod.find('ns:cEANTrib', ns))
            cProdANVISA = get_optional(prod.find('ns:med/ns:cProdANVISA', ns))

            qCom = str(round(float(qCom), DECIMAL_DIGITS))
            vUnCom = str(round(float(vUnCom), DECIMAL_DIGITS))
            vProd = str(round(float(vProd), DECIMAL_DIGITS))

            # Processing attributes of imposto
            imposto = det.find('ns:imposto', ns)
            vTotTrib = get_optional(imposto.find('ns:vTotTrib', ns))
            
            # Imposto ICMS
            inner_icms = imposto.find('ns:ICMS', ns).findall('*')
            if len(inner_icms) > 1: raise Exception('Múltiplos campos dentro de ICMS')
            inner_icms = inner_icms[0]

            icms_orig = inner_icms.find('ns:orig', ns).text
            icms_CST = get_optional(inner_icms.find('ns:CST', ns))
            icms_vBCSTRet = get_optional(inner_icms.find('ns:vBCSTRet', ns))
            icms_pST = get_optional(inner_icms.find('ns:pST', ns))
            icms_vICMSSTRet = get_optional(inner_icms.find('ns:vICMSSTRet', ns))

            # Imposto PIS
            pis = imposto.find('ns:PIS/ns:PISAliq', ns)

            pis_CST = pis.find('ns:CST', ns).text
            pis_vBC = str(round(float(pis.find('ns:vBC', ns).text), DECIMAL_DIGITS))
            pis_pPIS = str(round(float(pis.find('ns:pPIS', ns).text), DECIMAL_DIGITS))
            pis_vPIS = str(round(float(pis.find('ns:vPIS', ns).text), DECIMAL_DIGITS))

            # Imposto COFINS
            cofins = imposto.find('ns:COFINS/ns:COFINSAliq', ns)

            cofins_CST = cofins.find('ns:CST', ns).text
            cofins_vBC = str(round(float(cofins.find('ns:vBC', ns).text), DECIMAL_DIGITS))
            cofins_pCOFINS = str(round(float(cofins.find('ns:pCOFINS', ns).text), DECIMAL_DIGITS))
            cofins_vCOFINS = str(round(float(cofins.find('ns:vCOFINS', ns).text), DECIMAL_DIGITS))

            nfe = {
                'caminho': path,
                'xNome': xNome,
                'nNF': nNF,
                # Produto
                'cProd': cProd,
                'cEAN': cEAN,
                'xProd': xProd,
                'NCM': NCM,
                'cEANTrib': cEANTrib,
                'CEST': CEST,
                'cProdANVISA': cProdANVISA,
                'CFOP': CFOP,
                'uCom': uCom,
                'qCom': qCom,
                'vUnCom': vUnCom,
                'vProd': vProd,
                # Imposto
                'vTotTrib': vTotTrib,
                # Imposto ICMS
                'icms_orig': icms_orig,
                'icms_CST': icms_CST,
                'icms_vBCSTRet': icms_vBCSTRet,
                'icms_pST': icms_pST,
                'icms_vICMSSTRet': icms_vICMSSTRet,
                # Imposto PIS
                'pis_CST': pis_CST,
                'pis_vBC': pis_vBC,
                'pis_pPIS': pis_pPIS,
                'pis_vPIS': pis_vPIS,
                # Imposto COFINS
                'cofins_CST': cofins_CST,
                'cofins_vBC': cofins_vBC,
                'cofins_pCOFINS': cofins_pCOFINS,
                'cofins_vCOFINS': cofins_vCOFINS,
                # Totals
                'total_vBC': vBC,
                'total_vICMS': vICMS,
                'total_vICMSDeson': vICMSDeson,
                'total_vFCPUFDest': vFCPUFDest,
                'total_vICMSUFDest': vICMSUFDest,
                'total_vICMSUFRemet': vICMSUFRemet,
                'total_vFCP': vFCP,
                'total_vBCST': vBCST,
                'total_vST': vST,
                'total_vFCPST': vFCPST,
                'total_vFCPSTRet': vFCPSTRet,
                'total_vProd': vProd,
                'total_vFrete': vFrete,
                'total_vSeg': vSeg,
                'total_vDesc': vDesc,
                'total_vII': vII,
                'total_vIPI': vIPI,
                'total_vIPIDevol': vIPIDevol,
                'total_vPIS': vPIS,
                'total_vCOFINS': vCOFINS,
                'total_vOutro': vOutro,
                'total_vNF': vNF,
                'total_vTotTrib': vTotTrib
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
