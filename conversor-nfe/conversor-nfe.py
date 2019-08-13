#!/usr/bin/env python
# -*- coding=utf-8 -*-
import os
import csv
import sys, traceback
import xml.etree.ElementTree as ET

DIGITS=2

def get_optional(v): 
    '''Get attribute text if the object is not None
    '''
    if v is not None:
        return v.text
    else: 
        return None

def round_optional(v, digits=2): 
    '''If value is not None then round it to number of digits and convert it to string 
    '''
    if v is not None:
        return str(round(float(v.text), digits)) 
    else:
        return None

def process_nfes(local):
    '''Fetch xml NFe files recursively in current directory, parse the files
    extracting the relevant information and write the output to a csv file
    '''
    parsed = [_parse_xml(f) for f in _fetch_xml_files(local)]

    with open('output.csv', mode='w') as csv_file:
        fieldnames = [
            'xNome', 'nNF', 'cProd', 'xProd', 'emitUF', 'destUF',
            'NCM', 'cEAN', 'cEANTrib', 
            'CEST', 'cProdANVISA', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd', 'vDesc', 'vTotTrib', 
            'ICMS_orig', 'ICMS_CST', 'ICMS_vBCSTRet', 'ICMS_pST', 'ICMS_vICMSSTRet', 
            'ICMS_modBC', 'ICMS_pRedBC', 'ICMS_vBC', 'ICMS_pICMS', 'ICMS_vICMS', 
            'ICMS_vBCFCPP', 'ICMS_pFCP', 'ICMS_vFCP', 
            'ICMS_modBCST', 'ICMS_pMVAST', 'ICMS_vBCST', 'ICMS_pICMSST', 'ICMS_vICMSST', 
            'ICMS_vBCFCPST', 'ICMS_pFCPST', 'ICMS_vFCPST',
            'PIS_CST', 'PIS_vBC', 'PIS_pPIS', 'PIS_vPIS', 
            'PIS_qBCProd', 'PIS_vAliqProd', 
            'COFINS_CST', 'COFINS_vBC', 'COFINS_pCOFINS', 'COFINS_vCOFINS', 
            'COFINS_qBCProd', 'COFINS_vAliqProd', 
            'ICMSUFDest_vBCUFDest', 'ICMSUFDest_pFCPUFDest', 'ICMSUFDest_pICMSUFDest', 
            'ICMSUFDest_pICMSInter', 'ICMSUFDest_pICMSInterPart', 'ICMSUFDest_vFCPUFDest', 
            'ICMSUFDest_vICMSUFDest', 'ICMSUFDest_vICMSUFRemet', 
            'IPI_cEnq', 'IPI_CST', 'IPI_vBC', 'IPI_pIPI', 'IPI_vIPI',
            'IPI_qUnid', 'IPI_vUnid', 'IPI_vIPI', 
            'TOTAL_ICMS_vBC', 'TOTAL_vICMS', 'TOTAL_vICMSDeson', 'TOTAL_vFCPUFDest', 
            'TOTAL_vICMSUFDest', 'TOTAL_vICMSUFRemet', 'TOTAL_vFCP', 'TOTAL_vBCST', 
            'TOTAL_vST', 'TOTAL_vFCPST', 'TOTAL_vFCPSTRet', 'TOTAL_vProd', 
            'TOTAL_vFrete', 'TOTAL_vSeg', 'TOTAL_vDesc', 'TOTAL_vII', 'TOTAL_vIPI', 
            'TOTAL_vIPIDevol', 'TOTAL_vPIS', 'TOTAL_vCOFINS', 'TOTAL_vOutro', 
            'TOTAL_vNF', 'TOTAL_vTotTrib', 'caminho']
        
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter=';', 
                                lineterminator = '\n')
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
    '''Recursive search for .xml files
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
    '''Receives a path for a xml NFe file and parses the relevant data
    '''
    try:
        root = ET.parse(path)

        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'} 
        inf = root.find('ns:NFe/ns:infNFe', ns)        
        if inf is None:
            print(f"ATENÇÃO! O documento '{path}' não é válido e será ignorado!")
            return None

        nNF = inf.find('ns:ide/ns:nNF', ns).text
        
        emit = inf.find('ns:emit', ns)
        xNome = emit.find('ns:xNome', ns).text
        emitUF = emit.find('ns:enderEmit/ns:UF', ns).text
        destUF = inf.find('ns:dest/ns:enderDest/ns:UF', ns).text

        ICMSTot = inf.find('ns:total/ns:ICMSTot', ns)
        totals = {}
        totals['TOTAL_ICMS_vBC'] = ICMSTot.find('ns:vBC', ns).text
        totals['TOTAL_vICMS'] = ICMSTot.find('ns:vICMS', ns).text
        totals['TOTAL_vICMSDeson'] = round_optional(ICMSTot.find('ns:vICMSDeson', ns))
        totals['TOTAL_vFCPUFDest'] = round_optional(ICMSTot.find('ns:vFCPUFDest', ns))
        totals['TOTAL_vICMSUFDest'] = round_optional(ICMSTot.find('ns:vICMSUFDest', ns))
        totals['TOTAL_vICMSUFRemet'] = round_optional(ICMSTot.find('ns:vICMSUFRemet', ns))
        totals['TOTAL_vFCP'] = round_optional(ICMSTot.find('ns:vFCP', ns))
        totals['TOTAL_vBCST'] = round_optional(ICMSTot.find('ns:vBCST', ns))
        totals['TOTAL_vST'] = round_optional(ICMSTot.find('ns:vST', ns))
        totals['TOTAL_vFCPST'] = round_optional(ICMSTot.find('ns:vFCPST', ns))
        totals['TOTAL_vFCPSTRet'] = round_optional(ICMSTot.find('ns:vFCPSTRet', ns))
        totals['TOTAL_vProd'] = round_optional(ICMSTot.find('ns:vProd', ns))
        totals['TOTAL_vFrete'] = round_optional(ICMSTot.find('ns:vFrete', ns))
        totals['TOTAL_vSeg'] = round_optional(ICMSTot.find('ns:vSeg', ns))
        totals['TOTAL_vDesc'] = round_optional(ICMSTot.find('ns:vDesc', ns))
        totals['TOTAL_vII'] = round_optional(ICMSTot.find('ns:vII', ns))
        totals['TOTAL_vIPI'] = round_optional(ICMSTot.find('ns:vIPI', ns))
        totals['TOTAL_vIPIDevol'] = round_optional(ICMSTot.find('ns:vIPIDevol', ns))
        totals['TOTAL_vPIS'] = round_optional(ICMSTot.find('ns:vPIS', ns))
        totals['TOTAL_vCOFINS'] = round_optional(ICMSTot.find('ns:vCOFINS', ns))
        totals['TOTAL_vOutro'] = round_optional(ICMSTot.find('ns:vOutro', ns))
        totals['TOTAL_vNF'] = round_optional(ICMSTot.find('ns:vNF', ns))
        totals['TOTAL_vTotTrib'] = round_optional(ICMSTot.find('ns:vTotTrib', ns))

        nfes = []
        
        for det in inf.findall('ns:det', ns):
            # Join totals dict with the new one being created
            nfe = {
                'nNF': nNF,
                'xNome': xNome,
                'caminho': path,
                'emitUF': emitUF,
                'destUF': destUF,
                **totals
            }
            
            # Processing attributes of prod
            prod = det.find('ns:prod', ns)

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
            nfe['vDesc'] = round_optional(prod.find('ns:vDesc', ns))
            nfe['cEANTrib'] = get_optional(prod.find('ns:cEANTrib', ns))
            nfe['cProdANVISA'] = get_optional(prod.find('ns:med/ns:cProdANVISA', ns))

            nfe['qCom'] = str(round(float(nfe['qCom']), DIGITS))
            nfe['vUnCom'] = str(round(float(nfe['vUnCom']), DIGITS))
            nfe['vProd'] = str(round(float(nfe['vProd']), DIGITS))

            # Processing attributes of imposto
            imposto = det.find('ns:imposto', ns)
            nfe['vTotTrib'] = get_optional(imposto.find('ns:vTotTrib', ns))
            
            # Imposto ICMS
            inner_icms = imposto.find('ns:ICMS', ns).findall('*')
            if len(inner_icms) > 1: raise Exception('Múltiplos campos dentro de ICMS')
            inner_icms = inner_icms[0]

            nfe['ICMS_orig'] = inner_icms.find('ns:orig', ns).text
            nfe['ICMS_CST'] = get_optional(inner_icms.find('ns:CST', ns))
            nfe['ICMS_vBCSTRet'] = get_optional(inner_icms.find('ns:vBCSTRet', ns))
            nfe['ICMS_pST'] = get_optional(inner_icms.find('ns:pST', ns))
            nfe['ICMS_vICMSSTRet'] = get_optional(inner_icms.find('ns:vICMSSTRet', ns))
            
            nfe['ICMS_modBC'] = get_optional(inner_icms.find('ns:modBC', ns))
            nfe['ICMS_pRedBC'] = round_optional(inner_icms.find('ns:pRedBC', ns))
            nfe['ICMS_vBC'] = round_optional(inner_icms.find('ns:vBC', ns))
            nfe['ICMS_pICMS'] = round_optional(inner_icms.find('ns:pICMS', ns))
            nfe['ICMS_vICMS'] = round_optional(inner_icms.find('ns:vICMS', ns))
            nfe['ICMS_vBCFCPP'] = round_optional(inner_icms.find('ns:vBCFCPP', ns))
            nfe['ICMS_pFCP'] = round_optional(inner_icms.find('ns:pFCP', ns))
            nfe['ICMS_vFCP'] = round_optional(inner_icms.find('ns:vFCP', ns))
            
            # ICMS filds that may appear in ICMS70 and ICMS10
            nfe['ICMS_modBCST'] = get_optional(inner_icms.find('ns:modBCST', ns))
            nfe['ICMS_pMVAST'] = round_optional(inner_icms.find('ns:pMVAST', ns))
            nfe['ICMS_vBCST'] = round_optional(inner_icms.find('ns:vBCST', ns))
            nfe['ICMS_pICMSST'] = round_optional(inner_icms.find('ns:pICMSST', ns))
            nfe['ICMS_vICMSST'] = round_optional(inner_icms.find('ns:vICMSST', ns))
            nfe['ICMS_vBCFCPST'] = round_optional(inner_icms.find('ns:vBCFCPST', ns))
            nfe['ICMS_pFCPST'] = round_optional(inner_icms.find('ns:pFCPST', ns))
            nfe['ICMS_vFCPST'] = round_optional(inner_icms.find('ns:vFCPST', ns))

            # Imposto PIS
            pis = imposto.find('ns:PIS', ns).findall('*')
            if len(pis) > 1: raise Exception('Múltiplos campos dentro de PIS')
            pis = pis[0]
            
            tags = ['CST', 'vBC', 'pPIS', 'vPIS', 'qBCProd', 'vAliqProd']
            for c in pis.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em PIS: ' + tag)
            
            nfe['PIS_CST'] = pis.find('ns:CST', ns).text
            nfe['PIS_vBC'] = round_optional(pis.find('ns:vBC', ns))
            nfe['PIS_pPIS'] = round_optional(pis.find('ns:pPIS', ns))
            nfe['PIS_vPIS'] = round_optional(pis.find('ns:vPIS', ns))
            nfe['PIS_qBCProd'] = round_optional(pis.find('ns:qBCProd', ns))
            nfe['PIS_vAliqProd'] = round_optional(pis.find('ns:vAliqProd', ns))

            # Imposto COFINS
            cofins = imposto.find('ns:COFINS', ns).findall('*')
            if len(cofins) > 1: raise Exception('Múltiplos campos dentro de COFINS')
            cofins = cofins[0]

            tags = ['CST', 'vBC', 'pCOFINS', 'vCOFINS', 'qBCProd', 'vAliqProd']
            for c in cofins.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em COFINS: ' + tag)

            nfe['COFINS_CST'] = cofins.find('ns:CST', ns).text
            nfe['COFINS_vBC'] = round_optional(cofins.find('ns:vBC', ns))
            nfe['COFINS_pCOFINS'] = round_optional(cofins.find('ns:pCOFINS', ns))
            nfe['COFINS_vCOFINS'] = round_optional(cofins.find('ns:vCOFINS', ns))
            nfe['COFINS_qBCProd'] = round_optional(cofins.find('ns:qBCProd', ns))
            nfe['COFINS_vAliqProd'] = round_optional(cofins.find('ns:vAliqProd', ns))

            # Imposto ICMSUFDest
            icmsufdest = imposto.find('ns:ICMSUFDest', ns)
            if icmsufdest:
                nfe['ICMSUFDest_vBCUFDest'] = round_optional(icmsufdest.find('ns:vBCUFDest', ns))
                nfe['ICMSUFDest_pFCPUFDest'] = round_optional(icmsufdest.find('ns:pFCPUFDest', ns))
                nfe['ICMSUFDest_pICMSUFDest'] = round_optional(icmsufdest.find('ns:pICMSUFDest', ns))
                nfe['ICMSUFDest_pICMSInter'] = round_optional(icmsufdest.find('ns:pICMSInter', ns))
                nfe['ICMSUFDest_pICMSInterPart'] = round_optional(icmsufdest.find('ns:pICMSInterPart', ns))
                nfe['ICMSUFDest_vFCPUFDest'] = round_optional(icmsufdest.find('ns:vFCPUFDest', ns))
                nfe['ICMSUFDest_vICMSUFDest'] = round_optional(icmsufdest.find('ns:vICMSUFDest', ns))
                nfe['ICMSUFDest_vICMSUFRemet'] = round_optional(icmsufdest.find('ns:vICMSUFRemet', ns))
            
            # Imposto IPI
            ipi = imposto.find('ns:IPI', ns)
            if ipi:
                nfe['IPI_cEnq'] = get_optional(ipi.find('ns:cEnq', ns))
                
                ipint = ipi.find('ns:IPINT', ns)
                ipitrib = ipi.find('ns:IPITrib', ns)

                #If IPI is "tributed" or not
                if ipint:
                    nfe['IPI_CST'] = ipint.find('ns:CST', ns).text
                elif ipitrib:
                    nfe['IPI_CST'] = get_optional(ipitrib.find('ns:CST', ns))
                    
                    #Some documents have vBC and pIPI others have qUnid and vUnid
                    nfe['IPI_vBC'] = round_optional(ipitrib.find('ns:vBC', ns))
                    nfe['IPI_pIPI'] = get_optional(ipitrib.find('ns:pIPI', ns))
                    
                    nfe['IPI_qUnid'] = get_optional(ipitrib.find('ns:qUnid', ns))
                    nfe['IPI_vUnid'] = round_optional(ipitrib.find('ns:vUnid', ns))
                    
                    nfe['IPI_vIPI'] = get_optional(ipitrib.find('ns:vIPI', ns))
                
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
