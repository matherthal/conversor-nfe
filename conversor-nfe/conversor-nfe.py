#!/usr/bin/env python
# -*- coding=utf-8 -*-
import os
import csv
import sys, traceback
import xml.etree.ElementTree as ET
from datetime import datetime

def get_optional(v): 
    '''Get attribute text if the object is not None
    '''
    if v is not None:
        return v.text
    else: 
        return None

def process_nfes(local):
    '''Fetch xml NFe files recursively in current directory, parse the files
    extracting the relevant information and write the output to a csv file
    '''
    parsed = [_parse_xml(f) for f in _fetch_xml_files(local)]

    with open('output.csv', mode='w') as csv_file:
        fieldnames = [
            'emit_xNome', 'emit_UF', 'emit_CNPJ', 
            'dest_xNome', 'dest_UF', 'dest_CNPJ', 
            'nNF', 'refNFe', 'dhEmi_data', 'dhEmi_hora', 'dhSaiEnt_data', 
            'dhSaiEnt_hora', 'cProd', 'xProd', 
            'NCM', 'cEAN', 'cEANTrib', 'uTrib', 'qTrib', 'nRECOPI', 'CEST', 'cBenef', 'vFrete',
            'cProdANVISA', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd', 'vDesc', 
            'vTotTrib', 
            'ICMS_orig', 'ICMS_CST', 'ICMS_CSOSN', 'ICMS_vBCSTRet', 'ICMS_pST', 
            'ICMS_vICMSSTRet', 'ICMS_modBC', 'ICMS_pRedBC', 'ICMS_vBC', 
            'ICMS_pICMS', 'ICMS_vICMS', 'ICMS_vBCFCPP', 'ICMS_pFCP', 'ICMS_vFCP', 
            'ICMS_modBCST', 'ICMS_pMVAST', 'ICMS_vBCST', 'ICMS_pICMSST', 
            'ICMS_vICMSST', 'ICMS_vBCFCPST', 'ICMS_pFCPST', 'ICMS_vFCPST',
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
            'TOTAL_vNF', 'TOTAL_vTotTrib', 'infCpl', 'chNFe', 'nItem', 'caminho']

        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, delimiter=';', 
                                lineterminator = '\n')
        writer.writeheader()
        
        count = 0
        first_line = True
        for nfes in parsed:
            if nfes is None:
                continue

            count += 1
            for nfe in nfes:
                if nfe is not None and not first_line:
                    writer.writerow(nfe)
                first_line = False

    print('Aplicação executou com sucesso!')
    print(f'FORAM PROCESSADOS {count} DOCUMENTO(S)')

def _fetch_xml_files(path):
    '''Recursive search for .xml files
    ''' 
    paths = []

    current_file = os.path.basename(__file__)

    for (dirpath, dirnames, filenames) in os.walk(path):
        for f in filenames:
            if f.lower().endswith('.xml'):
                fpath = os.path.join(dirpath, f)
                print(fpath)
                paths.append(fpath)
            elif f != current_file:
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

        chNFe = get_optional(root.find('ns:protNFe/ns:infProt/ns:chNFe', ns))
        chNFe = f"{chNFe}_" if chNFe else ''

        inf = root.find('ns:NFe/ns:infNFe', ns)
        if inf is None:
            print(f"ATENÇÃO! O documento '{path}' não é válido e será ignorado!")
            return None

        ide = inf.find('ns:ide', ns)
        nNF = get_optional(ide.find('ns:nNF', ns))
        refNFe = get_optional(ide.find('ns:NFref/ns:refNFe', ns))

        dhEmi = get_optional(ide.find('ns:dhEmi', ns))
        if dhEmi:
            dhEmi = datetime.strptime(dhEmi, '%Y-%m-%dT%H:%M:%S%z')
            dhEmi_data = dhEmi.strftime('%Y-%m-%d')
            dhEmi_hora = dhEmi.strftime('%H:%M:%S')

        dhSaiEnt = get_optional(ide.find('ns:dhSaiEnt', ns))
        if dhSaiEnt:
            dhSaiEnt = datetime.strptime(dhSaiEnt, '%Y-%m-%dT%H:%M:%S%z')

        dhSaiEnt_data = dhSaiEnt.strftime('%Y-%m-%d') if dhSaiEnt else ''
        dhSaiEnt_hora = dhSaiEnt.strftime('%H:%M:%S') if dhSaiEnt else ''
        
        emit = inf.find('ns:emit', ns)
        emit_xNome = get_optional(emit.find('ns:xNome', ns))
        emit_UF = get_optional(emit.find('ns:enderEmit/ns:UF', ns))
        emit_CNPJ = get_optional(emit.find('ns:CNPJ', ns))

        dest = inf.find('ns:dest', ns)
        dest_xNome = get_optional(dest.find('ns:xNome', ns)) if dest else ''
        dest_UF = get_optional(dest.find('ns:enderDest/ns:UF', ns)) if dest else ''
        dest_CNPJ = get_optional(dest.find('ns:CNPJ', ns)) if dest else ''

        infCpl = get_optional(inf.find('ns:infAdic/ns:infCpl', ns))

        ICMSTot = inf.find('ns:total/ns:ICMSTot', ns)
        totals = {}
        totals['TOTAL_ICMS_vBC'] = get_optional(ICMSTot.find('ns:vBC', ns))
        totals['TOTAL_vICMS'] = get_optional(ICMSTot.find('ns:vICMS', ns))
        totals['TOTAL_vICMSDeson'] = get_optional(ICMSTot.find('ns:vICMSDeson', ns))
        totals['TOTAL_vFCPUFDest'] = get_optional(ICMSTot.find('ns:vFCPUFDest', ns))
        totals['TOTAL_vICMSUFDest'] = get_optional(ICMSTot.find('ns:vICMSUFDest', ns))
        totals['TOTAL_vICMSUFRemet'] = get_optional(ICMSTot.find('ns:vICMSUFRemet', ns))
        totals['TOTAL_vFCP'] = get_optional(ICMSTot.find('ns:vFCP', ns))
        totals['TOTAL_vBCST'] = get_optional(ICMSTot.find('ns:vBCST', ns))
        totals['TOTAL_vST'] = get_optional(ICMSTot.find('ns:vST', ns))
        totals['TOTAL_vFCPST'] = get_optional(ICMSTot.find('ns:vFCPST', ns))
        totals['TOTAL_vFCPSTRet'] = get_optional(ICMSTot.find('ns:vFCPSTRet', ns))
        totals['TOTAL_vProd'] = get_optional(ICMSTot.find('ns:vProd', ns))
        totals['TOTAL_vFrete'] = get_optional(ICMSTot.find('ns:vFrete', ns))
        totals['TOTAL_vSeg'] = get_optional(ICMSTot.find('ns:vSeg', ns))
        totals['TOTAL_vDesc'] = get_optional(ICMSTot.find('ns:vDesc', ns))
        totals['TOTAL_vII'] = get_optional(ICMSTot.find('ns:vII', ns))
        totals['TOTAL_vIPI'] = get_optional(ICMSTot.find('ns:vIPI', ns))
        totals['TOTAL_vIPIDevol'] = get_optional(ICMSTot.find('ns:vIPIDevol', ns))
        totals['TOTAL_vPIS'] = get_optional(ICMSTot.find('ns:vPIS', ns))
        totals['TOTAL_vCOFINS'] = get_optional(ICMSTot.find('ns:vCOFINS', ns))
        totals['TOTAL_vOutro'] = get_optional(ICMSTot.find('ns:vOutro', ns))
        totals['TOTAL_vNF'] = get_optional(ICMSTot.find('ns:vNF', ns))
        totals['TOTAL_vTotTrib'] = get_optional(ICMSTot.find('ns:vTotTrib', ns))

        nfes = []

        first_line = True
        for det in inf.findall('ns:det', ns):
            nItem = det.attrib['nItem'] if 'nItem' in det.attrib else ''

            # Join totals dict with the new one being created
            nfe = {
                'nNF': nNF,
                'refNFe': refNFe,
                'dhEmi_data': dhEmi_data,
                'dhEmi_hora': dhEmi_hora,
                'dhSaiEnt_data': dhSaiEnt_data,
                'dhSaiEnt_hora': dhSaiEnt_hora,
                'emit_xNome': emit_xNome,
                'emit_UF': emit_UF,
                'emit_CNPJ': emit_CNPJ,
                'dest_UF': dest_UF,
                'dest_xNome': dest_xNome,
                'dest_CNPJ': dest_CNPJ,
                'infCpl': infCpl,
                'chNFe': chNFe,
                'nItem': nItem,
                'caminho': path,
                **totals
            }
            
            # Processing attributes of prod
            prod = det.find('ns:prod', ns)

            nfe['cProd'] = get_optional(prod.find('ns:cProd', ns))
            nfe['cEAN'] = get_optional(prod.find('ns:cEAN', ns))
            nfe['xProd'] = get_optional(prod.find('ns:xProd', ns))
            nfe['NCM'] = get_optional(prod.find('ns:NCM', ns))
            nfe['CEST'] = get_optional(prod.find('ns:CEST', ns))
            nfe['cBenef'] = get_optional(prod.find('ns:cBenef', ns))
            nfe['CFOP'] = get_optional(prod.find('ns:CFOP', ns))
            nfe['uCom'] = get_optional(prod.find('ns:uCom', ns))
            nfe['qCom'] = get_optional(prod.find('ns:qCom', ns))
            nfe['vUnCom'] = get_optional(prod.find('ns:vUnCom', ns))
            nfe['vProd'] = get_optional(prod.find('ns:vProd', ns))
            nfe['vDesc'] = get_optional(prod.find('ns:vDesc', ns))
            nfe['cEANTrib'] = get_optional(prod.find('ns:cEANTrib', ns))
            nfe['uTrib'] = get_optional(prod.find('ns:uTrib', ns))
            nfe['qTrib'] = get_optional(prod.find('ns:qTrib', ns))
            nfe['nRECOPI'] = get_optional(prod.find('ns:nRECOPI', ns))
            if nfe['nRECOPI']:
                nfe['nRECOPI'] = "'" + nfe['nRECOPI'] + "'"
            nfe['cProdANVISA'] = get_optional(prod.find('ns:med/ns:cProdANVISA', ns))
            nfe['vFrete'] = get_optional(prod.find('ns:vFrete', ns))

            # Processing attributes of imposto
            imposto = det.find('ns:imposto', ns)
            nfe['vTotTrib'] = get_optional(imposto.find('ns:vTotTrib', ns))
            
            # Imposto ICMS
            inner_icms = imposto.find('ns:ICMS', ns)
            if inner_icms:
                inner_icms = inner_icms.findall('*')
                if len(inner_icms) > 1: 
                    raise Exception('Múltiplos campos dentro de ICMS')
                inner_icms = inner_icms[0]

                nfe['ICMS_orig'] = get_optional(inner_icms.find('ns:orig', ns))
                nfe['ICMS_CST'] = get_optional(inner_icms.find('ns:CST', ns))
                nfe['ICMS_CSOSN'] = get_optional(inner_icms.find('ns:CSOSN', ns))
                nfe['ICMS_vBCSTRet'] = get_optional(inner_icms.find('ns:vBCSTRet', ns))
                nfe['ICMS_pST'] = get_optional(inner_icms.find('ns:pST', ns))
                nfe['ICMS_vICMSSTRet'] = get_optional(inner_icms.find('ns:vICMSSTRet', ns))
                
                nfe['ICMS_modBC'] = get_optional(inner_icms.find('ns:modBC', ns))
                nfe['ICMS_pRedBC'] = get_optional(inner_icms.find('ns:pRedBC', ns))
                nfe['ICMS_vBC'] = get_optional(inner_icms.find('ns:vBC', ns))
                nfe['ICMS_pICMS'] = get_optional(inner_icms.find('ns:pICMS', ns))
                nfe['ICMS_vICMS'] = get_optional(inner_icms.find('ns:vICMS', ns))
                nfe['ICMS_vBCFCPP'] = get_optional(inner_icms.find('ns:vBCFCPP', ns))
                nfe['ICMS_pFCP'] = get_optional(inner_icms.find('ns:pFCP', ns))
                nfe['ICMS_vFCP'] = get_optional(inner_icms.find('ns:vFCP', ns))
                
                # ICMS filds that may appear in ICMS70 and ICMS10
                nfe['ICMS_modBCST'] = get_optional(inner_icms.find('ns:modBCST', ns))
                nfe['ICMS_pMVAST'] = get_optional(inner_icms.find('ns:pMVAST', ns))
                nfe['ICMS_vBCST'] = get_optional(inner_icms.find('ns:vBCST', ns))
                nfe['ICMS_pICMSST'] = get_optional(inner_icms.find('ns:pICMSST', ns))
                nfe['ICMS_vICMSST'] = get_optional(inner_icms.find('ns:vICMSST', ns))
                nfe['ICMS_vBCFCPST'] = get_optional(inner_icms.find('ns:vBCFCPST', ns))
                nfe['ICMS_pFCPST'] = get_optional(inner_icms.find('ns:pFCPST', ns))
                nfe['ICMS_vFCPST'] = get_optional(inner_icms.find('ns:vFCPST', ns))
            else:
                # Nota de serviço não tem ICMS
                icms_fields = [
                    'ICMS_orig', 'ICMS_CST', 'ICMS_CSOSN', 'ICMS_vBCSTRet', 
                    'ICMS_pST', 'ICMS_vICMSSTRet', 'ICMS_modBC', 'ICMS_pRedBC',
                    'ICMS_vBC', 'ICMS_pICMS', 'ICMS_vICMS', 'ICMS_vBCFCPP',
                    'ICMS_pFCP', 'ICMS_vFCP', 'ICMS_modBCST', 'ICMS_pMVAST',
                    'ICMS_vBCST', 'ICMS_pICMSST', 'ICMS_vICMSST', 
                    'ICMS_vBCFCPST', 'ICMS_pFCPST', 'ICMS_vFCPST'
                ]
                for field in icms_fields:
                    nfe[field] = ''

            # Imposto PIS
            pis = imposto.find('ns:PIS', ns).findall('*')
            if len(pis) > 1: raise Exception('Múltiplos campos dentro de PIS')
            pis = pis[0]
            
            tags = ['CST', 'vBC', 'pPIS', 'vPIS', 'qBCProd', 'vAliqProd']
            for c in pis.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em PIS: ' + tag)
            
            nfe['PIS_CST'] = get_optional(pis.find('ns:CST', ns))
            nfe['PIS_vBC'] = get_optional(pis.find('ns:vBC', ns))
            nfe['PIS_pPIS'] = get_optional(pis.find('ns:pPIS', ns))
            nfe['PIS_vPIS'] = get_optional(pis.find('ns:vPIS', ns))
            nfe['PIS_qBCProd'] = get_optional(pis.find('ns:qBCProd', ns))
            nfe['PIS_vAliqProd'] = get_optional(pis.find('ns:vAliqProd', ns))

            # Imposto COFINS
            cofins = imposto.find('ns:COFINS', ns).findall('*')
            if len(cofins) > 1: raise Exception('Múltiplos campos dentro de COFINS')
            cofins = cofins[0]

            tags = ['CST', 'vBC', 'pCOFINS', 'vCOFINS', 'qBCProd', 'vAliqProd']
            for c in cofins.findall('*'):
                tag = c.tag.split('}')[1]
                if tag not in tags:
                    raise Exception('Campos desconhecidos em COFINS: ' + tag)

            nfe['COFINS_CST'] = get_optional(cofins.find('ns:CST', ns))
            nfe['COFINS_vBC'] = get_optional(cofins.find('ns:vBC', ns))
            nfe['COFINS_pCOFINS'] = get_optional(cofins.find('ns:pCOFINS', ns))
            nfe['COFINS_vCOFINS'] = get_optional(cofins.find('ns:vCOFINS', ns))
            nfe['COFINS_qBCProd'] = get_optional(cofins.find('ns:qBCProd', ns))
            nfe['COFINS_vAliqProd'] = get_optional(cofins.find('ns:vAliqProd', ns))

            # Imposto ICMSUFDest
            icmsufdest = imposto.find('ns:ICMSUFDest', ns)
            if icmsufdest:
                nfe['ICMSUFDest_vBCUFDest'] = get_optional(icmsufdest.find('ns:vBCUFDest', ns))
                nfe['ICMSUFDest_pFCPUFDest'] = get_optional(icmsufdest.find('ns:pFCPUFDest', ns))
                nfe['ICMSUFDest_pICMSUFDest'] = get_optional(icmsufdest.find('ns:pICMSUFDest', ns))
                nfe['ICMSUFDest_pICMSInter'] = get_optional(icmsufdest.find('ns:pICMSInter', ns))
                nfe['ICMSUFDest_pICMSInterPart'] = get_optional(icmsufdest.find('ns:pICMSInterPart', ns))
                nfe['ICMSUFDest_vFCPUFDest'] = get_optional(icmsufdest.find('ns:vFCPUFDest', ns))
                nfe['ICMSUFDest_vICMSUFDest'] = get_optional(icmsufdest.find('ns:vICMSUFDest', ns))
                nfe['ICMSUFDest_vICMSUFRemet'] = get_optional(icmsufdest.find('ns:vICMSUFRemet', ns))
            
            # Imposto IPI
            ipi = imposto.find('ns:IPI', ns)
            if ipi:
                nfe['IPI_cEnq'] = get_optional(ipi.find('ns:cEnq', ns))
                
                ipint = ipi.find('ns:IPINT', ns)
                ipitrib = ipi.find('ns:IPITrib', ns)

                #If IPI is "tributed" or not
                if ipint:
                    nfe['IPI_CST'] = get_optional(ipint.find('ns:CST', ns))
                elif ipitrib:
                    nfe['IPI_CST'] = get_optional(ipitrib.find('ns:CST', ns))
                    
                    #Some documents have vBC and pIPI others have qUnid and vUnid
                    nfe['IPI_vBC'] = get_optional(ipitrib.find('ns:vBC', ns))
                    nfe['IPI_pIPI'] = get_optional(ipitrib.find('ns:pIPI', ns))
                    
                    nfe['IPI_qUnid'] = get_optional(ipitrib.find('ns:qUnid', ns))
                    nfe['IPI_vUnid'] = get_optional(ipitrib.find('ns:vUnid', ns))
                    
                    nfe['IPI_vIPI'] = get_optional(ipitrib.find('ns:vIPI', ns))


            if nItem and int(nItem)==1:
                nfes.append({n: '' for n in nfe.keys()})

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
