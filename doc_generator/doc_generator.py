import pandas as pd
import math as mt
import time
from datetime import datetime, timedelta
from docx2pdf import convert
from docxtpl import DocxTemplate
import xlwings as xw
from num2words import num2words
from pathlib import Path, PureWindowsPath


class Input_Output():
    @staticmethod
    def get_contract_df(path):
        contract = pd.read_csv(path, sep=';', encoding='latin-1', 
                                usecols=['Nome', 'CNPJ/CPF', 'Endereço',
                                            'Complemento', 'Cidade', 'Estado', 'CEP',
                                            'Potência do Sistema', 'Módulo Fabricante',
                                            'Módulo Modelo', 'Módulo Quantidade', 
                                            'Inversor Fabricante 1 (O número é incremental)',
                                            'Inversor Modelo 1 (O número é incremental)',
                                            'Quantidade de Inversores Utilizados',
                                            'Tipo de Telhado', 'Data de Geração de Proposta',
                                            'Kits Custo Total (Kit Fechado ou Módulo+Inversor+Otimizador)',
                                            'Prestação de serviços técnicos conforme escopo de fornecimento + Materiais extras (Cabos, Disjuntores e DPS)',
                                            'Preço Total Venda',
                                            'Preço Total Venda por Extenso',
                                            'Geração Mensal'])
        contract_df = pd.DataFrame(contract)
        # Renomenado Colunas
        contract_df = contract_df.rename(columns={
                                        'Nome': "cliente_nome",
                                        'CNPJ/CPF': "cliente_cnpj_cpf",
                                        'Endereço': "cliente_endereco",
                                        'Complemento': "cliente_complemento",
                                        'Cidade': "cliente_cidade",
                                        'Estado': "cliente_estado",
                                        'CEP': "cliente_cep",
                                        'Potência do Sistema': "potencia_sistema",
                                        'Módulo Fabricante': "modulo_fabricante",
                                        'Módulo Modelo': "modulo_modelo",
                                        'Módulo Quantidade': "modulo_quantidade",
                                        'Inversor Fabricante 1 (O número é incremental)': "inversor_fabricante_1",
                                        'Inversor Modelo 1 (O número é incremental)': "inversor_modelo_1",
                                        'Quantidade de Inversores Utilizados': "inversores_utilizados",
                                        'Tipo de Telhado': "tipo_telhado",
                                        'Data de Geração de Proposta': "data_de_inicio_utc",
                                        'Data de Geração de Proposta': "data_de_inicio_epoch",
                                        'Kits Custo Total (Kit Fechado ou Módulo+Inversor+Otimizador)': "kits_custo_total",
                                        'Preço Total Venda': "preço",
                                        'Preço Total Venda por Extenso': "preco_por_extenso",
                                        'Geração Mensal': "geraçao_mensal",
                                        'Prestação de serviços técnicos conforme escopo de fornecimento + Materiais extras (Cabos, Disjuntores e DPS)': "prestacao_de_servico"
        })
        
        return contract_df

    @staticmethod
    def get_client_name(wb):
        # Instantiate sheet
        client_sheet = wb.sheets['CLIENT']
        # Get client
        client = client_sheet['B1'].value
        return client

    @staticmethod
    def export_contract_info(wb, contract_df):
        # Instantiate sheet
        sheet = wb.sheets['PANEL']
        # Prepare DataFrame
        contract_df = contract_df.transpose()
        # Output DataFrame to workbook
        sheet['A1'].value = contract_df
        sheet['A1'].options(pd.DataFrame, expand='table').value
        sheet['A1'].value = 'Placeholder'
        sheet['B1'].value = 'Value'
        
    @staticmethod
    def export_contract_docx_pdf(client, contract_df, paths):
        # Get correct template path
        docx_path = Operations.get_correct_template_contract_path(client, contract_df, paths)
        # Instantiate DocxTemplate with contract_template
        document = DocxTemplate(docx_path)
        # Create dict for substitution
        context = {}
        # Get client row
        client_row = contract_df[contract_df['cliente_nome'].str.match(client)].index[0]
        # client_row = contract_df.loc[contract_df['cliente_nome'] == client].index[0]
        # Substitute placeholders with desired info
        for col in contract_df:
            context[col] = str(contract_df.at[client_row, col])
        document.render(context)
        path = paths['output']['contracts'] / Path(f"Contrato - {client}.docx")
        document.save(path)
        convert(path)


class Operations():
    @staticmethod
    def prepare_values(client, contract_df, wb):
        contract_df = Operations.arr_geraçao(client, contract_df)
        contract_df = Operations.epoch_to_time(client, contract_df)
        contract_df = Operations.end_date(client, contract_df)
        contract_df = Operations.captalize(client, contract_df)
        contract_df = Operations.format_cep(client, contract_df)
        contract_df = Operations.warranties(contract_df, wb)
        contract_df = Operations.deadline(contract_df, wb)
        contract_df = Operations.payment(client, contract_df, wb)
        contract_df = Operations.alphabetize(contract_df)
        return contract_df

    @staticmethod
    def get_correct_template_contract_path(client, contract_df, paths):
        cpf_cnpj = contract_df.loc[contract_df['cliente_nome'] == client, "cliente_cnpj_cpf"].values[0]
        if len(cpf_cnpj) != 14:
            for key, value in paths['input'].items():
                if "pj" in key:
                    return value
        else:
            for key, value in paths['input'].items():
                if "pf" in key:
                    return value

    @staticmethod
    def format_to_2_decimals(string_number):
        if not isinstance(string_number, str):
            if "." not in str(string_number):
                string_number = str(string_number) + ",00"
            elif "." == str(string_number)[-2]:
                string_number = str(string_number).replace(".", ",")
                string_number = str(string_number) + "0"
            elif "." == str(string_number)[-3]:
                string_number = str(string_number).replace(".", ",")
            else:
                return None
        else:
            string_number = string_number.replace(".", "")
            string_number = string_number.replace(",", ".")
        return string_number

    @staticmethod
    def format_to_currency(string_number):
        string_1, string_2 = string_number[:-3], string_number[-2:]
        separated_string_1 = []
        concatenated = ""
        for i in range(-1, -len(string_1) - 1, -1):
            if (i % 3) == 0:
                if i + 3 == 0:
                    separated_string_1.append(string_1[i:])
                else:
                    separated_string_1.append(string_1[i:i+3])
            elif i == -len(string_1):
                if len(separated_string_1) != 0:
                    for concat in separated_string_1:
                        concatenated += concat
                    separated_string_1.append(string_1.replace(concatenated, ""))
                else:
                    separated_string_1.append(string_1)
        currency = ""
        for term_index in range(-1, -len(separated_string_1)-1, -1):
            if term_index == -len(separated_string_1):
                currency += separated_string_1[term_index] + "," + string_2
            else:
                currency += separated_string_1[term_index] + "."
        return (currency, string_1, string_2)

    @staticmethod
    def get_currency_formats(string_number):
        string_number = Operations.format_to_2_decimals(string_number)
        currency = Operations.format_to_currency(string_number)
        if currency[2] == '00':
            written = num2words(currency[1], lang='pt_BR') + " reais"   
        else:
            written = num2words(currency[1], lang='pt_BR') + " reais e " + num2words(currency[2], lang='pt_BR') + " centavos" 
        written = written.title()
        written = written.replace("E", "e")
        return (currency[0], written)

    @staticmethod
    def payment(client, contract_df, wb):
        # AM2 Payment
        contract_df['am2_pagamento'] = None
        # Instantiate sheet
        client_sheet = wb.sheets['CLIENT']
        # Provider
        material_cost = contract_df.loc[contract_df['cliente_nome'] == client, 'kits_custo_total'].values[0]
        material_cost = Operations.get_currency_formats(material_cost)
        contract_df['fornecedor_pagamento'] = f"O CONTRATANTE deverá pagar ao fornecedor o valor referente aos custos de materiais e frete. Esses custos totalizam R$ {material_cost[0]} ({material_cost[1]})"
        # AM2
        contract_df['am2_tipo_pagamento'] = client_sheet['B7'].value
        service_cost = contract_df.loc[contract_df['cliente_nome'] == client, 'prestacao_de_servico'].values[0]
        service_cost = Operations.get_currency_formats(service_cost)
        if client_sheet['B7'].value == "À Vista":
            contract_df['am2_pagamento'] = f"O CONTRATANTE deverá pagar os serviços de projeto, execução e comissionamento do sistema fotovoltaico à vista após a conclusão dos serviços, no valor de R$ {service_cost[0]} ({service_cost[1]})"
        elif client_sheet['B7'].value == "Parcelado":
            installments_num = client_sheet['E7'].value
            installments_cost = client_sheet['G7'].value
            installments_cost = Operations.get_currency_formats(installments_cost)
            if client_sheet['D7'].value:
                entry_cost = client_sheet['D7'].value
                entry_cost = Operations.get_currency_formats(entry_cost)
                contract_df['am2_pagamento'] = f"O CONTRATANTE deverá pagar R$ {entry_cost[0]} ({entry_cost[1]}) à vista, após a assinatura do contrato, e o restante em {num2words(installments_num, lang='pt_BR')} parcelas mensais e consecutivas no valor de R$ {installments_cost[0]} ({installments_cost[1]})"
            elif client_sheet['E7'].value:
                contract_df['am2_pagamento'] = f"O CONTRATANTE deverá pagar após a entrega dos materiais, o valor de R$ {service_cost[0]} ({service_cost[1]}) em {num2words(installments_num, lang='pt_BR')} parcelas de R$ {installments_cost[0]} ({installments_cost[1]})"
            else:
                client_sheet['A12'].value = 'ERRO: Possibilidade não considerada.'
        else:
            client_sheet['A12'].value = 'ERRO: Possibilidade não considerada.'
        
        return contract_df

    @staticmethod
    def warranties(contract_df, wb):
        # Instantiate sheet
        client_sheet = wb.sheets['CLIENT']
        # Set warranties
        if client_sheet['B2'].value == "Micro Inversor":
            contract_df['inversor_tipo'] = client_sheet['B2'].value
            if client_sheet['B3'].value:
                contract_df['inversor_garantia'] = str(int(client_sheet['B3'].value))
            else:
                contract_df['inversor_garantia'] = '15'
        elif client_sheet['B2'].value == "Inversor String":
            contract_df['inversor_tipo'] = "Inversor String"
            if client_sheet['B4'].value:
                contract_df['inversor_garantia'] = str(int(client_sheet['B4'].value))
            else:
                contract_df['inversor_garantia'] = '5'
        if client_sheet['B5'].value:
            contract_df['modulo_garantia'] = str(int(client_sheet['B5'].value))
        else:    
            contract_df['modulo_garantia'] = '12'
        return contract_df

    @staticmethod
    def deadline(contract_df, wb):
        # Instantiate sheet
        client_sheet = wb.sheets['CLIENT']
        if client_sheet['B6'].value:
            contract_df['prazo_conclusao'] = str(int(client_sheet['B6'].value))
        else:    
            contract_df['prazo_conclusao'] = '12'
        return contract_df

    @staticmethod
    def arr_geraçao(client, contract_df):
        gm = float(contract_df.loc[contract_df['cliente_nome'] == client, 'geraçao_mensal'].values)
        contract_df.loc[contract_df['cliente_nome'] == client, 'geraçao_mensal_arredondada'] = str(int(mt.floor(gm/10)*10 - 10))
        return contract_df

    @staticmethod
    def epoch_to_time(client, contract_df):
        etstr = contract_df.loc[contract_df['cliente_nome'] == client, 'data_de_inicio_epoch'].values[0]
        epoch_time_str = etstr[0]+etstr[2:5]+etstr[6:9]+etstr[10:13]+etstr[14:17]
        epoch_time = int(int(epoch_time_str)/1000)
        my_time = time.strftime('%d/%m/%Y', time.localtime(epoch_time))
        contract_df.loc[contract_df['cliente_nome'] == client, 'data_de_inicio_utc'] = my_time
        return contract_df

    @staticmethod
    def end_date(client, contract_df):
        sd = contract_df.loc[contract_df['cliente_nome'] == client, 'data_de_inicio_utc'].values[0]
        end_date = datetime(year=int(sd[6:]), month=int(sd[3:5]), day=int(sd[:2])) + timedelta(days=60)
        contract_df['data_de_termino'] = end_date.strftime('%d/%m/%Y')
        return contract_df

    @staticmethod
    def format_cep(client, contract_df):
        cep = str(contract_df.loc[contract_df['cliente_nome'] == client, 'cliente_cep'].values[0])
        cep = cep[:5] + "-" + cep[5:]
        contract_df['cliente_cep'] = cep
        return contract_df

    @staticmethod
    def captalize(client, contract_df):
        preco_sm = contract_df.loc[contract_df['cliente_nome'] == client, 'preco_por_extenso']
        preco_allcaps = preco_sm.values[0].title()
        contract_df.loc[contract_df['cliente_nome'] == client, 'preco_por_extenso'] = preco_allcaps
        return contract_df

    @staticmethod
    def alphabetize(contract_df):
        sorted_cols = []
        for col in contract_df:
            sorted_cols.append(col)
        sorted_cols.sort()
        contract_df = contract_df.reindex(sorted_cols, axis="columns")
        return contract_df


def main():
    paths = {
        'input': {
            'csv': Path(__file__).parents[1] / Path('Data/SolarMarketData.csv'),
            'contract_template_pf': Path(__file__).parents[1] / Path('Templates/Contrato/Template_Contrato_PF.docx'),
            'contract_template_pj': Path(__file__).parents[1] / Path('Templates/Contrato/Template_Contrato_PJ.docx')
        },
        'io':{
            'excel': Path(__file__).parents[0] / Path('doc_generator.xlsm')
        },
        'output': {
            'contracts': Path(__file__).parents[1] / Path('Generated Documents/Contratos')
        }
    }

    io = Input_Output()
    op = Operations()
    
    # Get contract variables from csv file
    contract_df = io.get_contract_df(paths['input']['csv'])
    # Open workbook
    xw.Book(paths['io']['excel']).set_mock_caller()
    # Call workbook
    wb = xw.Book.caller()
    # Get client
    client = io.get_client_name(wb)
    # Prepare values
    contract_df = op.prepare_values(client, contract_df, wb)
    # Export contract variables to excel controler file
    io.export_contract_info(wb, contract_df)
    # Export docx and pdf contract
    io.export_contract_docx_pdf(client, contract_df, paths)



if __name__ == "__main__":
    xw.Book(Path(__file__).parents[0] / Path('doc_generator.xlsm')).set_mock_caller()
    main()
