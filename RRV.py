import re
import time
import win32com.client
from playwright.async_api import async_playwright
import asyncio
import datetime
from tqdm import tqdm

async def run():
    async with async_playwright() as p:
        data_atual = datetime.datetime.now()
        
        mes = data_atual.month
        ano = data_atual.year
        
        '''
        1
        ======================================================================================
        ESSA PARTE DO CÓDIGO É RESPONSÁVEL POR INICIAR ALL O PROCESSO, IRÁ ABRIR O GOOGLE,
        POR SUAS INFORMAÇÕES DE LOGIN E SENHA E CLICAR NO CAMPO DO BOTÃO QUE FAZ A SOLICITAÇÃO
        DO CÓFIGO PARA ACESSAR O SITE.
        ======================================================================================
        '''
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(color_scheme='dark', record_video_dir='video/')
        page = await browser.new_page()

        await page.goto('https://operacao.ccee.org.br/ui/home')  # Navegue até a página desejada
        # Substitua 'seletor' pelo seletor CSS do elemento
        await page.locator('#mat-input-0').fill('') # INSIRA SEU LOGIN 
        await page.locator('#mat-input-1').fill('') # INSIRA SUA SENHA
        await page.click('.btn-principal')
        await page.locator('xpath=//*[@id="formulario"]/div/input[2]').click()
        
        
        
        '''
        2
        ===================================================================================
        ESSA PARTE DO CÓDIGO PEGA O CÓDIGO NO E-MAIL E COLOCA NO CAMPO PARA ACESSAR A CCEE
        ===================================================================================
        '''
        async def obter_codigo_autorizacao():
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)  # "6" refere-se ao índice da caixa de entrada
            messages = inbox.Items
            last_message = None
            time.sleep(35)
            while True:
                try:
                    message = messages.GetLast()
                    if message != last_message:  # Verifica se há um novo email
                        last_message = message
                        # Usando expressão regular para buscar o código de autorização
                        padrao_codigo = re.compile(r"Codigo de autorizacao: \s*(\d+)", re.IGNORECASE)
                        match = padrao_codigo.search(message.body)
                        if match:
                            codigo_autorizacao = match.group(1)
                            await page.locator('#campoColetaUsuario').fill(codigo_autorizacao) # O CÓDIGO É INSERIDO NESSE CAMPO 
                            await page.locator('#enviar').click()
                            break
                    else:
                        print("Código de autorização não encontrado no email.")
                except Exception as e:
                    print("Red")
                    break
                
                
                
        '''
        3
        ===================================================================================
        ESSA PARTE DO CÓDIGO IRÁ PERCORRER A PÁGINA E CLICAR NOS DEMAIS CAMPOS, ATÉ CHEGAR 
        A PARTE DO DOWLOAD QUE FICA DENTRO DE UM LOOP
        ===================================================================================
        '''
        
        await obter_codigo_autorizacao()
        await page.click('.ng-tns-c37-57')
        await page.locator('xpath=/html/body/div[2]/div[2]/div/div/div/span[2]/span/button').click()
        # await page.pause()
        await page.frame_locator("iframe").get_by_label("Menu drop-down Painel de").click()
        await page.frame_locator("iframe").get_by_label("8.RRV").click()
        await page.frame_locator("iframe").get_by_label("Receita de Venda", exact=True).click()
        await page.frame_locator("iframe").get_by_text("RV002 - Detalhamento da").click()
        
        await page.frame_locator("iframe").get_by_role("textbox", name=f"/0{str(mes + 1)}").click() 
        await page.frame_locator("iframe").get_by_title(f"{ano}/0{str(mes - 3)}").click()
        
        time.sleep(2)
        await page.frame_locator("iframe").get_by_title(f"{ano}/0{str(mes - 3)}").press("Tab")  
        await page.frame_locator('iframe').locator('.data').nth(1).click()
        await page.frame_locator("iframe").get_by_title(f"{ano}_0{str(mes - 3)}_RECEITA DE VENDA FINAL").click()
        
        '''
        4
        ===================================================================================
        ESSA PARTE DO CÓDIGO É A LISTA QUE ALIMENTA O LOOP E COLOCA O NOME NOS DOCUMENTOS,
        CASO SEJA INSERDO ALGUMA USINA, POR FAVOR ADICIONAR NO ÚLTIMO ITEM DA LISTA
        ===================================================================================
        '''
        
        textos = [
    "CARCARA I", "CARCARA II", "CARNAUBA", "EOL ACRE I", "EOL CAICARA I",
    "EOL CAICARA II", "EOL CANUDOS II", "EOL CANUDOS III", "EOL JUNCO I",
    "EOL JUNCO II", "EOL POTIGUAR B31", "EOL POTIGUAR B32", "EOL POTIGUAR B33",
    "EOL VILA ACRE II", "EOL VILA AMAZONAS V", "EOL VILA PARA I", "EOL VILA PARA II",
    "EOL VILA PARA III", "EOL VILA PARAIBA I", "EOL VILA PARAIBA II", "EOL VILA PARAIBA III",
    "EOL VILA PARAIBA IV", "EOLICA CANUDOS II", "EOLICA CANUDOS III", "FOUNDEVER - BARRA FUNDA/SP",
    "OBRAMAX", "REDUTO", "SANTO CRISTO", "SAO JOAO", "SOL SERRA DO MEL I",
    "SOL SERRA DO MEL II", "SOL SERRA DO MEL III", "SOL SERRA DO MEL IV", "SOL SERRA DO MEL V",
    "SOL SERRA DO MEL VI", "TERRAL", "VENTOS DE VILA ACRE II", "VENTOS DE VILA CEARA I",
    "VENTOS DE VILA CEARA II", "VENTOS DE VILA PARAIBA I", "VENTOS DE VILA PARAIBA II",
    "VOLTALIA COM"
]
        
        
        
        '''
        ===================================================================================
        ESSA PARTE DO CÓDIGO É A LISTA QUE ALIMENTA O LOOP E COLOCA O NOME NOS DOCUMENTOS,
        CASO SEJA INSERDO ALGUMA USINA, POR FAVOR ADICIONAR NO ÚLTIMO ITEM DA LISTA
        ===================================================================================
        '''
        
        #  ESSE PRIMEIRO LOOP IRÁ BAIXAR AS PLANILHAS EXCEL 
        
        for i, txt in  tqdm(enumerate(textos),desc="Exportando para Excel", total=len(textos)):
            
            await page.frame_locator("iframe").get_by_text(txt, exact=True).click()
            await page.frame_locator("iframe").get_by_role("button", name="Aplicar").click()
            time.sleep(20)
            await page.frame_locator("iframe").get_by_role("button", name="Opções de Página").click()
            time.sleep(1.5)
            await page.frame_locator("iframe").get_by_label("Exportar para Excel").click()
            time.sleep(1.5)
            async with page.expect_download() as download_info:
                await page.frame_locator("iframe").get_by_label("Exportar Painel Inteiro").click()
            download = await download_info.value
            print(download.suggested_filename)
            # await page.pause()
            sumario, _ = download.suggested_filename.split('.xlsx')
            print(sumario)
            time.sleep(3)
            await page.frame_locator("iframe").get_by_role("link", name="OK").click()
            await download.save_as(f'Documentos/RRV/RRV {download.suggested_filename.replace(sumario, textos[i] + "  0" + str(mes) + "." + str(ano))}')
            
        # Exemplo de uso
        print('ATÉ AQUI, NOS AJUDOU O SENHOR !!!!')
        
        await context.close()
        await browser.close()
        
asyncio.run(run())




