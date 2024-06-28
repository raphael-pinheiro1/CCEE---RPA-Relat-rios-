import re
import time
import win32com.client
from playwright.async_api import async_playwright
import asyncio
import datetime
from tqdm import tqdm

print(f"""
            
            \033[31m███╗   ██╗███████╗████████╗███████╗██╗     ██╗██╗  ██╗
            ████╗  ██║██╔════╝╚══██╔══╝██╔════╝██║     ██║╚██╗██╔╝
            ██╔██╗ ██║█████╗     ██║   █████╗  ██║     ██║ ╚███╔╝ 
            ██║╚██╗██║██╔══╝     ██║   ██╔══╝  ██║     ██║ ██╔██╗ 
            ██║ ╚████║███████╗   ██║   ██║     ███████╗██║██╔╝ ██╗
            ╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝     ╚══════╝╚═╝╚═╝  ╚═╝\033[0m
            
                            PEGUE SUA PIPOCA!!
""")


async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(color_scheme='dark', record_video_dir='video/')
        page = await browser.new_page()

        await page.goto('https://analytics.voltalia.com/accounts/login/?next=/')  # Navegue até a página desejada
        # await page.get_by_role("link", name="Login with Voltalia Account")
        
        await page.get_by_role("link", name="Login with Voltalia Account").click()
        
        await page.get_by_label("Insira seu email ou telefone").fill("r.teixeira@voltalia.com")
        await page.get_by_label("Insira seu email ou telefone").press("Enter")
        await page.get_by_placeholder("Senha").fill("X9GGC98gqv@123")
        await page.get_by_placeholder("Senha").press("Enter")
        await page.get_by_role("button", name="Sim").click()
        await page.get_by_label("Pular").click()
        await page.pause()
        
        
        await context.close()
        await browser.close()
        
asyncio.run(run())

    # page.get_by_role("link", name="Login with Voltalia Account").click()
    # page.get_by_placeholder("Email ou telefone").click()
    # page.get_by_label("Opções de entrada").click()
    # page.get_by_label("Entrar em uma organização").click()
    # page.get_by_role("button", name="Voltar").click()
    # page.get_by_role("button", name="Voltar").click()
    # page.get_by_label("Insira seu email ou telefone").fill("r.teixeira@voltalia.com")
    # page.get_by_label("Insira seu email ou telefone").press("Enter")
    # page.locator("#i0118").press("CapsLock")
    # page.get_by_placeholder("Senha").fill("X9GGC98")
    # page.get_by_placeholder("Senha").press("CapsLock")
    # page.get_by_placeholder("Senha").fill("X9GGC98gqv@123")
    # page.get_by_placeholder("Senha").press("Enter")
    # page.get_by_label("Inserir código").fill("554407")
    # page.get_by_label("Inserir código").press("Enter")
    # page.get_by_label("Não mostrar isso novamente").check()
    # page.get_by_role("button", name="Sim").click()
    # page.get_by_label("Pular").click()