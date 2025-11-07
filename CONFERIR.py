#OLÁ!
#PROCEDIMENTO: CONFERIR;
#POR: IZABEL CRISTINA -SESDEC;
# ===============================================
# BIBLIOTECAS UTILIZADAS
# ===============================================
from playwright.sync_api import sync_playwright, TimeoutError
import pyautogui
import keyboard
import time
import sys
import re # Nova biblioteca para expressões regulares (usada na checagem de página)

# ===============================================
# VARIÁVEIS DE CONFIGURAÇÃO
# ===============================================
robo_deve_parar = False
CHROME_DEBUG_URL = "http://localhost:9222"
tecla_de_panico = "Esc" 
NOME_DO_DEPARTAMENTO = "SESDEC-NPA" # Variável movida para o topo para facilitar a manutenção

# ===============================================
# FUNÇÕES DE SEGURANÇA E PÂNICO
# ===============================================
def parar_execucao():
    """Função chamada pela tecla de pânico."""
    global robo_deve_parar
    print("\n!!! TECLA ESC ACIONADA! ENCERRANDO AUTOMACAO !!!")
    robo_deve_parar = True
    
def verificar_panico_e_sair():
    """Verifica se a tecla de pânico foi acionada e encerra o script."""
    global robo_deve_parar
    if robo_deve_parar:
        print("Encerrando a execução por comando de pânico.")
        pyautogui.alert('Tecla ESC acionada. Automação encerrada.', title='Pânico Acionado')
        sys.exit()

# Adiciona a tecla de pânico ao teclado
keyboard.add_hotkey(tecla_de_panico, parar_execucao)
print(f"Robô iniciado. Pressione a tecla '{tecla_de_panico}' a qualquer momento para abortar com seguranca.")

# Aviso inicial para o usuário
pyautogui.confirm(text='Aperte OK quando o E-ESTADO estiver logado na página inicial no depurador do Google Chrome', title='Depurador do Chrome' , buttons=['OK'])

# ===============================================
# INÍCIO DA AUTOMAÇÃO
# ===============================================
with sync_playwright() as p:
    try:
        verificar_panico_e_sair()

        # CONECTAR AO NAVEGADOR JÁ ABERTO:
        print(f"Tentando se conectar ao Chrome na porta de depuração: {CHROME_DEBUG_URL}")
        browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
        print("Conexão estabelecida com sucesso!")

        # OBTER A PÁGINA DO E-ESTADO:
        janela = browser.contexts[0]
        
        if not janela.pages:
            raise Exception("Nenhuma página encontrada na instância de depuração do Chrome.")

        guia = next((page for page in janela.pages if "e-estado.ro.gov.br" in page.url), janela.pages[0])
        print(f"Assumindo o controle da página com o título: '{guia.title()}'")
        
        guia.wait_for_load_state('networkidle', timeout=30000)
        
        # ----------------------------------------------------
        # 1. SEQUÊNCIA DE CLIQUES DE NAVEGAÇÃO
        # ----------------------------------------------------
        print("1. Iniciando navegação: Seleção de Escopo...")

        # CLIQUE 1: Inventário - Membro
        guia.get_by_text("Inventário - Membro").click() 
        time.sleep(0.5)
        verificar_panico_e_sair()

        # CLIQUE 2: Inventário (Item pai do menu lateral)
        guia.get_by_role("link", name="Inventário").click()
        time.sleep(0.5)
        verificar_panico_e_sair()

        # CLIQUE 3: Inventários (Sub-item que lista as comissões)
        guia.get_by_role("link", name="Inventários").click()
        guia.wait_for_load_state('networkidle', timeout=30000)
        time.sleep(0.5)
        verificar_panico_e_sair()

        # CLIQUE 4: Em andamento (Clica no status para entrar no detalhe do inventário)
        guia.get_by_text("Em andamento", exact=True).click() 
        time.sleep(0.5)
        verificar_panico_e_sair()

        # CLIQUE 5: Acessar comissão
        SELETOR_ACESSAR_COMISSAO = 'a:has-text("Acessar comissão")'
        guia.locator(SELETOR_ACESSAR_COMISSAO).click()
        guia.wait_for_load_state('networkidle', timeout=30000)
        time.sleep(0.5)
        verificar_panico_e_sair()


        # ----------------------------------------------------
        # 2. BUSCA, ROLAGEM E ACESSO A BENS (PASSO CRÍTICO)
        # ----------------------------------------------------
        
        print(f"2. Buscando, Rolando e Acessando Bens para o departamento: {NOME_DO_DEPARTAMENTO}")
        
        SELETOR_CAMPO_BUSCA = 'input[placeholder="Pesquisar"]'
        SELETOR_BOTAO_LUPA = '.input-group-text .fa-search'
        SELETOR_ACESSAR_BENS = 'a.btn.btn-info.btn-sm:has-text("Acessar bens")'

        # 1. ROLAGEM INICIAL (Ordem solicitada: Rola primeiro para expor a busca)
        print("   1ª Rolagem: Forçando rolagem para expor os elementos de interação...")
        guia.mouse.wheel(0, 10000) # Rola o viewport principal
        time.sleep(0.5)
        
        # 2. BUSCAR/FILTRAR O DEPARTAMENTO
        print(f"   Buscando o departamento: {NOME_DO_DEPARTAMENTO}")
        
        # Preenche e clica na lupa
        guia.fill(SELETOR_CAMPO_BUSCA, NOME_DO_DEPARTAMENTO)
        print("   Clicando na lupa para aplicar o filtro...")
        guia.click(SELETOR_BOTAO_LUPA)
        
        # Espera estendida para a atualização da lista (5s para sincronização)
        print("   Aguardando para que o filtro se aplique completamente...")
        time.sleep(0.5)  

        # 3. ROLAGEM DE SINCRONIZAÇÃO (Pode ser necessária para re-expor o botão)
        print("   2ª Rolagem: Forçando rolagem para re-expor o botão 'Acessar bens'...")
        guia.mouse.wheel(0, 10000) # Rola novamente até o final
        time.sleep(0.5) 

        # 4. CLICAR no botão "Acessar bens"
        print("   Esperando e Clicando no botão 'Acessar bens'...")
        
        # Espera que o botão esteja visível (após a 2ª rolagem)
        guia.wait_for_selector(SELETOR_ACESSAR_BENS, state='visible', timeout=10000)
        
        localizador_acessar_bens = guia.locator(SELETOR_ACESSAR_BENS)
        
        # CLIQUE 6: Acessar bens
        localizador_acessar_bens.first.click()
        
        guia.wait_for_load_state('networkidle', timeout=30000)
        time.sleep(0.5) 
        verificar_panico_e_sair()
 
        # ----------------------------------------------------
        # 3. LOOP DE PROCESSAMENTO E PAGINAÇÃO (CORREÇÃO DE ESTABILIDADE E INDENTAÇÃO)
        # ----------------------------------------------------

        print("\n3. Iniciando Loop de Processamento e Paginação...")

        # SELETORES
        SELETOR_PROXIMA_PAGINA = 'li.page-item:not(.disabled) button[aria-label="Go to next page"]'
        SELETOR_PAGINA_ATIVA = 'li.page-item.active button[aria-label^="Go to page"]'
        SELETOR_LINHAS_A_PROCESSAR = 'tr:has-text("NÃO CONFERIDO"):has-text("NÃO RECLASSIFICADO")'
        SELETOR_MODAL_DIALOG = '[role="dialog"]' 
        
        # === NOVO: LÓGICA PARA ENCONTRAR A PÁGINA ATUAL CORRETAMENTE ===
        try:
            # Tenta encontrar o botão da página ativa
            botao_pagina_ativa = guia.locator(SELETOR_PAGINA_ATIVA).first
            
            # Pega o atributo aria-label (ex: "Go to page 4")
            aria_label_pagina = botao_pagina_ativa.get_attribute("aria-label")
            
            # Extrai o número da página usando regex (expressão regular)
            match = re.search(r'page (\d+)', aria_label_pagina)
            if match:
                pagina_atual = int(match.group(1))
                print(f"[INÍCIO] Robô detectou que a página atual é a: {pagina_atual}")
            else:
                pagina_atual = 1
                print("[INÍCIO] Não foi possível ler a página ativa. Assumindo a página 1.")
        except:
            pagina_atual = 1
            print("[INÍCIO] Erro ao buscar o seletor da página ativa. Assumindo a página 1.")
        # ===============================================================

        # Inicia o loop externo de páginas
        while True:
            verificar_panico_e_sair()
            print(f"\n=======================================================")
            print(f"| INICIANDO PROCESSAMENTO DA PÁGINA Nº {pagina_atual} |")
            print(f"=======================================================")
            
            # ----------------------------------------------------
            # Loop interno: PROCESSAMENTO DE ITENS
            # ----------------------------------------------------
            
            itens_processados_nesta_pagina = 0
            while True:
                verificar_panico_e_sair()

                # 1. TENTATIVA DE LOCALIZAÇÃO DO PRÓXIMO ITEM PENDENTE
                # Usaremos uma contagem de todos os itens PENDENTES na página
                todos_itens_pendentes = guia.locator(SELETOR_LINHAS_A_PROCESSAR)
                
                # === NOVO: CHECAGEM AGRESSIVA DO FIM DA LISTA ===
                # Aumentamos o tempo de espera para a contagem estabilizar
                time.sleep(1.0) 
                
                if todos_itens_pendentes.count() == 0:
                    print(f"--- [FIM ITENS] Contagem zero de itens pendentes. Processados {itens_processados_nesta_pagina} itens nesta página. ---")
                    break # <--- SAI IMEDIATAMENTE do loop interno (de itens)

                proximo_item = todos_itens_pendentes.first
                
                try:
                    # Garante que o item está anexado (visível) antes de processar
                    # Mantemos o timeout em 5s para dar tempo para o próximo item subir na lista
                    proximo_item.wait_for(state="attached", timeout=5000)
                except TimeoutError:
                    print(f"--- [FIM ITENS] Timeout ao buscar o próximo item. Processados {itens_processados_nesta_pagina} itens nesta página. ---")
                    break # Sai do loop de itens
                
                itens_processados_nesta_pagina += 1
                print(f"--- Processando Item {itens_processados_nesta_pagina} ---")

                # 2. Localiza o botão 'Reclassificar'
                alvo_reclassificar = proximo_item.get_by_role("button", name="Reclassificar")
                
                # Espera o botão ficar visível
                alvo_reclassificar.wait_for(state="visible", timeout=5000) 
                
                # *** CORREÇÃO DEFINITIVA DO CLIQUE ***
                # 1. Rola para garantir a visibilidade
                print("    Rolando a tela para garantir a visibilidade do item...")
                alvo_reclassificar.scroll_into_view_if_needed()
                
                # 2. Aumenta a pausa APÓS a rolagem
                time.sleep(1.0) # Pausa maior (1.5s) para o navegador estabilizar o item no centro.
                
                # 3. CLIQUE 7: RECLASSIFICAR com timeout (re-tentativa interna do Playwright)
                print("    Clicando em 'Reclassificar'...")
                alvo_reclassificar.click(force=True, timeout=10000)
                
                # ----------------------------------------------------
                # 4. SELEÇÃO DOS DROPDOWNS DENTRO DO MODAL (USANDO ÍNDICES nth)
                # ----------------------------------------------------
                
                print("    Aguardando o carregamento COMPLETO e ESTABILIDADE do Modal...")
                
                # Espera que o container do modal fique visível (Esta é a sincronização principal agora)
                guia.locator(SELETOR_MODAL_DIALOG).wait_for(state="visible", timeout=30000) 
                
                # *** PAUSA DE ESTABILIDADE REFORÇADA ***
                # Aumentando a pausa após o modal aparecer
                print("    Pausa para habilitação dos campos...")
                time.sleep(0.5)

                print("    Selecionando opções...")
                
                # Localizador do conjunto de <select> dentro do modal
                seletores_dropdown = guia.locator(SELETOR_MODAL_DIALOG).locator('select')
                
                # *** CORREÇÃO: USANDO ÍNDICES (nth) EM VEZ DE IDS DINÂMICOS ***
                # 1º Select (Classificação)
                seletores_dropdown.nth(0).select_option(label="Servível", force=True) 
                time.sleep(0.5)
                
                # 2º Select (Motivo)
                seletores_dropdown.nth(1).select_option(label="Localizado", force=True)
                time.sleep(0.5)
                
                # 3º Select (Objetivo)
                seletores_dropdown.nth(2).select_option(label="Uso e usufruto do bem", force=True)
                time.sleep(0.5)
                
                # 4º Select (Estado de Conservação)
                seletores_dropdown.nth(3).select_option(label="Bom", force=True)
                time.sleep(0.5)
                
                verificar_panico_e_sair()

                # ----------------------------------------------------
                # 5. CLIQUES FINAIS: SALVAR E CONFERIR (ROLAGEM E PAUSA CORRIGIDA)
                # ----------------------------------------------------

                print("    Salvando e Conferindo...")
                
                # CLIQUE 12: SALVAR (Dentro do modal)
                guia.get_by_role("button", name="Salvar").click()
                guia.wait_for_load_state('networkidle', timeout=30000)
                time.sleep(1.0) # Aumento de pausa para garantir estabilidade pós-salvar
                verificar_panico_e_sair()

                # ROLAGEM CRÍTICA: PÓS-SALVAR
                print("    Rolagem para garantir visibilidade do botão 'Conferir'...")
                SELETOR_BOTAO_CONFERIR = guia.get_by_role("button", name="Conferir").first
                SELETOR_BOTAO_CONFERIR.scroll_into_view_if_needed() # Rola para a visão
                time.sleep(0.5) 
                
                # CLIQUE 13: CONFERIR (Finaliza o item)
                SELETOR_BOTAO_CONFERIR.click() 
                
                # Espera o carregamento e o item sumir
                guia.wait_for_load_state('networkidle', timeout=30000)
                
                # ESPERAR O DESAPARECIMENTO
                print("    Esperando o botão 'Reclassificar' do item atual desaparecer...")
                try:
                    # Esperamos o desaparecimento do botão do item recém-processado
                    alvo_reclassificar.wait_for(state="detached", timeout=10000)
                except TimeoutError:
                    print("[AVISO] O botão Reclassificar do item processado não desapareceu após 10s.") 

                # *** PAUSA EXTRA DE ESTABILIDADE (2 SEGUNDOS) ***
                print("    Pausa de 2 segundos para estabilização da lista antes de buscar o próximo item...")
                time.sleep(2.0) 
                
                verificar_panico_e_sair()

            # ----------------------------------------------------
            # 6. LÓGICA DE PAGINAÇÃO
            # ----------------------------------------------------
            
            print("[PAGINAÇÃO] Verificando se há próxima página...")
            
            proxima_pagina_locator = guia.locator(SELETOR_PROXIMA_PAGINA).first
            
            try:
                # Espera o botão 'Próximo' aparecer e ficar clicável (5s de timeout)
                proxima_pagina_locator.wait_for(state='visible', timeout=5000)
                
                # Se encontrou e está visível, clica
                print(f"[PAGINAÇÃO] Movendo para a página {pagina_atual + 1}...")
                proxima_pagina_locator.click()
                guia.wait_for_load_state('networkidle', timeout=30000)
                time.sleep(2) # Espera o carregamento da nova página
                pagina_atual += 1
                
            except TimeoutError:
                # Se o locator não estiver visível após 5s (fim da paginação ou erro)
                print("[FIM] Botão 'Próximo' não encontrado ou desabilitado. Fim de todas as páginas.")
                break # Sai do loop while True (loop de páginas)

        print("\n[SUCESSO] Processamento de todas as páginas concluído.")
        pyautogui.alert('Automação E-ESTADO concluída com sucesso.', title='Automação Concluída', button='OK')

    except TimeoutError as te:
        print(f"\n[ERRO FATAL - TIMEOUT] O tempo de espera por um elemento expirou. {te}")
        pyautogui.alert('O robô falhou por tempo limite. Verifique se a página carregou ou se o seletor está visível.', title='Erro de Automação', button='OK')
    except Exception as e:
        print(f"\n[ERRO FATAL] Ocorreu um erro inesperado: {e}")
        pyautogui.alert(f'O robô falhou. Erro: {e}', title='Erro de Automação', button='OK')
    finally:
        if 'browser' in locals():
            print("Fechando a conexão com o navegador...")
            browser.close()