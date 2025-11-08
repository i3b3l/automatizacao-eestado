# OLÁ!
# PROCEDIMENTO: CONFERIR;
#POR: IZABEL CRISTINA GALVÃO-SESDEC;
#ALTERAÇÕES NO CÓDIGO PODEM SER ACESSADAS NO MEU GITHUB: <https://github.com/i3b3l/automatizacao-eestado>.

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
NOME_DO_DEPARTAMENTO = "SESDEC-NPA"

### MUDANÇA 1: Lista para rastrear itens bloqueados e evitar loops infinitos
ids_bloqueados_a_pular = [] 

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
        # 3. LOOP DE PROCESSAMENTO E PAGINAÇÃO (CORREÇÃO DE ESTABILIDADE E INDENTAÇÃO)
        # ----------------------------------------------------

        print("\n3. Iniciando Loop de Processamento e Paginação...")

        # SELETORES
        SELETOR_PROXIMA_PAGINA = 'li.page-item:not(.disabled) button[aria-label="Go to next page"]'
        SELETOR_PAGINA_ATIVA = 'li.page-item.active button[aria-label^="Go to page"]'
        SELETOR_LINHAS_A_PROCESSAR = 'tr:has-text("NÃO CONFERIDO"):has-text("NÃO RECLASSIFICADO")'
        SELETOR_MODAL_DIALOG = '[role="dialog"]' 
        
        # === LÓGICA PARA ENCONTRAR A PÁGINA ATUAL CORRETAMENTE ===
        try:
            botao_pagina_ativa = guia.locator(SELETOR_PAGINA_ATIVA).first
            aria_label_pagina = botao_pagina_ativa.get_attribute("aria-label")
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
                todos_itens_pendentes = guia.locator(SELETOR_LINHAS_A_PROCESSAR)
                
                time.sleep(2.0) # Pausa de 2s para estabilidade da lista
                
                if todos_itens_pendentes.count() == 0:
                    print(f"--- [FIM ITENS] Contagem zero de itens pendentes. Processados {itens_processados_nesta_pagina} itens nesta página. ---")
                    break # <--- SAI IMEDIATAMENTE do loop interno (de itens)

                proximo_item = todos_itens_pendentes.first
                
                try:
                    proximo_item.wait_for(state="attached", timeout=5000)
                except TimeoutError:
                    print(f"--- [FIM ITENS] Timeout ao buscar o próximo item. Processados {itens_processados_nesta_pagina} itens nesta página. ---")
                    break # Sai do loop de itens
                
                # *** MUDANÇA 2: Extrair o ID do bem ANTES de processar ***
                id_do_bem_atual = "ID_NAO_ENCONTRADO"
                try:
                    # Tenta pegar o ID (número grande) do texto do item
                    id_text = proximo_item.locator("p.font-responsive strong").first.text_content(timeout=1000)
                    match_id = re.search(r'(\d{5,})', id_text) # Busca 5+ dígitos
                    if match_id:
                        id_do_bem_atual = match_id.group(1)
                        print(f"    ID do Bem detectado: {id_do_bem_atual}")
                    else:
                        print(f"    Aviso: Regex não encontrou ID do bem no texto '{id_text}'.")
                except Exception as e:
                    print(f"    Aviso: Falha ao extrair ID do bem. {e}")
                
                # *** MUDANÇA 3: Verificar se o ID já está na lista de pulo ***
                # Isso resolve a REPETIÇÃO do item bloqueado
                if id_do_bem_atual in ids_bloqueados_a_pular:
                    print(f"!!! Item {id_do_bem_atual} está na lista de bloqueados. PULANDO. !!!")
                    continue # Pula para o próximo item
                
                # Se não está na lista, continua o processamento...
                
                print("    Forçando rolagem da página para o topo...")
                guia.evaluate("window.scrollTo(0, 0)")
                time.sleep(0.5) 

                itens_processados_nesta_pagina += 1
                print(f"--- Processando Item {itens_processados_nesta_pagina} (ID: {id_do_bem_atual}) ---")

                # 2. Localiza o botão 'Reclassificar'
                alvo_reclassificar = proximo_item.get_by_role("button", name="Reclassificar")
                
                alvo_reclassificar.wait_for(state="visible", timeout=5000) 
                
                print("    Rolando a tela para garantir a visibilidade do item...")
                alvo_reclassificar.scroll_into_view_if_needed()
                
                time.sleep(2.0) # Pausa maior (2s) para o navegador estabilizar.
                
                # 3. CLIQUE 7: RECLASSIFICAR
                print("    Clicando em 'Reclassificar'...")
                alvo_reclassificar.click(force=True, timeout=10000)
                
                # ----------------------------------------------------
                # 4. PREENCHIMENTO DO MODAL E DETECÇÃO DE ERRO
                # ----------------------------------------------------
                
                print("    Aguardando o modal aparecer...")
                try:
                    guia.locator(SELETOR_MODAL_DIALOG).wait_for(state="visible", timeout=10000) 
                except TimeoutError:
                    print("[ERRO FATAL] Modal de Reclassificação NÃO ABRIU. (Bug do último item?).")
                    continue # Pula esta iteração
                
                print("    Modal abriu. Preenchendo campos...")
                time.sleep(2.0) # Pausa para os campos carregarem

                print("    Selecionando opções...")
                seletores_dropdown = guia.locator(SELETOR_MODAL_DIALOG).locator('select')
                
                seletores_dropdown.nth(0).select_option(label="Servível", force=True) 
                time.sleep(0.2)
                seletores_dropdown.nth(1).select_option(label="Localizado", force=True)
                time.sleep(0.2)
                seletores_dropdown.nth(2).select_option(label="Uso e usufruto do bem", force=True)
                time.sleep(0.2)
                seletores_dropdown.nth(3).select_option(label="Bom", force=True)
                time.sleep(0.2)
                
                verificar_panico_e_sair()

                # ----------------------------------------------------
                # 5. LÓGICA DE SALVAMENTO (ERRO vs SUCESSO)
                # ----------------------------------------------------

                print("    Clicando em 'Salvar'...")
                guia.get_by_role("button", name="Salvar").click()
                
                # Pausa CRÍTICA: Espera 3 segundos para o TOAST de erro aparecer
                print("    Aguardando resposta do servidor (Erro ou Sucesso)...")
                time.sleep(3.0) 

                SELETOR_TOAST_BLOQUEADO = 'div.toast-message:has-text("este bem está bloqueado")'
                
                if guia.locator(SELETOR_TOAST_BLOQUEADO).is_visible():
                    # --- CENÁRIO DE ERRO (O TOAST APARECEU) ---
                    print("\n!!! ERRO DETECTADO: BEM BLOQUEADO. PULANDO ESTE ITEM. !!!")
                    
                    # Adiciona na lista de pulo para evitar loop infinito
                    if id_do_bem_atual not in ids_bloqueados_a_pular:
                        ids_bloqueados_a_pular.append(id_do_bem_atual)
                        print(f"    ID {id_do_bem_atual} adicionado à lista de pulo.")
                    
                    # Clica no botão "Fechar" (vermelho) DENTRO do modal
                    try:
                        print("    Clicando no botão 'Fechar' do modal de erro...")
                        SELETOR_BOTAO_FECHAR_ERRO = 'button.btn-danger:has-text("Fechar")'
                        guia.locator(SELETOR_MODAL_DIALOG).locator(SELETOR_BOTAO_FECHAR_ERRO).click(force=True, timeout=5000)
                    except Exception as e:
                        print(f"    Aviso: Falha ao clicar em 'Fechar'. Tentando 'Escape'. {e}")
                        guia.keyboard.press("Escape") # Tenta fechar com ESC

                    # Espera o modal de erro fechar
                    try:
                        guia.locator(SELETOR_MODAL_DIALOG).wait_for(state="detached", timeout=5000)
                        print("    Modal de erro fechado.")
                    except TimeoutError:
                        print("    Aviso: Modal de erro não fechou.")
                    
                    continue # Pula para o próximo item (NÃO CLICA EM CONFERIR)
                
                else:
                    # --- CENÁRIO DE SUCESSO (Nenhum toast de erro apareceu) ---
                    print("    Salvamento bem-sucedido. Clicando em 'Conferir'...")
                    
                    # Espera o modal de preenchimento fechar sozinho
                    try:
                        guia.locator(SELETOR_MODAL_DIALOG).wait_for(state="detached", timeout=10000)
                        print("    Modal de preenchimento fechado com sucesso.")
                    except TimeoutError:
                        print("    [ALERTA] O Modal demorou para fechar. Tentando 'Escape'...")
                        guia.keyboard.press("Escape")
                        time.sleep(1.0)
                    
                    # *** CORREÇÃO: Revertendo para o seletor original que funciona ***
                    # Clica em "Conferir" na tela principal (busca o primeiro "Conferir" da página)
                    SELETOR_BOTAO_CONFERIR = guia.get_by_role("button", name="Conferir").first
                    SELETOR_BOTAO_CONFERIR.scroll_into_view_if_needed() 
                    time.sleep(0.5) 
                    
                    print("    Clicando em 'Conferir'...")
                    SELETOR_BOTAO_CONFERIR.click() 
                    
                    guia.wait_for_load_state('networkidle', timeout=30000)
                    
                    # Espera o item sumir da lista
                    print("    Esperando o item ser processado e sumir da lista...")
                    try:
                        # alvo_reclassificar foi definido no início do loop do item
                        alvo_reclassificar.wait_for(state="detached", timeout=10000)
                    except TimeoutError:
                        print("[AVISO] O item processado não desapareceu da lista após 10s.") 

                    print("    Pausa de 1 segundo para estabilização da lista...")
                    time.sleep(1.0) 
                    
                    verificar_panico_e_sair()

            # ----------------------------------------------------
            # 6. LÓGICA DE PAGINAÇÃO
            # ----------------------------------------------------
            
            print("[PAGINAÇÃO] Verificando se há próxima página...")
            
            proxima_pagina_locator = guia.locator(SELETOR_PROXIMA_PAGINA).first
            
            try:
                proxima_pagina_locator.wait_for(state='visible', timeout=5000)
                
                print(f"[PAGINAÇÃO] Movendo para a página {pagina_atual + 1}...")
                proxima_pagina_locator.click()
                guia.wait_for_load_state('networkidle', timeout=30000)
                
                print("[PAGINAÇÃO] Esperando o primeiro item da nova página...")
                guia.locator(SELETOR_LINHAS_A_PROCESSAR).first.wait_for(state='attached', timeout=10000)
                
                time.sleep(1) 
                pagina_atual += 1
                
            except TimeoutError:
                print("[FIM] Botão 'Próximo' não encontrado ou desabilitado. Fim de todas as páginas.")
                break # Sai do loop while True (loop de páginas)

        print("\n[SUCESSO] Processamento de todos os páginas concluído.")
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