# precisa instalar o pip install pywin32 e pip install selenium no terminal de comando
import time
import subprocess
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select # Para caixas de <select>
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import win32com.client
import os

# --- Configuração ---
# (Você precisará baixar o "chromedriver" e colocar o caminho aqui,
# ou deixar o Selenium 4+ baixá-lo automaticamente)f
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 70) # Define um tempo de espera máximo (70 seg)

def validar_chamado_no_agilis(chamado, driver, wait):
    try:
        print(f"   Validando chamado: {chamado}")

        # --- PASSO 1: Clicar na LUPA ---
        print("   - Clicando na lupa...")
        try:
            lupa = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[@aria-label='Pesquisa']")
            ))
            lupa.click()
        except TimeoutException:
            lupa = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[.//*[@aria-label='Pesquisa']]")
            ))
            lupa.click()

        print("   - Lupa clicada com sucesso.")
        time.sleep(0.8)

        # --- PASSO 2: Limpar e digitar o chamado ---
        print(f"   - Digitando chamado {chamado}...")
        campo_busca = wait.until(EC.element_to_be_clickable(
            (By.ID, "subheader_search_box")
        ))
        campo_busca.click()
        campo_busca.send_keys(Keys.CONTROL + "a")
        campo_busca.send_keys(Keys.DELETE)
        campo_busca.send_keys(str(chamado))
        campo_busca.send_keys(Keys.RETURN)
        print("   - Enter pressionado.")
        time.sleep(2)

        # --- PASSO 3: Aguardar e clicar no painel colapsável ---
        print("   - Aguardando painel da conversa...")

        # Espera o painel aparecer com aria-expanded="false" (ainda fechado)
        painel = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "z-collapsiblepanel[data-conv_type='reply']")
        ))

        # Verifica se já está expandido, se não, clica para expandir
        aria_expanded = painel.get_attribute("aria-expanded")
        if aria_expanded != "true":
            print("   - Painel fechado, clicando para expandir...")
            header = painel.find_element(
                By.CSS_SELECTOR,
                "div.zcollapsiblepanel__header.zcollapsiblepanel--toggleableheader"
            )
            driver.execute_script("arguments[0].click();", header)  # JS click mais confiável
            print("   - Header clicado via JavaScript.")
        else:
            print("   - Painel já estava expandido.")

        # --- PASSO 4: Aguardar o conteúdo carregar DENTRO do Web Component ---
        print("   - Aguardando conteúdo do painel carregar...")

        # Espera o div.panel-body aparecer DENTRO do z-collapsiblepanel
        try:
            painel_body = WebDriverWait(painel, 15).until(
                lambda d: painel.find_element(By.CSS_SELECTOR, "div.panel-body.p0")
                if painel.find_elements(By.CSS_SELECTOR, "div.panel-body.p0")
                else None
            )
        except TimeoutException:
            # Fallback: tenta buscar no documento inteiro
            print("   - Tentativa fallback: buscando panel-body no documento...")
            painel_body = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "z-collapsiblepanel[data-conv_type='reply'] div.panel-body.p0")
            ))

        conteudo = painel_body.text
        print(f"   - Conteúdo lido: {conteudo[:80]}...")

        # --- PASSO 5: Verificar critérios ---

        # Critério 1: Tem código de rastreio → MANTER
        if "segue o código de rastreio:" in conteudo:
            print(f"   ✅ Chamado {chamado} MANTIDO (tem código de rastreio).")
            return True

        # Critério 2: Encomenda não recebida → REMOVER
        if "Encomenda não recebida na mensageria." in conteudo:
            print(f"   ❌ Chamado {chamado} REMOVIDO (encomenda não recebida).")
            return False

        # Critério 3: Elemento anterior ao último vazio ou inexistente → REMOVER
        paragrafos = painel_body.find_elements(By.TAG_NAME, "p")
        if len(paragrafos) < 2:
            print(f"   ❌ Chamado {chamado} REMOVIDO (menos de 2 parágrafos).")
            return False

        penultimo = paragrafos[-2].text.strip()
        if not penultimo:
            print(f"   ❌ Chamado {chamado} REMOVIDO (penúltimo parágrafo vazio).")
            return False

        # Nenhum critério de manutenção → REMOVER
        print(f"   ❌ Chamado {chamado} REMOVIDO (sem código de rastreio).")
        return False

    except TimeoutException:
        print(f"   ⚠️ Timeout ao validar chamado {chamado}. Mantendo por precaução.")
        return True
    except Exception as e:
        print(f"   ⚠️ Erro inesperado ao validar chamado {chamado}: {e}")
        return True   

# ----------- Parte 1 Agilis ----------- # 
# DEFINA SEUS DADOS DE LOGIN E A URL INICIAL
URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
try:
    # 0. ABRIR A PÁGINA E FAZER LOGIN
    driver.get(URL_INICIAL)
    print(f"Página aberta: {URL_INICIAL}")
    print("Aguardando tela de login...")

    # O seletor mais provável para esse botão é pelo texto.
    # Tentativa 1: Usando By.LINK_TEXT (se for uma tag <a>)
    try:
        selector_login_integrado = (By.LINK_TEXT, "Login Integrado Microsoft")
        wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
        
    # Tentativa 2: Usando By.XPATH (funciona para <button>, <div>, <span>, etc.)
    except:
        print("Não encontrou por LINK_TEXT. Tentando por XPATH...")
        # Este XPATH procura QUALQUER elemento que tenha o texto exato.
        selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
        wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()

    print("0. Cliquei em 'Login Integrado Microsoft'.")
    print("Aguardando autenticação SSO e carregamento da página principal...")

    # Preenche o e-mail (usando a variável segura)
    email_field_microsoft = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0116")) 
    )
    print("Preenchendo e-mail da Microsoft...")
    email_field_microsoft.send_keys(" ")#Coloque o seu email aqui
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    
    # Preenche a senha (usando a variável segura)
    password_field_microsoft = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0118")) 
    )
    print("Preenchendo senha da Microsoft...")
    password_field_microsoft.send_keys(" ")#Coloque a sua senha aqui
    
    # Tenta clicar no botão "Entrar" (com loop anti-stale)
    print("Procurando o botão 'Entrar'...")
    tentativas = 0
    clicado_entrar = False
    while not clicado_entrar and tentativas < 5:
        try:
            entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            entrar_button.click()
            clicado_entrar = True
            print("Botão 'Entrar' clicado.")
        except StaleElementReferenceException:
            tentativas += 1; time.sleep(0.5)
    if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

    # --- ESPERA PELO MFA MANUAL ---
    print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
    print("Aguardando aprovação do MFA no seu celular (até 180s)...")
    
    tentativas = 0
    clicado_manter = False
    while not clicado_manter and tentativas < 5:
        try:
            keep_logged_in_button = WebDriverWait(driver, 180).until( 
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            keep_logged_in_button.click() # Clica "Sim"
            clicado_manter = True
            print("MFA Aprovado! Botão 'Manter conectado' clicado.")
        except StaleElementReferenceException:
            tentativas += 1; time.sleep(0.5)
        except TimeoutException:
             print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
             clicado_manter = False; break
    if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

    print("Login da Microsoft concluído na janela pop-up.")

    # --- O RESTO DO SEU SCRIPT CONTINUA DAQUI ---

    # 1. CLICAR EM 'RELATÓRIOS'
    # Aguarda o menu 'Relatórios' ficar visível após o login
     
    selector_relatorios = (By.LINK_TEXT, "Relatórios")
    wait.until(EC.element_to_be_clickable(selector_relatorios)).click()
    print("1. Cliquei em 'Relatórios'.")
    print("2. Navegando no menu...")

    # 2. IR EM "CONTRATOS - ADM" -> "PRODUTIVIDADE CONTRATOS - ADM"
    # (Provavelmente precisa clicar no primeiro para o segundo aparecer)
    selector_contratos_adm = (By.LINK_TEXT, "Contratos - ADM")
    wait.until(EC.element_to_be_clickable(selector_contratos_adm)).click()
    print("   - Cliquei em 'Contratos - ADM'.")

    selector_produtividade = (By.LINK_TEXT, "Produtividade Contratos - ADM")
    wait.until(EC.element_to_be_clickable(selector_produtividade)).click()
    print("   - Cliquei em 'Produtividade Contratos - ADM'.")

    # 3. APERTAR "EDITAR"
    # (Pode ser por ID, NOME, ou texto. 'name' é um bom chute em apps Zoho)
    selector_editar = (By.CLASS_NAME, "linkborder") # Chute, pode ser "Editar"
    wait.until(EC.element_to_be_clickable(selector_editar)).click()
    print("3. Cliquei em 'Editar'.")

    # 4. PASSO 1 - SELECIONAR "COLETOR DE CUSTO ADM" E MOVER
    # Selecionar o item na lista da esquerda (pelo texto)
    selector_coletor = (By.XPATH, "//option[text()='Coletor de custo ADM']")
    wait.until(EC.element_to_be_clickable(selector_coletor)).click()
    print("4. Selecionei 'Coletor de custo ADM'.")
    

    # O seletor dele provavelmente é um 'class' ou 'onclick'
    selector_seta_direita = (By.CLASS_NAME, "moverightButton") # CHUTE!
    driver.find_element(*selector_seta_direita).click()
    print("   - Cliquei na seta para mover.")

    print("4.5. Expandindo 'Passo 2: Opções de filtragem'...")
    try:
        # O ID 'reportstep2' foi confirmado pela sua imagem do Inspecionar
        selector_opcoes_filtragem = (By.ID, "rcstep2src")
        # Encontra o elemento que você quer clicar (use "reportstep2" que é melhor)
        elemento_clique = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))

        # Usa JavaScript para forçar o clique
        driver.execute_script("arguments[0].click();", elemento_clique)

        print("   - SUCESSO: Cliquei em 'Opções de filtragem' (via JavaScript).")
        
        # Pequena pausa para a animação de expandir terminar
        time.sleep(1) 

    except TimeoutException:
        print("    - FALHA: Não foi possível encontrar 'Passo 2: Opções de filtragem' (ID: reportstep2).")
        # Se este passo falhar, o próximo (clicar no rádio) também vai falhar.
        raise # 'raise' vai parar o script e pular para o bloco 'except Exception'

    # --- PASSO 5: Selecionar o rádio 'Durante' ---
    print("5. Selecionando o filtro 'Durante'...")
    try: 
        # Usando CSS_SELECTOR para encontrar pelo atributo [value='predefined']
        selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
        
        # Espera o rádio ficar clicável
        wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
        print("    - SUCESSO: Filtro 'Durante' selecionado.")

    except TimeoutException:
        print("    - FALHA: Não foi possível encontrar o rádio 'Durante' (CSS_SELECTOR: input[value='predefined']).")
        raise # Para o script se não encontrar

    # 6. APERTAR "EXECUTAR RELATÓRIO"
    selector_executar = (By.ID, "addnew223222")
    wait.until(EC.element_to_be_clickable(selector_executar)).click()
    print("6. Cliquei em 'Executar relatório'.")
    print("--- Relatório executado, aguardando 10s para carregar...")
    time.sleep(3) # Pausa importante para o relatório carregar

    # --- 7. Baixar Relatório XLS Diretamente ---
    print("7. Iniciando o download direto do relatório XLS...")
    try:
        # Localizar o link de exportação pelo ID "exportxls" e clicar nele
        DOWNLOAD_XLS_LINK = (By.ID, "exportxls")
        wait.until(EC.element_to_be_clickable(DOWNLOAD_XLS_LINK)).click()
        print("   - Clique realizado no link 'Exportar arquivo como XLS'.")
    
        # IMPORTANTE: Adicionar uma pausa para o download começar e terminar.
        # A melhor abordagem é verificar a pasta de downloads até o arquivo aparecer.
        # Veja a explicação abaixo sobre como fazer isso.
        print("   - Aguardando o download ser concluído...")
        time.sleep(5) # Pausa simples de 15 segundos. O ideal é usar uma função de verificação.
    
        print("   - Relatório baixado com sucesso!")

    except Exception as e:
        print(f"ERRO ao tentar baixar o relatório XLS: {e}")
        # Adicione aqui o tratamento de erro
    
    print("--- Automação concluída com sucesso! ---")
    time.sleep(1) # Pausa para você ver o resultado

except Exception as e:
    print(f"ERRO: A automação falhou.")
    print(e)

finally:
    #driver.quit() # Comente "driver.quit()" para o navegador não fechar no final 
    print("Script finalizado.")

time.sleep(2)

# --- CONFIGURAÇÃO ---
PASTA_DOWNLOAD = "C:/Users/pedro.henrsilva/Downloads"
NOME_ASSUNTO = "Produtividade Contratos - ADM"
NOME_REMETENTE = "Agilis"
NOME_ANEXO = "Produtividade Contratos - ADM.xls"
NOME_ARQUIVO_FINAL = "Produtividade_EDITADO.xlsx"
# --------------------

def processar_relatorio_email():
    """
    Encontra o relatório baixado, edita no Excel (deletando linhas)
    e cria uma planilha de resumo formatada.
    """
    # Constantes de formatação do Excel
    xlUp = -4162
    xlCellTypeVisible = 12
    xlPasteValues = -4163 
    xlOpenXMLWorkbook = 51
    xlCenter = -4108           
    xlTop = -4160              
    xlContinuous = 1           
    # -----------------------------------------------
    arquivo_salvo_path = None
    excel = None
    wb = None

    # --- 1. ENCONTRAR O ARQUIVO MAIS RECENTE NA PASTA DE DOWNLOADS --- (NOVO BLOCO)
    try:
        print(f"Procurando o relatório mais recente na pasta: {PASTA_DOWNLOAD}")

        # Lista todos os arquivos na pasta que terminam com .xls (ignorando maiúsculas/minúsculas)
        arquivos_xls = [f for f in os.listdir(PASTA_DOWNLOAD) if f.lower().endswith('.xls')]

        if not arquivos_xls:
            raise FileNotFoundError(f"Nenhum arquivo .xls encontrado na pasta {PASTA_DOWNLOAD}")

        # Monta o caminho completo para cada arquivo e encontra o mais recente pela data de modificação
        caminhos_completos = [os.path.join(PASTA_DOWNLOAD, f) for f in arquivos_xls]
        arquivo_salvo_path = max(caminhos_completos, key=os.path.getmtime)
        
        print(f"Arquivo mais recente encontrado: {arquivo_salvo_path}")

    except Exception as e:
        print(f"ERRO ao tentar encontrar o arquivo de relatório na pasta de downloads: {e}")
        return
    # --- 2. ABRIR O EXCEL E EDITAR O ARQUIVO ---
    try:
        print("\n--- Iniciando edição no Excel ---")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  
        excel.DisplayAlerts = False 

        wb = excel.Workbooks.Open(arquivo_salvo_path)
        ws = wb.Worksheets(1) # ws = Planilha de dados processada

        # --- 2.1 a 2.6. Limpeza (SEM AutoFilter) ---
        print("Editando: Excluindo linhas 1-8...")
        ws.Rows("1:8").Delete()
        
        print("Editando: Removendo logos/imagens...")
        for shape in ws.Shapes:
            shape.Delete()
        
        print("Editando: Ajustando altura e largura...")
        ws.Rows.RowHeight = 12
        ws.Columns.ColumnWidth = 12
        
        # --- (EXISTENTE) 2.7. Deletar linhas que NÃO SÃO de "Correspondência"
        print("Editando: Removendo linhas que não são de 'Solicitação de Envio de Correspondência'...")
        filtro_texto = "Solicitação de Envio de Correspondência"

        # Achar a última linha (baseado na Coluna I, campo 9)
        # OBS: No seu código original, o comentário dizia Coluna H, mas o código usava 9 (que é a coluna I).
        # Mantive o uso da coluna 9 (I).
        last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        print(f"Verificando {last_row} linhas (de baixo para cima)...")

        # Loop de baixo para cima para deletar com segurança
        for i in range(last_row, 1, -1):  # (de last_row até a linha 2)
            cell_value = str(ws.Cells(i, 9).Value)  # Coluna I

            # Se o texto NÃO for encontrado, deleta a linha
            if filtro_texto not in cell_value:
                ws.Rows(i).Delete()
        print("Linhas indesejadas removidas.")


        # --- (NOVO) 2.8. Deletar linhas com horários específicos se for depois do meio-dia ---
        print("\nIniciando verificação de horário para remoção adicional...")

        # Pega a hora atual (apenas o componente da hora, em formato 24h)
        hora_atual = datetime.now().hour

        # Condição: Se a hora atual for maior que 12 (ou seja, 13h em diante)
        if hora_atual >= 12:
            print(f"A hora atual ({hora_atual}:00) é maior que 12:00. Removendo correspondências matutinas.")
            
            # Lista dos horários que queremos remover
            horarios_para_excluir = ["11", "10", "09:", "08:", "07:", "06:"]
            
            # É importante recalcular a última linha, pois a primeira limpeza pode ter alterado o total.
            # Usaremos a coluna K (11) como referência para esta verificação.
            last_row = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
            print(f"Verificando {last_row} linhas (de baixo para cima) na coluna K...")

            # Loop de baixo para cima novamente
            for i in range(last_row, 1, -1):
                cell_value_k = str(ws.Cells(i, 11).Value)  # Pega o valor da Coluna K

                # Verifica se algum dos horários da lista está no texto da célula
                for horario in horarios_para_excluir:
                    if horario in cell_value_k:
                        ws.Rows(i).Delete()
                        # Uma vez que a linha é deletada, podemos parar de verificar outros horários
                        # e passar para a próxima linha (i-1).
                        break
                        
            print("Remoção de correspondências matutinas concluída.")

        else:
            # Se não for depois do meio-dia, apenas informa e não faz nada.
            print(f"A hora atual ({hora_atual}:00) não é maior que 12:00. Nenhuma remoção adicional será feita.")
        
        # 2.9. Adicionar fórmulas (Agora sem SpecialCells)
        print("Editando: Adicionando fórmulas nas colunas M-R...")
        # Achar a *nova* última linha, após as deleções
        last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        print(f"Última linha de dados (pós-deleção): {last_row}")
        if last_row > 1:

            formula_N = (
                '=LET('
                'txt,F2,'
                'start,IFERROR(SEARCH("Código:",txt)+7,'
                'IFERROR(SEARCH("Coletor de Custo ADM",txt)+22,'
                'IFERROR(SEARCH("Centro de custo:",txt)+17,'
                'IFERROR(SEARCH(". Código:",txt)+8,"")))),'
                'raw,IF(start="","",MID(txt,start,999)),'
                'clean1,IF(raw="","",SUBSTITUTE(raw,CHAR(160)," ")),'
                'clean2,SUBSTITUTE(clean1,CHAR(10)," "),'
                'clean3,SUBSTITUTE(clean2,CHAR(13)," "),'
                'chunk,IF(raw="","",TRIM(clean3)),'
                'IF(chunk="","",'
                '   IF(ISNUMBER(--LEFT(chunk,1)),'
                '      LEFT(chunk,IFERROR(SEARCH(" ",chunk)-1,LEN(chunk))),'
                '      LEFT(chunk,10)'
                '   )'
                ')'
                ')'
            )

            # Aplica só em N2 e copia para baixo
            ws.Range("N2").Formula = formula_N
            if last_row > 2:
                ws.Range("N2").Copy(ws.Range(f"N3:N{last_row}"))

            # Coluna O: apenas referencia N
            ws.Range(f"O2:O{last_row}").FormulaLocal = "=N2"

            ws.Range(f"P2:P{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Correspondência:\")"
            ws.Range(f"Q2:Q{last_row}").FormulaLocal = "=TEXTOANTES(P2;\"Cidade\")"

            ws.Range(f"R2:R{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Documentos:\")"
            ws.Range(f"S2:S{last_row}").FormulaLocal = "=TEXTOANTES(R2;\"*\")"
            print("Fórmulas aplicadas com sucesso.")

        # 2.11 Criar Planilha de Resumo
        print("Criando: Planilha de Resumo...")
        ws_summary = wb.Worksheets.Add(After=ws)
        ws_summary.Name = "Resumo"
        ws_summary.Activate()

        # 2.12 Montar o Layout do Resumo
        print("Criando: Layout do Resumo (Título e Cabeçalho)...")
        today_date = time.strftime("%d/%m/%Y")
        
        title_range = ws_summary.Range("A1:D1")
        title_range.Merge()
        title_range.Value = f"MRV - DATA - {today_date}"
        title_range.Font.Bold = True
        title_range.HorizontalAlignment = xlCenter 
        
        ws_summary.Range("A2").Value = "Centro de Custo"
        ws_summary.Range("B2").Value = "Chamado"
        ws_summary.Range("C2").Value = "Serviço"
        ws_summary.Range("D2").Value = "Quantidade"
        summary_header = ws_summary.Range("A2:D2")
        summary_header.Font.Bold = True
        summary_header.Font.Color = 16777215 
        summary_header.Interior.Color = 12611584 
        summary_header.AutoFilter()
        ws_summary.Columns("A:D").ColumnWidth = 22

        # 2.13 Copiar Dados Visíveis para o Resumo (Agora sem SpecialCells)
        print("Copiando dados filtrados para o Resumo (somente valores)...")
        if last_row > 1:
            ws.Range(f"O2:O{last_row}").Copy()
            ws_summary.Range("A3").PasteSpecial(Paste=xlPasteValues)

            ws.Range(f"B2:B{last_row}").Copy()
            ws_summary.Range("B3").PasteSpecial(Paste=xlPasteValues)

            ws.Range(f"Q2:Q{last_row}").Copy()
            ws_summary.Range("C3").PasteSpecial(Paste=xlPasteValues)

            ws.Range(f"S2:S{last_row}").Copy()
            ws_summary.Range("D3").PasteSpecial(Paste=xlPasteValues)
            print("Dados copiados com sucesso.")
        else:
            print("AVISO: Nenhum dado filtrado para copiar.")
        
        excel.Application.CutCopyMode = False
        
        # 2.13.1 Adicionar Rodapé Dinâmico
        print("Criando: Rodapé do Resumo...")
        last_summary_row = ws_summary.Cells(ws_summary.Rows.Count, "A").End(xlUp).Row
        footer_row = max(last_summary_row + 2, 57)
        
        footer_cliente = ws_summary.Range(f"A{footer_row}:B{footer_row + 1}")
        footer_cliente.Merge()
        footer_cliente.Value = "CLIENTE:"
        footer_cliente.Font.Bold = True
        footer_cliente.VerticalAlignment = xlTop 

        footer_agf = ws_summary.Range(f"C{footer_row}:D{footer_row + 1}")
        footer_agf.Merge()
        footer_agf.Value = "AGF:"
        footer_agf.Font.Bold = True
        footer_agf.VerticalAlignment = xlTop 
        
        # 2.13.2 Adicionar Bordas
        print("Criando: Bordas da planilha Resumo...")
        final_used_row = footer_row + 1
        full_range = ws_summary.Range(f"A1:D{final_used_row}")
        full_range.Borders.LineStyle = xlContinuous
        
        # 2.14 Ativar resumo (NÃO precisamos mais limpar o filtro)
        # --- VALIDAÇÃO DOS CHAMADOS NO AGILIS ---
        print("\n--- Iniciando validação dos chamados no Agilis ---")

        wait_validacao = WebDriverWait(driver, 30)

        try:
            driver.get("https://agilis.mrv.com.br/HomePage.do?view_type=my_view")

            # Pega a última linha do resumo
            last_summary_row = ws_summary.Cells(ws_summary.Rows.Count, "B").End(xlUp).Row
            print(f"Total de chamados para validar: {last_summary_row - 2}")

            # Loop de baixo para cima para deletar com segurança
            for i in range(last_summary_row, 2, -1):  # De baixo para cima, a partir da linha 3
                chamado = ws_summary.Cells(i, 2).Value  # Coluna B

                if chamado and str(chamado).strip():
                    manter = validar_chamado_no_agilis(str(chamado).strip(), driver, wait_validacao)
                    if not manter:
                        ws_summary.Rows(i).Delete()
                        print(f"   Linha {i} deletada do Resumo.")
                else:
                    print(f"   Linha {i} ignorada (chamado vazio).")

        finally:
            driver.quit()
            print("--- Validação dos chamados concluída ---")

        # --- 3. SALVAR O ARQUIVO FINAL ---
        # Definição do caminho correto
        caminho_final = r"C:\Users\pedro.henrsilva\OneDrive - MRV\Área de Trabalho\AUTOMATIZAR_RELATORIO_CORREIOS"

        # Combina o diretório com o nome do arquivo
        caminho_completo = os.path.join(caminho_final, NOME_ARQUIVO_FINAL)

        # Salva o arquivo passando o caminho completo no primeiro argumento
        wb.SaveAs(caminho_completo, FileFormat=51) # 51 corresponde ao xlOpenXMLWorkbook (.xlsx)

        print(f"\n--- SUCESSO! ---")
        print(f"Arquivo editado e com resumo salvo como: {caminho_completo}")

    except Exception as e:
        print(f"ERRO durante a edição no Excel: {e}")

if __name__ == "__main__":
    processar_relatorio_email()