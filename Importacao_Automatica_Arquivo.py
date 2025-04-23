"""
Objetivo:
    - O código abaixo agenda a execução de uma função para mover arquivos do NOC em formato .xlsb  de uma pasta origem para um pasta de 
    destino substituindo os arquivos existentes no destino, diariamente às 06:00.
-------------------------------------------------------------------------------------------------------------------------------------
Benefícios:
    - O script é útil para automatizar a movimentação de arquivos em um horário específico, garantindo que os arquivos sejam transferidos 
    e atualizados de forma regular.

-------------------------------------------------------------------------------------------------------------------------------------
Detalhes Técnicos:
    - A pasta de origem é atualizada diariamente pelo sistema do NOC.
    - A pasta de destino é o local centralizado para o processamento e captura das informações pra subir pra outra base de dados.

-------------------------------------------------------------------------------------------------------------------------------------
Agendamento:
    - O script é executado diariamente às 06:00.
    - O script depende da geração dos arquivos .xlsb pelo sistema do NOC antes das 06:00.
"""

# Importa a biblioca schedule para o agendamento de tarefas e funções
import schedule

# Importa biblioteca time para conseguir ter acesso a diversos horários e conversões
import time

# Importa a biblioteca os para obter diversos comandos que vão permitir interagir com o sistema operacional
import os

# Importa a biblioteca shutil para operações de arquivos e pastas
import shutil

# Criação da função mover_arquivos
def mover_arquivos():

    # Define o caminho da pasta de origem, na qual está armazenado o arquivo que desejamos capturar (.xlsb).
    folder_origem = r'C:\Users\F8074436\OneDrive - TIM\Dashboard 2025'

    # Define o caminho da pasta de destino, que será onde os arquivos capturados serão movidos.
    folder_destino = r'\\internal\FileServer\TBR6\CSM\ProcessosSitesInternos\Sites\FiberPlanejamento\PlanFiber2020\Relatórios MIS\01.ETLs\NOC\BASES'

    # Verifica se a pasta de destino existe
    if not os.path.exists(folder_destino):

        # Se a pasta de destino não existir, cria a pasta
        os.makedirs(folder_destino)

    # Lista todos os arquivos da pasta de origem que terminam com a extensão .xlsb.
    arquivos = [f for f in os.listdir(folder_origem) if f.lower().endswith(".xlsb")]

    # Verifica se há lista de arquivos está vazia (ou seja, nenhum arquivo .xlsb encontrado)
    if not arquivos:

        # Imprime uma mensagem informando que nenhum arquivo .xlsb foi encontrado
        print("Nenhum arquivo .xlsb encontrado na pasta de origem.")

    else:

        # Mas caso ele encontre o arquivo, ele vai rodar esse loop (for) para cada arquivo que ele encontrou
        for arquivo in arquivos:

            # Monta o caminho completo do arquivo na pasta de origem
            caminho_origem = os.path.join(folder_origem, arquivo)

            # Monta o caminho completo do arquivo na pasta de destino
            caminho_destino = os.path.join(folder_destino, arquivo)

            # Verifica se o arquivo já existe na pasta de destino
            if os.path.exists(caminho_destino):

                # Se o arquivo existir, imprime uma mensagem informando que será substituído
                print(f"Substituindo arquivo existente: {arquivo}")

                # Remove o arquivo existente na pasta de destino
                os.remove(caminho_destino)

            # Move o arquivo da pasta de origem para a pasta de destino
            try:

                # Move o arquivo
                shutil.move(caminho_origem, caminho_destino)

                # Imprime uma mensagem de sucesso
                print(f"Arquivo movido com sucesso: {arquivo}")

            # Captura qualquer exceção que possa ocorrer durante a movimentação do arquivo
            except Exception as e:

                # Imprime uma mensagem de erro informando qual arquivo falhou e qual erro ocorreu
                print(f"Erro ao mover o arquivo {arquivo}: {e}")

# Agendar a função para ser executada todos os dias às 06:00
schedule.every().day.at("06:00").do(mover_arquivos)

# Loop infinito para manter o agendador em execução
while True:

    # Verifica se há alguma tarefa agendada para ser executada
    schedule.run_pending()

    # Pausa a execução do loop por 1 segundo (para não consumir muitos recursos da CPU)
    time.sleep(1)