
import subprocess
import os

# Obter o caminho atual da pasta
caminho_atual = os.getcwd()

# Imprimir o caminho atual
print(f"O caminho atual da pasta é: {caminho_atual}")


# Caminho para o script Python que você deseja executar
caminho_script = "Cdu&Med.py"

# Execute o script usando o comando 'python'
# subprocess.run(["python", caminho_script])
subprocess.run(["python", "Sumario.py"])