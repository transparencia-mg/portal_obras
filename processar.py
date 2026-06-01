from pathlib import Path
import subprocess

REPO = Path(__file__).parent

print("Verificando alterações...")

resultado = subprocess.run(
    ["git", "status", "--porcelain"],
    cwd=REPO,
    capture_output=True,
    text=True
)

if not resultado.stdout.strip():
    print("Nenhuma alteração encontrada.")
    exit()

print("Enviando alterações para o GitHub...")

subprocess.run(
    ["git", "add", "upload"],
    cwd=REPO,
    check=True
)

subprocess.run(
    [
        "git",
        "commit",
        "-m",
        "Atualização automática portal obras"
    ],
    cwd=REPO,
    check=True
)

subprocess.run(
    ["git", "push"],
    cwd=REPO,
    check=True
)

print("GitHub atualizado com sucesso.")