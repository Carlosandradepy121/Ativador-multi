import os
import sys
import subprocess

def main():
    print("=== Ativador Windows KMS - Compilador ===")
    print("\nEste script irá compilar o programa em um executável.")
    print("Certifique-se de que o Python está instalado e adicionado ao PATH.")
    
    input("\nPressione Enter para continuar...")
    
    try:
        # Instalar PyInstaller
        print("\nInstalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
        # Instalar PyQt6
        print("\nInstalando PyQt6...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyqt6"])
        
        # Compilar o programa
        print("\nCompilando o programa...")
        subprocess.check_call([
            "pyinstaller",
            "--noconfirm",
            "--onefile",
            "--windowed",
            "--name=Ativador Windows KMS",
            "ativar_windows_gui.py"
        ])
        
        print("\nCompilação concluída com sucesso!")
        print("O executável está na pasta 'dist'")
        print("\nPressione Enter para sair...")
        input()
        
    except Exception as e:
        print(f"\nErro durante a compilação: {str(e)}")
        print("Verifique se o Python está instalado corretamente.")
        print("\nPressione Enter para sair...")
        input()

if __name__ == "__main__":
    main() 