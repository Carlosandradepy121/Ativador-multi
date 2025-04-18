import sys
import subprocess
import re
import requests
import winreg
import os
import xml.etree.ElementTree as ET
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QVBoxLayout, QWidget, QMessageBox, QProgressBar,
                            QCheckBox, QComboBox, QFrame, QHBoxLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor, QPixmap

class WindowsVersion:
    def __init__(self):
        self.version = None
        self.edition = None
        self.build = None
        self.get_windows_info()

    def get_windows_info(self):
        try:
            # Obter informações do sistema
            result = subprocess.run(['systeminfo'], capture_output=True, text=True)
            output = result.stdout

            # Extrair versão
            version_match = re.search(r'Versão do SO:\s+([\d.]+)', output)
            if version_match:
                self.version = version_match.group(1)

            # Extrair edição
            edition_match = re.search(r'Edição do SO:\s+(.+)', output)
            if edition_match:
                self.edition = edition_match.group(1)

            # Extrair build
            build_match = re.search(r'Build do SO:\s+(\d+)', output)
            if build_match:
                self.build = build_match.group(1)

            # Determinar versão principal
            if "10.0.22000" in self.version or "10.0.22621" in self.version:
                self.version = "Windows 11"
            elif "10.0.10240" in self.version or "10.0.10586" in self.version:
                self.version = "Windows 10"
            elif "6.2" in self.version or "6.3" in self.version:
                self.version = "Windows 8"

            # Normalizar edição
            if "Core" in self.edition:
                self.edition = "Core"
            elif "Professional" in self.edition:
                self.edition = "Pro"
            elif "Enterprise" in self.edition:
                self.edition = "Enterprise"
            elif "Education" in self.edition:
                self.edition = "Education"
            elif "Home" in self.edition:
                self.edition = "Home"

        except Exception as e:
            print(f"Erro ao obter informações do Windows: {str(e)}")

class VolumeKeys:
    def __init__(self):
        self.keys = {
            'Windows 11': {
                'Home': [
                    'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99',
                    'YTMG3-N6DKC-DKB77-7M9GH-8HVX7',
                    '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH',
                    'PVMJN-6DFY6-9CCP6-7BKTT-D3WVR'
                ],
                'Pro': [
                    'VK7JG-NPHTM-C97JM-9MPGT-3V66T',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF'
                ],
                'Enterprise': [
                    'NPPR9-FWDCX-D2C8J-H872K-2YT43',
                    'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4',
                    'YYVX9-NTFWV-6MDM3-9PT4T-4M68B',
                    '44RPN-FTY23-9VTTB-MP9BX-T84FV'
                ],
                'Education': [
                    'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2',
                    '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ',
                    'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY',
                    '84NGF-MHBT6-FXBX8-QWJK7-DRR8H'
                ]
            },
            'Windows 10': {
                'Home': [
                    'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99',
                    'YTMG3-N6DKC-DKB77-7M9GH-8HVX7',
                    '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH',
                    'PVMJN-6DFY6-9CCP6-7BKTT-D3WVR',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
                    'R3BYW-CBNWT-F3JTP-FM942-BTDXY',
                    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
                    'C4M9W-WPRDG-QBB3F-VM9K8-KDQ9Y',
                    '2VCGQ-BRVJ4-2HGJ2-K36X9-J66JG',
                    'MGX79-TPQB9-KQ248-KXR2V-DHRTD',
                    'FJHWT-KDGHY-K2384-93CT7-323RC',
                    '6K2KY-BFH24-PJW6W-9GK29-TMPWP',
                    'N2434-X9D7W-8PF6X-8DV9T-8TYMD',
                    'J2WWN-Q4338-3GF6R-WF6DW-QQV6M',
                    '236TW-X778T-8MV9F-937GT-QVKBB',
                    '87VT2-FY2XW-F7K39-W3T8R-XMFGF',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
                ],
                'Pro': [
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
                    'R3BYW-CBNWT-F3JTP-FM942-BTDXY',
                    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
                    'C4M9W-WPRDG-QBB3F-VM9K8-KDQ9Y',
                    '2VCGQ-BRVJ4-2HGJ2-K36X9-J66JG',
                    'MGX79-TPQB9-KQ248-KXR2V-DHRTD',
                    'FJHWT-KDGHY-K2384-93CT7-323RC',
                    '6K2KY-BFH24-PJW6W-9GK29-TMPWP',
                    'N2434-X9D7W-8PF6X-8DV9T-8TYMD',
                    'J2WWN-Q4338-3GF6R-WF6DW-QQV6M',
                    '236TW-X778T-8MV9F-937GT-QVKBB',
                    '87VT2-FY2XW-F7K39-W3T8R-XMFGF',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
                    'R3BYW-CBNWT-F3JTP-FM942-BTDXY',
                    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
                    'C4M9W-WPRDG-QBB3F-VM9K8-KDQ9Y',
                    '2VCGQ-BRVJ4-2HGJ2-K36X9-J66JG'
                ],
                'Enterprise': [
                    'NPPR9-FWDCX-D2C8J-H872K-2YT43',
                    'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4',
                    'YYVX9-NTFWV-6MDM3-9PT4T-4M68B',
                    '44RPN-FTY23-9VTTB-MP9BX-T84FV',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
                    'R3BYW-CBNWT-F3JTP-FM942-BTDXY',
                    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
                    'C4M9W-WPRDG-QBB3F-VM9K8-KDQ9Y',
                    '2VCGQ-BRVJ4-2HGJ2-K36X9-J66JG',
                    'MGX79-TPQB9-KQ248-KXR2V-DHRTD',
                    'FJHWT-KDGHY-K2384-93CT7-323RC',
                    '6K2KY-BFH24-PJW6W-9GK29-TMPWP',
                    'N2434-X9D7W-8PF6X-8DV9T-8TYMD',
                    'J2WWN-Q4338-3GF6R-WF6DW-QQV6M',
                    '236TW-X778T-8MV9F-937GT-QVKBB',
                    '87VT2-FY2XW-F7K39-W3T8R-XMFGF',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
                ],
                'Education': [
                    'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2',
                    '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ',
                    'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY',
                    '84NGF-MHBT6-FXBX8-QWJK7-DRR8H',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
                    'R3BYW-CBNWT-F3JTP-FM942-BTDXY',
                    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
                    'C4M9W-WPRDG-QBB3F-VM9K8-KDQ9Y',
                    '2VCGQ-BRVJ4-2HGJ2-K36X9-J66JG',
                    'MGX79-TPQB9-KQ248-KXR2V-DHRTD',
                    'FJHWT-KDGHY-K2384-93CT7-323RC',
                    '6K2KY-BFH24-PJW6W-9GK29-TMPWP',
                    'N2434-X9D7W-8PF6X-8DV9T-8TYMD',
                    'J2WWN-Q4338-3GF6R-WF6DW-QQV6M',
                    '236TW-X778T-8MV9F-937GT-QVKBB',
                    '87VT2-FY2XW-F7K39-W3T8R-XMFGF',
                    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
                    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
                    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
                    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
                    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
                ]
            },
            'Windows 8': {
                'Core': [
                    'FB4WR-32NVD-4RW79-XQFWH-CYQG3',
                    'BN3D2-R7TKB-3YPBD-8DRP2-27GG4',
                    'NG4HW-VH26C-733KW-K6F98-J8CK4',
                    'XKY4K-2NRWR-8F6P2-448RF-CRYQH'
                ],
                'Pro': [
                    'NG4HW-VH26C-733KW-K6F98-J8CK4',
                    'XKY4K-2NRWR-8F6P2-448RF-CRYQH',
                    'FB4WR-32NVD-4RW79-XQFWH-CYQG3',
                    'BN3D2-R7TKB-3YPBD-8DRP2-27GG4'
                ],
                'Enterprise': [
                    '32JNW-9KQ84-P47T8-D8GGY-CWCK7',
                    'MHF9N-XY6XB-WVXMC-BTDCT-MKKG7',
                    'JMNMF-RHW7P-DMY6X-RF3DR-X2BQT',
                    'TNM78-863YX-2K6QX-4G2YH-2R9XB'
                ]
            }
        }

    def get_key(self, version, edition):
        if version in self.keys and edition in self.keys[version]:
            # Retorna uma lista de todas as chaves disponíveis
            return self.keys[version][edition]
        return None

class OfficeVersionManager:
    def __init__(self):
        self.versions = {
            'Office 2021 Pro Plus (x64)': {
                'xml_file': 'office2021_proplus_x64.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x64',
                'version': '2021',
                'edition': 'ProPlus'
            },
            'Office 2021 Pro Plus (x86)': {
                'xml_file': 'office2021_proplus_x86.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x86',
                'version': '2021',
                'edition': 'ProPlus'
            },
            'Office 2021 Project (x64)': {
                'xml_file': 'office2021_project_x64.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x64',
                'version': '2021',
                'edition': 'Project'
            },
            'Office 2021 Project (x86)': {
                'xml_file': 'office2021_project_x86.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x86',
                'version': '2021',
                'edition': 'Project'
            },
            'Office 2019 Pro Plus (x64)': {
                'xml_file': 'office2019_proplus_x64.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x64',
                'version': '2019',
                'edition': 'ProPlus'
            },
            'Office 2019 Pro Plus (x86)': {
                'xml_file': 'office2019_proplus_x86.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x86',
                'version': '2019',
                'edition': 'ProPlus'
            },
            'Office 2019 Project (x64)': {
                'xml_file': 'office2019_project_x64.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x64',
                'version': '2019',
                'edition': 'Project'
            },
            'Office 2019 Project (x86)': {
                'xml_file': 'office2019_project_x86.xml',
                'setup_file': 'setup.exe',
                'architecture': 'x86',
                'version': '2019',
                'edition': 'Project'
            }
        }
        self.office_path = os.path.join(os.getcwd(), "Projectoffice24")
        self.create_xml_files()

    def create_xml_files(self):
        if not os.path.exists(self.office_path):
            os.makedirs(self.office_path)

        for version_name, config in self.versions.items():
            xml_path = os.path.join(self.office_path, config['xml_file'])
            
            # Criar o XML base
            root = ET.Element("Configuration")
            
            # Adicionar configurações comuns
            add = ET.SubElement(root, "Add")
            add.set("OfficeClientEdition", config['architecture'])
            add.set("Channel", "PerpetualVL2021" if config['version'] == "2021" else "PerpetualVL2019")
            
            # Configurar produtos
            product = ET.SubElement(add, "Product")
            product.set("ID", self.get_product_id(config['edition'], config['version']))
            
            # Configurar linguagem
            language = ET.SubElement(product, "Language")
            language.set("ID", "pt-br")
            
            # Configurar exibições
            display = ET.SubElement(root, "Display")
            display.set("Level", "None")
            display.set("AcceptEULA", "TRUE")
            
            # Configurar propriedades
            property = ET.SubElement(root, "Property")
            property.set("Name", "AUTOACTIVATE")
            property.set("Value", "1")
            
            # Salvar o arquivo XML
            tree = ET.ElementTree(root)
            tree.write(xml_path, encoding='utf-8', xml_declaration=True)

    def get_product_id(self, edition, version):
        product_ids = {
            'ProPlus': {
                '2021': 'ProPlus2021Volume',
                '2019': 'ProPlus2019Volume'
            },
            'Project': {
                '2021': 'ProjectPro2021Volume',
                '2019': 'ProjectPro2019Volume'
            }
        }
        return product_ids[edition][version]

    def get_version_info(self, version_name):
        return self.versions.get(version_name)

class OfficeInstaller:
    def __init__(self):
        self.office_path = os.path.join(os.getcwd(), "Projectoffice24")
        self.version_manager = OfficeVersionManager()

    def install_office(self, version_name):
        try:
            # Verificar se a pasta existe
            if not os.path.exists(self.office_path):
                return False, "Pasta Projectoffice24 não encontrada!"

            # Obter informações da versão selecionada
            version_info = self.version_manager.get_version_info(version_name)
            if not version_info:
                return False, "Versão do Office não encontrada!"

            # Verificar se o arquivo de configuração existe
            xml_path = os.path.join(self.office_path, version_info['xml_file'])
            if not os.path.exists(xml_path):
                return False, "Arquivo de configuração não encontrado!"

            # Executar o instalador do Office
            setup_path = os.path.join(self.office_path, version_info['setup_file'])
            result = subprocess.run([setup_path, '/configure', xml_path],
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                return True, f"Office {version_name} instalado com sucesso!"
            else:
                return False, f"Erro ao instalar Office: {result.stderr}"

        except Exception as e:
            return False, f"Erro ao instalar Office: {str(e)}"

class RegistryManager:
    def __init__(self):
        self.kms_host = "carlosandradepy-39464.portmap.host"
        self.kms_port = "44258"

    def set_kms_settings(self):
        try:
            # Abrir a chave do registro
            key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform")
            
            # Definir o host do KMS
            winreg.SetValueEx(key, "KeyManagementServiceName", 0, winreg.REG_SZ, self.kms_host)
            
            # Definir a porta do KMS
            winreg.SetValueEx(key, "KeyManagementServicePort", 0, winreg.REG_SZ, self.kms_port)
            
            winreg.CloseKey(key)
            return True, "Configurações do KMS atualizadas com sucesso!"
        except Exception as e:
            return False, f"Erro ao configurar KMS no registro: {str(e)}"

class ActivationThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    status = pyqtSignal(str)

    def __init__(self, install_office=False, version_name=""):
        super().__init__()
        self.install_office = install_office
        self.version_name = version_name

    def run(self):
        try:
            # Verificar privilégios de administrador
            if not self.is_admin():
                self.error.emit("Este programa precisa ser executado como administrador!")
                return

            # Configurar KMS no registro
            self.status.emit("Configurando KMS no registro...")
            registry_manager = RegistryManager()
            success, message = registry_manager.set_kms_settings()
            if not success:
                self.error.emit(message)
                return

            # Instalar Office se solicitado
            if self.install_office:
                self.status.emit("Instalando Microsoft Office...")
                office_installer = OfficeInstaller()
                success, message = office_installer.install_office(self.version_name)
                if not success:
                    self.error.emit(message)
                    return
                self.status.emit(message)

            # Obter informações do Windows
            self.status.emit("Detectando versão do Windows...")
            win_info = WindowsVersion()
            if not win_info.version or not win_info.edition:
                self.error.emit("Não foi possível detectar a versão do Windows!")
                return

            # Determinar versão principal
            windows_version = "Windows 10" if "10" in win_info.version else "Windows 11"
            self.status.emit(f"Windows detectado: {windows_version} {win_info.edition}")

            # Obter lista de chaves de volume
            volume_keys = VolumeKeys()
            keys = volume_keys.get_key(windows_version, win_info.edition)
            if not keys:
                self.error.emit(f"Não foi possível encontrar chaves para {windows_version} {win_info.edition}")
                return

            # Tentar cada chave até encontrar uma que funcione
            total_keys = len(keys)
            for i, key in enumerate(keys):
                self.status.emit(f"Tentando chave {i+1} de {total_keys}...")
                self.progress.emit(int((i / total_keys) * 100))

                # Tentar instalar a chave
                result = subprocess.run(['slmgr', '/ipk', key], 
                                      capture_output=True, text=True)
                
                if result.returncode == 0:
                    self.status.emit(f"Chave {key} instalada com sucesso!")
                    break
                else:
                    self.status.emit(f"Chave {key} falhou, tentando próxima...")

            # Ativar Windows
            self.status.emit("Ativando Windows...")
            result = subprocess.run(['slmgr', '/ato'], 
                                  capture_output=True, text=True)
            if result.returncode != 0:
                self.error.emit(f"Erro ao ativar Windows: {result.stderr}")
                return

            # Verificar status
            self.status.emit("Verificando status da ativação...")
            result = subprocess.run(['slmgr', '/dli'], 
                                  capture_output=True, text=True)
            self.finished.emit(result.stdout)

        except Exception as e:
            self.error.emit(f"Erro inesperado: {str(e)}")

    def is_admin(self):
        try:
            return subprocess.run(['net', 'session'], 
                                capture_output=True).returncode == 0
        except:
            return False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ativador de Windows KMS")
        self.setFixedSize(600, 700)
        
        # Inicializar o gerenciador de versões
        self.version_manager = OfficeVersionManager()
        
        # Configurar estilo
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 3px;
                text-align: center;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
            }
        """)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Cabeçalho
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #0078d4; border-radius: 10px;")
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 20, 20, 20)
        
        # Título
        title = QLabel("Ativador de Windows KMS")
        title.setFont(QFont('Arial', 24, QFont.Weight.Bold))
        title.setStyleSheet("color: white;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(title)
        
        # Subtítulo
        subtitle = QLabel("Ative seu Windows e instale o Office com facilidade")
        subtitle.setFont(QFont('Arial', 12))
        subtitle.setStyleSheet("color: white;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(subtitle)
        
        layout.addWidget(header_frame)
        
        # Frame de conteúdo
        content_frame = QFrame()
        content_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        content_layout = QVBoxLayout(content_frame)
        content_layout.setSpacing(15)
        
        # Seção do Office
        office_label = QLabel("Selecione a versão do Office:")
        office_label.setFont(QFont('Arial', 12, QFont.Weight.Bold))
        content_layout.addWidget(office_label)
        
        self.office_version_combo = QComboBox()
        self.office_version_combo.addItem("Não instalar Office")
        for version_name in self.version_manager.versions.keys():
            self.office_version_combo.addItem(version_name)
        self.office_version_combo.setFixedHeight(35)
        content_layout.addWidget(self.office_version_combo)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #cccccc;")
        content_layout.addWidget(separator)
        
        # Informações do sistema
        system_label = QLabel("Informações do Sistema:")
        system_label.setFont(QFont('Arial', 12, QFont.Weight.Bold))
        content_layout.addWidget(system_label)
        
        self.system_info = QLabel("")
        self.system_info.setWordWrap(True)
        self.system_info.setStyleSheet("padding: 10px; background-color: #f8f8f8; border-radius: 5px;")
        content_layout.addWidget(self.system_info)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(25)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content_layout.addWidget(self.progress_bar)
        
        # Status
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("padding: 10px; background-color: #f8f8f8; border-radius: 5px;")
        content_layout.addWidget(self.status_label)
        
        layout.addWidget(content_frame)
        
        # Botão de ativação
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 0, 0, 0)
        
        self.activate_button = QPushButton("Ativar Windows")
        self.activate_button.setFixedSize(200, 45)
        self.activate_button.setFont(QFont('Arial', 12, QFont.Weight.Bold))
        button_layout.addWidget(self.activate_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(button_frame)
        
        # Conectar sinais
        self.activate_button.clicked.connect(self.activate_windows)

    def activate_windows(self):
        self.status_label.setText("Iniciando processo de ativação...")
        self.activate_button.setEnabled(False)
        self.progress_bar.setValue(0)
        
        selected_version = self.office_version_combo.currentText()
        install_office = selected_version != "Não instalar Office"
        
        self.thread = ActivationThread(install_office, selected_version)
        self.thread.finished.connect(self.activation_success)
        self.thread.error.connect(self.activation_error)
        self.thread.progress.connect(self.update_progress)
        self.thread.status.connect(self.update_status)
        self.thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if value == 100:
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #cccccc;
                    border-radius: 3px;
                    text-align: center;
                    background-color: white;
                }
                QProgressBar::chunk {
                    background-color: #4CAF50;
                }
            """)
        else:
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #cccccc;
                    border-radius: 3px;
                    text-align: center;
                    background-color: white;
                }
                QProgressBar::chunk {
                    background-color: #0078d4;
                }
            """)

    def update_status(self, message):
        self.status_label.setText(message)
        if "erro" in message.lower() or "falha" in message.lower():
            self.status_label.setStyleSheet("padding: 10px; background-color: #ffebee; border-radius: 5px; color: #c62828;")
        elif "sucesso" in message.lower():
            self.status_label.setStyleSheet("padding: 10px; background-color: #e8f5e9; border-radius: 5px; color: #2e7d32;")
        else:
            self.status_label.setStyleSheet("padding: 10px; background-color: #f8f8f8; border-radius: 5px;")

    def activation_success(self, message):
        self.status_label.setText(message)
        self.activate_button.setEnabled(True)
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Sucesso", "Windows ativado com sucesso!")

    def activation_error(self, error_message):
        self.status_label.setText(error_message)
        self.activate_button.setEnabled(True)
        self.progress_bar.setValue(0)
        QMessageBox.critical(self, "Erro", error_message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 