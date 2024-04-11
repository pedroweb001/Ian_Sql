import sys
import os
import mysql.connector
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QFileDialog, QTextEdit, QMessageBox, QDialog, QGridLayout, QLabel, QLineEdit, QPushButton
from PyQt5.QtGui import QKeySequence
from PyQt5.QtCore import Qt
import win32com.client as wincl

class ConectarDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle('Conectar ao Banco de Dados')

        layout = QGridLayout()

        self.hostname_label = QLabel('Host:')
        self.hostname_edit = QLineEdit('localhost')
        self.hostname_edit.setAccessibleName("Host:")
        layout.addWidget(self.hostname_label, 0, 0)
        layout.addWidget(self.hostname_edit, 0, 1)

        self.username_label = QLabel('Nome de Usuário:')
        self.username_edit = QLineEdit('root')
        self.username_edit.setAccessibleName("Nome de Usuário:")
        layout.addWidget(self.username_label, 1, 0)
        layout.addWidget(self.username_edit, 1, 1)

        self.password_label = QLabel('Senha:')
        self.password_edit = QLineEdit()
        self.password_edit.setAccessibleName("Senha:")
        self.password_edit.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_label, 2, 0)
        layout.addWidget(self.password_edit, 2, 1)

        self.port_label = QLabel('Porta:')
        self.port_edit = QLineEdit('3306')  # Valor padrão para a porta
        self.port_edit.setAccessibleName("Porta:")
        layout.addWidget(self.port_label, 3, 0)
        layout.addWidget(self.port_edit, 3, 1)

        self.conectar_button = QPushButton('Conectar')
        self.conectar_button.clicked.connect(lambda: self.conectar())
        layout.addWidget(self.conectar_button, 4, 0, 1, 2)

        self.cancelar_button = QPushButton('Cancelar')
        self.cancelar_button.clicked.connect(self.close)
        layout.addWidget(self.cancelar_button, 5, 0, 1, 2)

        self.setLayout(layout)

    def conectar(self):
        hostname = self.hostname_edit.text()
        username = self.username_edit.text()
        password = self.password_edit.text()
        port = self.port_edit.text()
        print(f"Conectando ao banco de dados com: Host: {hostname}, Usuário: {username}, Senha: {password}, Porta: {port}")
        try:
            conn = mysql.connector.connect(
                host=hostname,
                user=username,
                password=password,
                port=port
            )
            conn.close()
            QMessageBox.information(self, "Sucesso", "Conexão estabelecida com sucesso!")
            return conn  # Retorna a conexão estabelecida
        except mysql.connector.Error as err:
            QMessageBox.critical(self, "Erro", f"Erro ao conectar ao banco de dados: {err}")
            return None

class MeuEditorDeTexto(QTextEdit):
    def __init__(self, conexao, parent=None):
        super().__init__(parent)
        self.conexao = conexao
        self.voz = wincl.Dispatch("SAPI.SpVoice")

    def keyPressEvent(self, event):
        if event.modifiers() == Qt.ShiftModifier and event.key() == Qt.Key_Return:
            texto = self.toPlainText().strip()
            if texto:
                if self.conexao is None:
                    QMessageBox.warning(self, "Aviso", "Você não está conectado a nenhuma instância do banco de dados. Conecte-se primeiro.")
                    return
                self.executar_comandos(texto)
        elif event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Return:
            texto = self.toPlainText().strip()
            if texto:
                if self.conexao is None:
                    QMessageBox.warning(self, "Aviso", "Você não está conectado a nenhuma instância do banco de dados. Conecte-se primeiro.")
                    return
                self.executar_script(texto)
        else:
            super().keyPressEvent(event)

    def executar_comandos(self, texto):
        comandos = texto.split(';')
        for comando in comandos[:-1]:  # Exclui o último item da lista, que é uma string vazia após o último ';'
            comando = comando.strip()
            if comando:
                resultado = self.executar_consulta(comando + ';')  # Adiciona o ';' ao final de cada comando
                if resultado:
                    QMessageBox.information(self, "Comando Executado", resultado)

    def executar_consulta(self, consulta_sql):
        try:
            cursor = self.conexao.cursor()
            if consulta_sql.lower().startswith("drop") or consulta_sql.lower().startswith("create") or consulta_sql.lower().startswith("use"):
                cursor.execute(consulta_sql)
                self.conexao.commit()
                self.voz.Speak("Comando executado com sucesso!")
                return "Comando executado com sucesso!"
            else:
                self.voz.Speak("Comando não executado: não é um comando DDL.")
                return "Comando não executado: não é um comando DDL."
        except Exception as e:
            print(f"Erro ao executar consulta SQL: {e}")
            self.voz.Speak("Erro ao executar o comando.")
            return f"Erro ao executar o comando: {str(e)}"

    def executar_script(self, script):
        try:
            cursor = self.conexao.cursor()
            cursor.execute(script)
            self.conexao.commit()
            self.voz.Speak("Script executado com sucesso!")
            return "Script executado com sucesso!"
        except Exception as e:
            print(f"Erro ao executar script SQL: {e}")
            self.voz.Speak("Erro ao executar o script.")
            return f"Erro ao executar o script: {str(e)}"

    def executar_ddl(self, ddl):
        try:
            cursor = self.conexao.cursor()
            cursor.execute(ddl)
            self.conexao.commit()
            self.voz.Speak("DDL executado com sucesso!")
            return "DDL executado com sucesso!"
        except Exception as e:
            print(f"Erro ao executar comando DDL: {e}")
            self.voz.Speak("Erro ao executar o comando DDL.")
            return f"Erro ao executar o comando DDL: {str(e)}"

class Iprincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Ian Sql')
        self.setGeometry(100, 100, 600, 400)
        self.editor = QTextEdit()
        self.setCentralWidget(self.editor)
        self.editor.hide()

        menu = self.menuBar()
        file_menu = menu.addMenu('&Arquivo')

        abrir_projeto_action = QAction('Abrir Script', self)
        abrir_projeto_action.setShortcut(QKeySequence('CTRL+O'))
        abrir_projeto_action.triggered.connect(self.abrir_projeto)
        file_menu.addAction(abrir_projeto_action)

        criar_projeto_action = QAction('Criar Script', self)
        criar_projeto_action.setShortcut(QKeySequence('Alt+C'))
        criar_projeto_action.triggered.connect(self.criar_projeto)
        file_menu.addAction(criar_projeto_action)

        salvar_como_action = QAction('Salvar Como', self)
        salvar_como_action.setShortcut(QKeySequence('Alt+S'))
        salvar_como_action.triggered.connect(self.salvar_como)
        file_menu.addAction(salvar_como_action)

        salvar_action = QAction('Salvar', self)
        salvar_action.setShortcut(QKeySequence('Ctrl+S'))
        salvar_action.triggered.connect(self.salvar)
        file_menu.addAction(salvar_action)

        fechar_editor_action = QAction('Fechar Editor', self)
        fechar_editor_action.setShortcut(QKeySequence('Ctrl+W'))
        fechar_editor_action.triggered.connect(self.fechar_editor)
        file_menu.addAction(fechar_editor_action)

        sair_action = QAction('Sair', self)
        sair_action.setShortcut(QKeySequence('Alt+F4'))
        sair_action.triggered.connect(self.sair)
        file_menu.addAction(sair_action)

        self.editor.textChanged.connect(self.salvar)

        db_menu = menu.addMenu('&Banco de Dados')

        conectar_action = QAction('Conectar', self)
        conectar_action.triggered.connect(self.conectar_banco)
        db_menu.addAction(conectar_action)

        desconectar_action = QAction('Desconectar', self)
        desconectar_action.triggered.connect(self.desconectar_banco)
        db_menu.addAction(desconectar_action)
        self.show()

    def criar_projeto(self):
        script_dir = os.path.dirname(os.path.realpath(__file__))
        default_folder = os.path.join(script_dir, "scripts")
        if not os.path.exists(default_folder):
            os.makedirs(default_folder)
        folder_path = QFileDialog.getExistingDirectory(self, 'Escolha onde salvar o projeto', default_folder, QFileDialog.ShowDirsOnly)
        if folder_path:
            self.editor.clear()
            self.editor.setFocus()
            self.editor.show()
            print(f"Projeto será criado em: {folder_path}")

    def abrir_projeto(self):
        if self.editor.conexao is None:
            QMessageBox.warning(self, "Aviso", "Você não está conectado a nenhuma instância do banco de dados. Conecte-se primeiro.")
            return

        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Abrir Arquivo', filter='Arquivos SQL (*.sql)')

        if file_path:
            with open(file_path, 'r') as file:
                conteudo = file.read()
                self.editor.setPlainText(conteudo)
            print(f"Arquivo {file_path} aberto com sucesso.")

    def salvar_como(self):
        script_dir = os.path.dirname(os.path.realpath(__file__))
        default_folder = os.path.join(script_dir, "scripts")

        file_path, _ = QFileDialog.getSaveFileName(self, 'Salvar Arquivo', default_folder, "Arquivos SQL (*.sql)")
        if file_path:
            self.file_path = file_path
            with open(self.file_path, 'w') as f:
                texto = self.editor.toPlainText()
                f.write(texto)
            print(f"Arquivo salvo em: {file_path}")

    def salvar(self):
        texto = self.editor.toPlainText()
        if not texto.strip():
            return
        if hasattr(self, 'file_path'):
            with open(self.file_path, 'w') as f:
                f.write(texto)
            print(f"Arquivo salvo em: {self.file_path}")

    def fechar_editor(self):
        texto = self.editor.toPlainText().strip()
        if texto:
            resposta = QMessageBox.question(self, 'Salvar Alterações', 'Deseja salvar as alterações antes de fechar?', QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if resposta == QMessageBox.Yes:
                self.salvar()
            elif resposta == QMessageBox.Cancel:
                return
        self.editor.hide()
        self.editor.clearFocus()
        self.mostrar_janela_principal()

    def sair(self):
        self.close()

    def mostrar_janela_principal(self):
        self.editor.clear()
        self.show()

    def conectar_banco(self):
        dialog = ConectarDialog(self)  # Instanciando o diálogo de conexão
        conexao = dialog.exec_()  # Exibindo o diálogo e obtendo a conexão
        if conexao:
            if hasattr(self, 'editor'):
                self.editor.conexao = conexao  # Atualiza a conexão existente no editor
            else:
                self.editor = MeuEditorDeTexto(conexao)  # Cria um novo editor de texto com a conexão
                self.setCentralWidget(self.editor)  # Define o editor de texto como widget central
            self.show()  # Mostra as alterações na interface

    def desconectar_banco(self):
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Iprincipal()
    ex.mostrar_janela_principal()
    sys.exit(app.exec_())