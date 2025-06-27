import winreg
import sys
import os
import ctypes
from contextlib import contextmanager

@contextmanager
def registry_key(hive, path, access=winreg.KEY_ALL_ACCESS):
    """Context manager para manejo seguro de chaves do registro"""
    key = None
    try:
        key = winreg.CreateKey(hive, path)
        yield key
    except Exception as e:
        print(f"  ❌ Erro ao acessar chave {path}: {e}")
        yield None
    finally:
        if key:
            winreg.CloseKey(key)

def is_admin():
    """Verifica se está executando como administrador"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def backup_registry_keys():
    """Cria backup das configurações atuais (opcional)"""
    print("💾 Criando backup das configurações atuais...")
    # Aqui você poderia implementar um backup das chaves antes de modificar
    return True

def habilitar_vba_office():
    """
    Habilita VBA e macros no Office via registro do Windows
    """
    print("🔧 Configurando acesso VBA e MACROS no Office...")
    print("=" * 50)
    
    # Verificar privilégios
    admin_status = "✅ ADMINISTRADOR" if is_admin() else "⚠️ USUÁRIO COMUM"
    print(f"🔐 Status: {admin_status}")
    
    # Versões do Office (incluindo versões mais recentes)
    office_versions = {
        "16.0": "Office 2016/2019/2021/365",
        "15.0": "Office 2013", 
        "14.0": "Office 2010",
        "12.0": "Office 2007"
    }
    
    # Aplicações para configurar
    applications = ["Word", "Excel", "PowerPoint", "Access"]
    
    success_count = 0
    total_attempts = 0
    
    for version, version_name in office_versions.items():
        print(f"\n📋 Configurando {version_name} ({version})...")
        
        for app in applications:
            total_attempts += 1
            
            # Configurar HKEY_CURRENT_USER (sempre funciona)
            with registry_key(winreg.HKEY_CURRENT_USER, 
                            f"SOFTWARE\\Microsoft\\Office\\{version}\\{app}\\Security") as key:
                if key:
                    try:
                        # Configurações de segurança VBA
                        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                        winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
                        winreg.SetValueEx(key, "Level", 0, winreg.REG_DWORD, 1)
                        
                        # Configurações adicionais para maior compatibilidade
                        winreg.SetValueEx(key, "ExtensionHardening", 0, winreg.REG_DWORD, 0)
                        
                        print(f"  ✅ {app} - Configurações aplicadas (usuário atual)")
                        success_count += 1
                        
                    except Exception as e:
                        print(f"  ❌ {app} - Erro ao configurar: {e}")
            
            # Tentar configurar HKEY_LOCAL_MACHINE (requer admin)
            if is_admin():
                with registry_key(winreg.HKEY_LOCAL_MACHINE, 
                                f"SOFTWARE\\Microsoft\\Office\\{version}\\{app}\\Security") as key:
                    if key:
                        try:
                            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                            winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
                            winreg.SetValueEx(key, "Level", 0, winreg.REG_DWORD, 1)
                            winreg.SetValueEx(key, "ExtensionHardening", 0, winreg.REG_DWORD, 0)
                            print(f"  🔐 {app} - Configuração global aplicada")
                        except Exception as e:
                            print(f"  ⚠️ {app} - Erro na configuração global: {e}")
    
    # Verificar configurações aplicadas
    print(f"\n🔍 Verificando configurações aplicadas...")
    verificar_configuracoes_silencioso()
    
    # Resultado final
    print("\n" + "=" * 50)
    print("🎯 Configuração concluída!")
    print(f"✅ {success_count}/{total_attempts} configurações aplicadas")
    
    if success_count > 0:
        print("\n⚠️ IMPORTANTE - PRÓXIMOS PASSOS:")
        print("   1. Feche TODOS os programas do Office abertos")
        print("   2. Abra o Word/Excel novamente") 
        print("   3. Teste com um documento simples primeiro")
        print("   4. Execute seu script Python")
        print("\n🔒 AVISO DE SEGURANÇA:")
        print("   • Macros foram habilitadas com nível baixo de segurança")
        print("   • Considere reverter após o uso se necessário")
        print("   • Sempre verifique a origem de documentos com macros")
        return True
    else:
        print("\n❌ Nenhuma configuração foi aplicada.")
        print("💡 Sugestões:")
        print("   • Execute como Administrador para configuração global")
        print("   • Verifique se o Office está instalado")
        print("   • Configure manualmente se necessário")
        return False

def verificar_configuracoes_silencioso():
    """Verificação silenciosa para uso interno"""
    office_versions = {
        "16.0": "Office 2016/2019/2021/365",
        "15.0": "Office 2013",
        "14.0": "Office 2010", 
        "12.0": "Office 2007"
    }
    
    applications = ["Word", "Excel", "PowerPoint"]
    configs_ativas = 0
    
    for version, version_name in office_versions.items():
        for app in applications:
            try:
                with registry_key(winreg.HKEY_CURRENT_USER, 
                                f"SOFTWARE\\Microsoft\\Office\\{version}\\{app}\\Security",
                                winreg.KEY_READ) as key:
                    if key:
                        try:
                            access_vbom, _ = winreg.QueryValueEx(key, "AccessVBOM")
                            if access_vbom == 1:
                                configs_ativas += 1
                        except:
                            pass
            except:
                continue
    
    if configs_ativas > 0:
        print(f"  📊 {configs_ativas} configurações VBA ativas encontradas")

def verificar_configuracoes():
    """
    Verifica as configurações atuais sem alterar nada
    """
    print("🔍 Verificando configurações atuais do VBA...")
    print("=" * 50)
    
    office_versions = {
        "16.0": "Office 2016/2019/2021/365",
        "15.0": "Office 2013",
        "14.0": "Office 2010", 
        "12.0": "Office 2007"
    }
    
    applications = ["Word", "Excel", "PowerPoint", "Access"]
    total_configs = 0
    
    for version, version_name in office_versions.items():
        version_found = False
        
        for app in applications:
            try:
                with registry_key(winreg.HKEY_CURRENT_USER, 
                                f"SOFTWARE\\Microsoft\\Office\\{version}\\{app}\\Security",
                                winreg.KEY_READ) as key:
                    if key:
                        try:
                            access_vbom, _ = winreg.QueryValueEx(key, "AccessVBOM")
                            vba_warnings, _ = winreg.QueryValueEx(key, "VBAWarnings")
                            
                            if not version_found:
                                print(f"\n📋 {version_name}:")
                                version_found = True
                            
                            vba_status = "✅ HABILITADO" if access_vbom == 1 else "❌ DESABILITADO"
                            macro_status = "✅ HABILITADO" if vba_warnings == 1 else "❌ DESABILITADO"
                            
                            print(f"  {app}:")
                            print(f"    VBA Access: {vba_status}")
                            print(f"    Macros: {macro_status}")
                            
                            if access_vbom == 1:
                                total_configs += 1
                            
                        except:
                            if not version_found:
                                print(f"\n📋 {version_name}:")
                                version_found = True
                            print(f"  {app}: ⚠️ Não configurado")
                            
            except:
                continue
    
    print(f"\n📊 Resumo: {total_configs} configurações VBA ativas")
    
    if total_configs == 0:
        print("💡 Nenhuma configuração VBA encontrada. Execute a opção 1 para configurar.")

def restaurar_seguranca():
    """
    Restaura configurações de segurança padrão
    """
    print("🔒 Restaurando configurações de segurança padrão...")
    print("=" * 50)
    
    office_versions = {
        "16.0": "Office 2016/2019/2021/365",
        "15.0": "Office 2013",
        "14.0": "Office 2010", 
        "12.0": "Office 2007"
    }
    
    applications = ["Word", "Excel", "PowerPoint", "Access"]
    restored_count = 0
    
    for version, version_name in office_versions.items():
        print(f"\n📋 Restaurando {version_name}...")
        
        for app in applications:
            try:
                with registry_key(winreg.HKEY_CURRENT_USER, 
                                f"SOFTWARE\\Microsoft\\Office\\{version}\\{app}\\Security") as key:
                    if key:
                        # Configurações seguras padrão
                        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 0)
                        winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 2)  # Desabilitar com notificação
                        winreg.SetValueEx(key, "Level", 0, winreg.REG_DWORD, 2)  # Nível médio
                        winreg.SetValueEx(key, "ExtensionHardening", 0, winreg.REG_DWORD, 1)
                        
                        print(f"  ✅ {app} - Segurança restaurada")
                        restored_count += 1
                        
            except Exception as e:
                print(f"  ❌ {app} - Erro: {e}")
    
    print(f"\n🔒 {restored_count} configurações restauradas para padrão seguro")

def main():
    print("🔧 CONFIGURADOR VBA - Office")
    print("=" * 40)
    
    # Verificar argumentos da linha de comando
    if len(sys.argv) > 1:
        if sys.argv[1] == "--check":
            verificar_configuracoes()
            return
        elif sys.argv[1] == "--restore":
            restaurar_seguranca()
            return
    
    print("Escolha uma opção:")
    print("1. 🔓 Configurar VBA (habilitar macros)")
    print("2. 🔍 Verificar configurações atuais")
    print("3. 🔒 Restaurar segurança padrão")
    print("0. 👋 Sair")
    
    try:
        choice = input("\nDigite sua opção (0-3): ").strip()
    except (KeyboardInterrupt, EOFError):
        print("\n👋 Saindo...")
        return
    
    if choice == "1":
        if not is_admin():
            print("\n⚠️ Executando como usuário comum.")
            print("💡 Para configuração global, execute como Administrador.")
            
        success = habilitar_vba_office()
        if success:
            print("\n🚀 Configuração concluída! Teste seu script agora.")
            
    elif choice == "2":
        verificar_configuracoes()
        
    elif choice == "3":
        confirm = input("\n⚠️ Confirma restaurar segurança padrão? (s/N): ").strip().lower()
        if confirm in ['s', 'sim', 'y', 'yes']:
            restaurar_seguranca()
            print("\n✅ Segurança restaurada. Reinicie o Office.")
        else:
            print("❌ Operação cancelada.")
            
    elif choice == "0":
        print("👋 Saindo...")
    else:
        print("❌ Opção inválida! Use 0-3.")

if __name__ == "__main__":
    main()
