from agent import AgenteOfficeIA, abrir_arquivo
import os


def abrir_arquivo_interativo():
    """
    Função interativa para abrir arquivos
    """
    while True:
        print("\n" + "=" * 50)
        print("📂 ABRIR ARQUIVO")
        print("=" * 50)

        # Lista arquivos no diretório atual
        arquivos = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.docx', '.pdf', '.txt'))]

        if arquivos:
            print("\n📋 Arquivos disponíveis:")
            for i, arquivo in enumerate(arquivos, 1):
                print(f"   {i}. {arquivo}")

        print("\n💡 Digite:")
        print("   - Nome do arquivo (ex: vendas.xlsx)")
        print("   - Número do arquivo da lista")
        print("   - 'sair' para voltar")

        entrada = input("\n➤ ").strip()

        if entrada.lower() in ['sair', 'exit', 'q']:
            print("👋 Até mais!")
            break

        # Se digitou um número
        if entrada.isdigit():
            idx = int(entrada) - 1
            if 0 <= idx < len(arquivos):
                arquivo = arquivos[idx]
            else:
                print("❌ Número inválido!")
                continue
        else:
            arquivo = entrada

        # Verifica se o arquivo existe
        if os.path.exists(arquivo):
            try:
                abrir_arquivo(arquivo)
                print(f"✅ Abrindo: {arquivo}")
            except Exception as e:
                print(f"❌ Erro ao abrir: {e}")
        else:
            print(f"❌ Arquivo '{arquivo}' não encontrado!")
            print("💡 Certifique-se de digitar o nome completo com extensão")


if __name__ == '__main__':
    abrir_arquivo_interativo()