from agent import AgenteOfficeIA, abrir_arquivo


def criar_excel_manual():
    """
    Cria arquivo Excel manualmente com dados inseridos pelo usuário
    """
    print("\n" + "=" * 60)
    print("📊 CRIAR PLANILHA EXCEL MANUAL")
    print("=" * 60)

    # Nome do arquivo
    nome_arquivo = input("\n📁 Nome do arquivo (sem extensão): ").strip()
    if not nome_arquivo:
        nome_arquivo = "planilha"
    nome_arquivo = f"{nome_arquivo}.xlsx"

    # Cabeçalhos
    print("\n📋 Defina os cabeçalhos (colunas)")
    print("💡 Digite os nomes separados por vírgula")
    print("   Exemplo: Nome, Email, Telefone, Idade")

    cabecalhos_input = input("\n➤ Cabeçalhos: ").strip()
    if not cabecalhos_input:
        cabecalhos = ["Coluna 1", "Coluna 2", "Coluna 3"]
    else:
        cabecalhos = [c.strip() for c in cabecalhos_input.split(',')]

    # Dados
    print(f"\n📝 Agora insira os dados (você definiu {len(cabecalhos)} colunas)")
    print("💡 Digite os valores separados por vírgula")
    print("💡 Digite 'fim' quando terminar")
    print(f"   Exemplo para {', '.join(cabecalhos)}:")
    print("   João Silva, joao@email.com, 11999999999, 30")

    dados = []
    linha_num = 1

    while True:
        print(f"\n🔹 Linha {linha_num}:")
        entrada = input("➤ ").strip()

        if entrada.lower() in ['fim', 'sair', 'exit', 'q']:
            break

        if not entrada:
            print("⚠️  Linha vazia ignorada")
            continue

        # Divide os valores
        valores = [v.strip() for v in entrada.split(',')]

        # Ajusta para o número de colunas
        if len(valores) < len(cabecalhos):
            valores.extend([''] * (len(cabecalhos) - len(valores)))
        elif len(valores) > len(cabecalhos):
            valores = valores[:len(cabecalhos)]

        dados.append(valores)
        linha_num += 1
        print(f"✅ Linha {linha_num - 1} adicionada")

    # Verifica se há dados
    if not dados:
        print("\n⚠️  Nenhum dado inserido. Criando arquivo com cabeçalhos apenas.")

    # Cria o Excel
    print(f"\n🔧 Criando arquivo {nome_arquivo}...")
    agente = AgenteOfficeIA()
    agente.criar_excel(nome_arquivo, dados, cabecalhos)

    # Pergunta se quer abrir
    abrir = input("\n📂 Abrir arquivo agora? (s/n): ").strip().lower()
    if abrir in ['s', 'sim', 'y', 'yes']:
        abrir_arquivo(nome_arquivo)

    print(f"\n✅ Arquivo criado: {nome_arquivo}")
    print(f"   📊 {len(dados)} linhas de dados")


def criar_word_manual():
    """
    Cria arquivo Word manualmente com conteúdo inserido pelo usuário
    """
    print("\n" + "=" * 60)
    print("📄 CRIAR DOCUMENTO WORD MANUAL")
    print("=" * 60)

    # Nome do arquivo
    nome_arquivo = input("\n📁 Nome do arquivo (sem extensão): ").strip()
    if not nome_arquivo:
        nome_arquivo = "documento"
    nome_arquivo = f"{nome_arquivo}.docx"

    # Título
    titulo = input("\n📌 Título do documento: ").strip()
    if not titulo:
        titulo = "Documento"

    # Conteúdo
    print("\n📝 Agora insira o conteúdo")
    print("💡 Digite os parágrafos (Enter após cada um)")
    print("💡 Digite 'fim' em uma linha vazia para terminar")
    print("💡 Deixe uma linha vazia para adicionar espaço")

    paragrafos = []
    linha_num = 1

    while True:
        print(f"\n🔹 Parágrafo {linha_num}:")
        entrada = input("➤ ").strip()

        if entrada.lower() in ['fim', 'sair', 'exit', 'q']:
            break

        # Permite parágrafos vazios para espaçamento
        paragrafos.append(entrada)

        if entrada:  # Só conta linhas não vazias
            linha_num += 1
            print(f"✅ Parágrafo adicionado")

    # Verifica se há conteúdo
    if not paragrafos or all(not p for p in paragrafos):
        print("\n⚠️  Nenhum conteúdo inserido.")
        paragrafos = ["Este é um documento vazio."]

    # Cria o Word
    print(f"\n🔧 Criando arquivo {nome_arquivo}...")
    agente = AgenteOfficeIA()
    agente.criar_word(nome_arquivo, titulo, paragrafos)

    # Pergunta se quer abrir
    abrir = input("\n📂 Abrir arquivo agora? (s/n): ").strip().lower()
    if abrir in ['s', 'sim', 'y', 'yes']:
        abrir_arquivo(nome_arquivo)

    print(f"\n✅ Arquivo criado: {nome_arquivo}")
    print(f"   📄 {len([p for p in paragrafos if p])} parágrafos")


def menu_principal():
    """
    Menu principal para escolher o que criar
    """
    while True:
        print("\n" + "=" * 60)
        print("🤖 CRIAR ARQUIVOS MANUALMENTE")
        print("=" * 60)
        print("\n📋 Escolha uma opção:")
        print("   1. Criar Excel (.xlsx)")
        print("   2. Criar Word (.docx)")
        print("   3. Sair")

        opcao = input("\n➤ Opção: ").strip()

        if opcao == '1':
            criar_excel_manual()
        elif opcao == '2':
            criar_word_manual()
        elif opcao in ['3', 'sair', 'exit', 'q']:
            print("\n👋 Até mais!")
            break
        else:
            print("\n❌ Opção inválida! Tente novamente.")


if __name__ == '__main__':
    menu_principal()