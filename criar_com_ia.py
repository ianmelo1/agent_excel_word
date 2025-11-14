from agent import AgenteOfficeIA, abrir_arquivo
import json
import os
from dotenv import load_dotenv

# Carrega variáveis do .env
load_dotenv()


def criar_excel_com_ia():
    """
    Cria Excel automaticamente usando IA para gerar dados
    """
    print("\n" + "=" * 60)
    print("🤖 CRIAR EXCEL COM IA")
    print("=" * 60)

    # Verifica API key
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("\n⚠️  GOOGLE_API_KEY não configurada!")
        print("💡 Configure com: export GOOGLE_API_KEY='sua-chave'")
        api_key = input("\n🔑 Ou cole sua API key aqui: ").strip()
        if not api_key:
            print("❌ Cancelado.")
            return

    agente = AgenteOfficeIA(api_key=api_key)

    # Nome do arquivo
    nome_arquivo = input("\n📁 Nome do arquivo (sem extensão): ").strip()
    if not nome_arquivo:
        nome_arquivo = "planilha_ia"
    nome_arquivo = f"{nome_arquivo}.xlsx"

    # Descrição do que o usuário quer
    print("\n📝 Descreva o que você quer na planilha")
    print("💡 Exemplos:")
    print("   - Lista de 10 produtos com preços e categorias")
    print("   - Tabela de vendas mensais de 2024")
    print("   - Lista de funcionários com departamento e salário")
    print("   - Controle de estoque com 15 itens")

    descricao = input("\n➤ O que você quer: ").strip()
    if not descricao:
        print("❌ Descrição vazia. Cancelado.")
        return

    # Número de linhas
    num_linhas = input("\n🔢 Quantas linhas de dados? (padrão: 10): ").strip()
    if not num_linhas or not num_linhas.isdigit():
        num_linhas = 10
    else:
        num_linhas = int(num_linhas)

    # Gera dados com IA
    print(f"\n🤖 Gerando dados com IA...")
    print("⏳ Aguarde...")

    prompt = f"""Crie dados para uma planilha Excel baseado nesta descrição:

"{descricao}"

Gere EXATAMENTE {num_linhas} linhas de dados.

Retorne APENAS um JSON válido neste formato (sem markdown, sem explicações):
{{
    "cabecalhos": ["Coluna1", "Coluna2", "Coluna3"],
    "dados": [
        ["valor1", "valor2", "valor3"],
        ["valor1", "valor2", "valor3"]
    ]
}}

IMPORTANTE:
- Gere dados realistas e variados
- Use valores apropriados para cada coluna
- EXATAMENTE {num_linhas} linhas em "dados"
- Retorne APENAS o JSON, sem texto adicional"""

    try:
        resposta = agente.perguntar_ia(prompt)

        # Tenta extrair JSON da resposta
        resposta = resposta.strip()

        # Remove markdown se houver
        if resposta.startswith('```'):
            resposta = resposta.split('```')[1]
            if resposta.startswith('json'):
                resposta = resposta[4:]
            resposta = resposta.strip()

        # Parse JSON
        dados_json = json.loads(resposta)
        cabecalhos = dados_json.get("cabecalhos", [])
        dados = dados_json.get("dados", [])

        print(f"\n✅ IA gerou:")
        print(f"   📋 {len(cabecalhos)} colunas")
        print(f"   📊 {len(dados)} linhas")

        # Mostra preview
        print("\n👀 Preview dos dados:")
        print(f"   Colunas: {', '.join(cabecalhos)}")
        if dados:
            print(f"   Primeira linha: {dados[0]}")

        # Confirma
        confirma = input("\n✅ Criar planilha com esses dados? (s/n): ").strip().lower()
        if confirma not in ['s', 'sim', 'y', 'yes']:
            print("❌ Cancelado.")
            return

        # Cria Excel
        print(f"\n🔧 Criando {nome_arquivo}...")
        agente.criar_excel(nome_arquivo, dados, cabecalhos)

        # Pergunta se quer abrir
        abrir = input("\n📂 Abrir arquivo agora? (s/n): ").strip().lower()
        if abrir in ['s', 'sim', 'y', 'yes']:
            abrir_arquivo(nome_arquivo)

        print(f"\n🎉 Planilha criada com sucesso!")

    except json.JSONDecodeError as e:
        print(f"\n❌ Erro ao processar resposta da IA")
        print(f"💡 A IA retornou: {resposta[:200]}...")
        print(f"🔧 Erro: {e}")
    except Exception as e:
        print(f"\n❌ Erro: {e}")


def criar_word_com_ia():
    """
    Cria documento Word automaticamente usando IA
    """
    print("\n" + "=" * 60)
    print("🤖 CRIAR WORD COM IA")
    print("=" * 60)

    # Verifica API key
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("\n⚠️  GOOGLE_API_KEY não configurada!")
        print("💡 Configure com: export GOOGLE_API_KEY='sua-chave'")
        api_key = input("\n🔑 Ou cole sua API key aqui: ").strip()
        if not api_key:
            print("❌ Cancelado.")
            return

    agente = AgenteOfficeIA(api_key=api_key)

    # Nome do arquivo
    nome_arquivo = input("\n📁 Nome do arquivo (sem extensão): ").strip()
    if not nome_arquivo:
        nome_arquivo = "documento_ia"
    nome_arquivo = f"{nome_arquivo}.docx"

    # Título
    titulo = input("\n📌 Título do documento: ").strip()
    if not titulo:
        titulo = "Documento Gerado por IA"

    # Tipo de documento
    print("\n📋 Que tipo de documento você quer?")
    print("💡 Exemplos:")
    print("   - Relatório sobre vendas do último trimestre")
    print("   - Artigo sobre inteligência artificial")
    print("   - Proposta comercial para serviço de consultoria")
    print("   - Ata de reunião sobre projeto X")
    print("   - Manual de instruções para usar sistema Y")

    descricao = input("\n➤ Descreva o documento: ").strip()
    if not descricao:
        print("❌ Descrição vazia. Cancelado.")
        return

    # Tamanho
    print("\n📏 Tamanho do documento:")
    print("   1. Curto (1-2 parágrafos)")
    print("   2. Médio (3-5 parágrafos)")
    print("   3. Longo (6+ parágrafos)")

    tamanho_opt = input("\n➤ Opção (padrão: 2): ").strip()
    tamanho_map = {
        '1': 'curto com 1-2 parágrafos',
        '2': 'médio com 3-5 parágrafos',
        '3': 'longo com 6-8 parágrafos'
    }
    tamanho = tamanho_map.get(tamanho_opt, tamanho_map['2'])

    # Gera conteúdo com IA
    print(f"\n🤖 Gerando documento com IA...")
    print("⏳ Aguarde...")

    prompt = f"""Escreva um documento {tamanho} sobre:

"{descricao}"

IMPORTANTE:
- Escreva de forma profissional e bem estruturada
- Divida em parágrafos claros
- Use linguagem formal mas acessível
- Seja objetivo e informativo
- NÃO use markdown, negrito ou itálico
- NÃO use títulos ou subtítulos além do conteúdo
- Apenas texto puro em parágrafos

Retorne APENAS o conteúdo do documento, sem introduções ou explicações."""

    try:
        conteudo = agente.perguntar_ia(prompt)

        # Remove possível markdown
        conteudo = conteudo.replace('**', '').replace('*', '')

        # Divide em parágrafos
        paragrafos = [p.strip() for p in conteudo.split('\n') if p.strip()]

        print(f"\n✅ IA gerou:")
        print(f"   📄 {len(paragrafos)} parágrafos")
        print(f"   📝 {len(conteudo)} caracteres")

        # Mostra preview
        print("\n👀 Preview (primeiros 200 caracteres):")
        print(f"   {conteudo[:200]}...")

        # Confirma
        confirma = input("\n✅ Criar documento com esse conteúdo? (s/n): ").strip().lower()
        if confirma not in ['s', 'sim', 'y', 'yes']:
            print("❌ Cancelado.")
            return

        # Cria Word
        print(f"\n🔧 Criando {nome_arquivo}...")
        agente.criar_word(nome_arquivo, titulo, paragrafos)

        # Pergunta se quer abrir
        abrir = input("\n📂 Abrir arquivo agora? (s/n): ").strip().lower()
        if abrir in ['s', 'sim', 'y', 'yes']:
            abrir_arquivo(nome_arquivo)

        print(f"\n🎉 Documento criado com sucesso!")

    except Exception as e:
        print(f"\n❌ Erro: {e}")


def analisar_excel_e_gerar_relatorio():
    """
    Lê um Excel existente, analisa com IA e gera relatório Word
    """
    print("\n" + "=" * 60)
    print("📊➡️📄 ANALISAR EXCEL E GERAR RELATÓRIO")
    print("=" * 60)

    # Verifica API key
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("\n⚠️  GOOGLE_API_KEY não configurada!")
        api_key = input("\n🔑 Cole sua API key aqui: ").strip()
        if not api_key:
            print("❌ Cancelado.")
            return

    agente = AgenteOfficeIA(api_key=api_key)

    # Lista arquivos Excel
    arquivos_excel = [f for f in os.listdir('.') if f.endswith('.xlsx')]

    if not arquivos_excel:
        print("\n⚠️  Nenhum arquivo Excel encontrado no diretório atual")
        return

    print("\n📋 Arquivos Excel disponíveis:")
    for i, arquivo in enumerate(arquivos_excel, 1):
        print(f"   {i}. {arquivo}")

    # Escolhe arquivo
    escolha = input("\n➤ Escolha o arquivo (número ou nome): ").strip()

    if escolha.isdigit():
        idx = int(escolha) - 1
        if 0 <= idx < len(arquivos_excel):
            arquivo_excel = arquivos_excel[idx]
        else:
            print("❌ Número inválido!")
            return
    else:
        arquivo_excel = escolha
        if not arquivo_excel.endswith('.xlsx'):
            arquivo_excel += '.xlsx'

    if not os.path.exists(arquivo_excel):
        print(f"❌ Arquivo '{arquivo_excel}' não encontrado!")
        return

    # Lê Excel
    print(f"\n📖 Lendo {arquivo_excel}...")
    dados = agente.ler_excel(arquivo_excel)

    print(f"✅ {len(dados)} linhas lidas")

    # Analisa com IA
    print("\n🤖 Analisando dados com IA...")
    print("⏳ Aguarde...")

    # Pega amostra dos dados
    amostra = dados[:min(20, len(dados))]

    prompt = f"""Analise os dados desta planilha Excel e crie um relatório executivo completo.

Dados (primeiras {len(amostra)} linhas):
{json.dumps(amostra, ensure_ascii=False, indent=2)}

Crie um relatório com:
1. RESUMO EXECUTIVO: visão geral dos dados
2. ANÁLISE DETALHADA: insights principais e padrões identificados
3. ESTATÍSTICAS: números e métricas importantes
4. CONCLUSÕES: principais descobertas
5. RECOMENDAÇÕES: sugestões baseadas nos dados

Escreva de forma profissional, objetiva e estruturada.
Use parágrafos separados para cada seção.
NÃO use markdown ou formatação especial."""

    try:
        analise = agente.perguntar_ia(prompt)

        # Remove markdown se houver
        analise = analise.replace('**', '').replace('*', '')

        print(f"\n✅ Análise gerada ({len(analise)} caracteres)")

        # Nome do relatório
        nome_relatorio = arquivo_excel.replace('.xlsx', '_relatorio.docx')

        # Cria Word
        print(f"\n🔧 Criando relatório {nome_relatorio}...")

        titulo = f"Relatório: {arquivo_excel}"
        paragrafos = [
            "Este relatório foi gerado automaticamente por IA a partir da análise dos dados da planilha.",
            "",
            analise
        ]

        agente.criar_word(nome_relatorio, titulo, paragrafos)

        # Pergunta se quer abrir
        abrir = input("\n📂 Abrir relatório agora? (s/n): ").strip().lower()
        if abrir in ['s', 'sim', 'y', 'yes']:
            abrir_arquivo(nome_relatorio)

        print(f"\n🎉 Relatório criado com sucesso!")
        print(f"   📊 Fonte: {arquivo_excel}")
        print(f"   📄 Relatório: {nome_relatorio}")

    except Exception as e:
        print(f"\n❌ Erro: {e}")


def menu_principal():
    """
    Menu principal para escolher o que fazer com IA
    """
    while True:
        print("\n" + "=" * 60)
        print("🤖 CRIAR ARQUIVOS COM IA (GEMINI)")
        print("=" * 60)
        print("\n📋 Escolha uma opção:")
        print("   1. Criar Excel com IA")
        print("   2. Criar Word com IA")
        print("   3. Analisar Excel e gerar relatório")
        print("   4. Sair")

        opcao = input("\n➤ Opção: ").strip()

        if opcao == '1':
            criar_excel_com_ia()
        elif opcao == '2':
            criar_word_com_ia()
        elif opcao == '3':
            analisar_excel_e_gerar_relatorio()
        elif opcao in ['4', 'sair', 'exit', 'q']:
            print("\n👋 Até mais!")
            break
        else:
            print("\n❌ Opção inválida! Tente novamente.")


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("⚙️  CONFIGURAÇÃO")
    print("=" * 60)

    api_key = os.environ.get("GOOGLE_API_KEY")
    if api_key:
        print("✅ GOOGLE_API_KEY detectada")
    else:
        print("⚠️  GOOGLE_API_KEY não configurada")
        print("\n💡 Para configurar:")
        print("   export GOOGLE_API_KEY='sua-chave-aqui'")
        print("\n💡 Ou você pode colar quando solicitado")
        print("\n🔑 Obtenha sua chave em: https://makersuite.google.com/app/apikey")

    menu_principal()