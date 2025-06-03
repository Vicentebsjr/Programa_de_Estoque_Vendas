#preparação para criar executável
"""
import os
import pandas as pd

def inicializar_planilhas():
    if not os.path.exists('clientes.xlsx'):
        df_clientes = pd.DataFrame(columns=['nome', 'cpf', 'telefone', 'email', 'divida'])
        df_clientes.to_excel('clientes.xlsx', index=False)
        print('Planilha de clientes criada.')

    if not os.path.exists('produtos.xlsx'):
        df_produtos = pd.DataFrame(columns=['codigo', 'nome', 'quantidade', 'preco_custo', 'preco_venda'])
        df_produtos.to_excel('produtos.xlsx', index=False)
        print('Planilha de produtos criada.')

    if not os.path.exists('vendas.xlsx'):
        df_vendas = pd.DataFrame(columns=['data', 'codigo_produto', 'quantidade', 'preco_venda', 'tipo_venda'])
        df_vendas.to_excel('vendas.xlsx', index=False)
        print('Planilha de vendas criada.')

# Chame a função de inicialização ao iniciar o programa
inicializar_planilhas()
"""


# Início do código em si


import pandas as pd
import os
import random
from datetime import datetime

# Função para carregar dados de um arquivo Excel
def carregar_dados(nome_arquivo, sheet_name):
    try:
        # Carregar os dados do arquivo Excel
        df = pd.read_excel(nome_arquivo, sheet_name=sheet_name)

        # Converter o campo 'data' para datetime, se existir
        if 'data' in df.columns:
            df['data'] = pd.to_datetime(df['data'])

        # Converter os dados para um dicionário de registros
        return df.to_dict(orient='records')

    except FileNotFoundError:
        return []

# Função para salvar dados em um arquivo Excel
def salvar_dados(nome_arquivo, sheet_name, dados):
    df = pd.DataFrame(dados)
    if os.path.exists(nome_arquivo):
        with pd.ExcelWriter(nome_arquivo, engine='openpyxl', mode='a') as writer:
            if sheet_name in writer.book.sheetnames:
                writer.book.remove(writer.book[sheet_name])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Função para formatar valores monetários
def formatar_valor(valor):
    valor = valor.replace(',', '.')
    return round(float(valor), 2)

# Gerar código de cliente
def gerar_codigo_cliente():
    codigo = 'c' + str(random.randint(1000, 9999))
    return codigo

# Cadastrar cliente
def cadastrar_cliente():
    nome = input('Digite o nome do cliente: ').strip()
    telefone = input('Digite o telefone do cliente: ').strip()
    endereco = input('Digite o endereço do cliente: ').strip()
    codigo = gerar_codigo_cliente()

    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    clientes.append({
        'codigo': codigo,
        'nome': nome,
        'telefone': telefone,
        'endereco': endereco,
        'divida': 0.0
    })
    salvar_dados('clientes.xlsx', 'Clientes', clientes)
    print(f'Cliente {nome} cadastrado com sucesso! Código: {codigo}')

# Editar cliente
def editar_cliente():
    codigo = input('Digite o código do cliente a ser editado: ').strip()
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    cliente_encontrado = False

    for cliente in clientes:
        if cliente['codigo'] == codigo:
            cliente_encontrado = True
            nome = input(f'Nome ({cliente["nome"]}): ').strip() or cliente['nome']
            telefone = input(f'Telefone ({cliente["telefone"]}): ').strip() or cliente['telefone']
            endereco = input(f'Endereço ({cliente["endereco"]}): ').strip() or cliente['endereco']

            cliente['nome'] = nome
            cliente['telefone'] = telefone
            cliente['endereco'] = endereco

            salvar_dados('clientes.xlsx', 'Clientes', clientes)
            print(f'Cliente {codigo} atualizado com sucesso!')
            return

    if not cliente_encontrado:
        print('Cliente não encontrado.')

# Remover cliente
def remover_cliente():
    codigo = input('Digite o código do cliente a ser removido: ').strip()
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    cliente_encontrado = False

    for cliente in clientes:
        if cliente['codigo'] == codigo:
            cliente_encontrado = True
            clientes.remove(cliente)
            salvar_dados('clientes.xlsx', 'Clientes', clientes)
            print(f'Cliente {codigo} removido com sucesso!')
            return

    if not cliente_encontrado:
        print('Cliente não encontrado.')

# Acessar histórico do cliente
def acessar_historico_cliente():
    codigo = input('Digite o código do cliente para acessar o histórico de vendas: ').strip()
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    cliente_encontrado = False

    for cliente in clientes:
        if cliente['codigo'] == codigo:
            cliente_encontrado = True
            print(f'Histórico de vendas para o cliente {cliente["nome"]}:')
            vendas_cliente = carregar_dados('vendas.xlsx', 'Vendas')
            vendas_cliente = [venda for venda in vendas_cliente if venda.get('codigo_cliente') == codigo]

            if vendas_cliente:
                for venda in vendas_cliente:
                    if venda['tipo_venda'] == 'À vista':
                        print(f'Produto: {venda["nome_produto"]}, Quantidade: {venda["quantidade"]}, Preço de Venda: {venda["preco_venda"]:.2f}, Data: {venda["data"]}')
                    else:
                        print(f'Produto: {venda["nome_produto"]}, Quantidade: {venda["quantidade"]}, Preço de Venda: {venda["preco_venda"]:.2f}, Tipo: {venda["tipo_venda"]}, Data: {venda["data"]}, Valor da dívida: R$ {venda["divida"]:.2f}')
            else:
                print('Nenhuma venda registrada para este cliente.')
            return

    if not cliente_encontrado:
        print('Cliente não encontrado.')

# Abater dívida de cliente
def abater_divida_cliente():
    codigo = input('Digite o código do cliente para abater a dívida: ').strip()
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    cliente_encontrado = False

    for cliente in clientes:
        if cliente['codigo'] == codigo:
            cliente_encontrado = True
            divida_atual = cliente['divida']
            valor_abatido = formatar_valor(input(f'Digite o valor a ser abatido da dívida atual de R${divida_atual:.2f}: '))

            if valor_abatido <= divida_atual:
                cliente['divida'] -= valor_abatido
                salvar_dados('clientes.xlsx', 'Clientes', clientes)
                print(f'Dívida atualizada. Dívida restante do cliente {cliente["nome"]}: R$ {cliente["divida"]:.2f}')
            else:
                print('O valor abatido não pode ser maior que a dívida atual.')
            return

    if not cliente_encontrado:
        print('Cliente não encontrado.')

# Consultar dívida de cliente
def consultar_divida_cliente():
    codigo = input('Digite o código do cliente para consultar a dívida: ').strip()
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    cliente_encontrado = False

    for cliente in clientes:
        if cliente['codigo'] == codigo:
            cliente_encontrado = True
            print(f'Dívida do cliente {cliente["nome"]}: R$ {cliente["divida"]:.2f}')
            return

    if not cliente_encontrado:
        print('Cliente não encontrado.')

# Consultar lista de clientes
def consultar_lista_clientes():
    clientes = carregar_dados('clientes.xlsx', 'Clientes')

    if clientes:
        print('Lista de Clientes:')
        for cliente in clientes:
            print(f'Código: {cliente["codigo"]}, Nome: {cliente["nome"]}, Telefone: {cliente["telefone"]}, Endereço: {cliente["endereco"]}, Dívida: R$ {cliente["divida"]:.2f}')
    else:
        print('Nenhum cliente cadastrado.')

# Vender produto
def vender_produto():
    codigo_produto = input('Digite o código do produto: ').strip()
    quantidade = int(input('Digite a quantidade vendida: '))
    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    cliente_associado = None
    desconto = False
    valor_final = 0.0
    desconto_r = 0.0
    desconto_percentual = 0.0

    for produto in produtos:
        if produto['codigo'] == codigo_produto:
            if produto['quantidade'] >= quantidade:
                valor_final = produto["preco_venda"] * quantidade
                if input("Houve desconto na venda? (S/N): ").upper() == "S":
                    desconto = True
                    valor_final = formatar_valor(input("Digite o valor final de venda: "))
                    desconto_r = produto["preco_venda"] * quantidade - valor_final
                    desconto_percentual = (desconto_r / (produto["preco_venda"] * quantidade)) * 100

                produto['quantidade'] -= quantidade
                salvar_dados('produtos.xlsx', 'Produtos', produtos)

                # Verificar se o cliente é cadastrado e associar a venda
                clientes = carregar_dados('clientes.xlsx', 'Clientes')
                codigo_cliente = input('Digite o código do cliente (ou deixe em branco para venda sem cliente): ').strip()

                if codigo_cliente:
                    cliente_associado = next((cliente for cliente in clientes if cliente['codigo'] == codigo_cliente), None)
                    if cliente_associado:
                        cliente_associado['divida'] += valor_final

                venda = {
                    'nome_produto': produto['nome'],
                    'quantidade': quantidade,
                    'preco_venda': valor_final,
                    'data': datetime.now(),
                    'codigo_cliente': codigo_cliente if cliente_associado else None,
                    'divida': valor_final if cliente_associado else None,
                    'tipo_venda': 'À vista' if not cliente_associado else 'A prazo'
                }
                vendas = carregar_dados('vendas.xlsx', 'Vendas')
                vendas.append(venda)
                salvar_dados('vendas.xlsx', 'Vendas', vendas)
                print(f'Produto {produto["nome"]} vendido com sucesso!')
            else:
                print('Quantidade em estoque insuficiente.')
            return
    print('Produto não encontrado.')

# Cadastrar produto
def cadastrar_produto():
    nome = input('Digite o nome do produto: ').strip()
    quantidade = int(input('Digite a quantidade do produto: '))
    preco_custo = formatar_valor(input('Digite o preço de custo do produto: '))
    preco_venda = formatar_valor(input('Digite o preço de venda do produto: '))
    codigo = input('Digite o código do produto: ').strip()

    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    produtos.append({
        'codigo': codigo,
        'nome': nome,
        'quantidade': quantidade,
        'preco_custo': preco_custo,
        'preco_venda': preco_venda
    })
    salvar_dados('produtos.xlsx', 'Produtos', produtos)
    print(f'Produto {nome} cadastrado com sucesso!')

# Editar produto
def editar_produto():
    codigo = input('Digite o código do produto a ser editado: ').strip()
    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    produto_encontrado = False

    for produto in produtos:
        if produto['codigo'] == codigo:
            produto_encontrado = True
            nome = input(f'Nome ({produto["nome"]}): ').strip() or produto['nome']
            quantidade = input(f'Quantidade ({produto["quantidade"]}): ').strip() or produto['quantidade']
            preco_custo = input(f'Preço de Custo ({produto["preco_custo"]:.2f}): ').strip() or produto['preco_custo']
            preco_venda = input(f'Preço de Venda ({produto["preco_venda"]:.2f}): ').strip() or produto['preco_venda']

            produto['nome'] = nome
            produto['quantidade'] = int(quantidade)
            produto['preco_custo'] = formatar_valor(preco_custo)
            produto['preco_venda'] = formatar_valor(preco_venda)

            salvar_dados('produtos.xlsx', 'Produtos', produtos)
            print(f'Produto {codigo} atualizado com sucesso!')
            return

    if not produto_encontrado:
        print('Produto não encontrado.')

# Remover produto
def remover_produto():
    codigo = input('Digite o código do produto a ser removido: ').strip()
    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    produto_encontrado = False

    for produto in produtos:
        if produto['codigo'] == codigo:
            produto_encontrado = True
            produtos.remove(produto)
            salvar_dados('produtos.xlsx', 'Produtos', produtos)
            print(f'Produto {codigo} removido com sucesso!')
            return

    if not produto_encontrado:
        print('Produto não encontrado.')

# Consultar lista de produtos
def consultar_lista_produtos():
    produtos = carregar_dados('produtos.xlsx', 'Produtos')

    if produtos:
        print('Lista de Produtos:')
        for produto in produtos:
            print(f'Código: {produto["codigo"]}, Nome: {produto["nome"]}, Quantidade: {produto["quantidade"]}, Preço de Custo: R$ {produto["preco_custo"]:.2f}, Preço de Venda: R$ {produto["preco_venda"]:.2f}')
    else:
        print('Nenhum produto cadastrado.')

# Consultar produto específico
def consultar_produto():
    codigo = input('Digite o código do produto para consulta: ').strip()
    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    produto_encontrado = False

    for produto in produtos:
        if produto['codigo'] == codigo:
            produto_encontrado = True
            print(f'Código: {produto["codigo"]}, Nome: {produto["nome"]}, Quantidade: {produto["quantidade"]}, Preço de Custo: R$ {produto["preco_custo"]:.2f}, Preço de Venda: R$ {produto["preco_venda"]:.2f}')
            return

    if not produto_encontrado:
        print('Produto não encontrado.')

# Relatórios
def mostrar_tudo_vendido_no_mes():
    mes = int(input('Digite o mês (1-12): '))
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_mes = [venda for venda in vendas if venda['data'].month == mes]

    if vendas_mes:
        print(f'Vendas do mês {mes}:')
        for venda in vendas_mes:
            print(venda)
    else:
        print('Nenhuma venda registrada neste mês.')

def historico_mensal_vendas():
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_por_mes = {}

    for venda in vendas:
        mes = venda['data'].strftime('%Y-%m')
        if mes not in vendas_por_mes:
            vendas_por_mes[mes] = []
        vendas_por_mes[mes].append(venda)

    for mes, vendas_mes in vendas_por_mes.items():
        print(f'Vendas do mês {mes}:')
        for venda in vendas_mes:
            print(venda)

def vendas_diarias():
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_por_dia = {}

    for venda in vendas:
        dia = venda['data'].strftime('%Y-%m-%d')
        if dia not in vendas_por_dia:
            vendas_por_dia[dia] = []
        vendas_por_dia[dia].append(venda)

    for dia, vendas_dia in vendas_por_dia.items():
        print(f'Vendas do dia {dia}:')
        for venda in vendas_dia:
            print(venda)

def vendas_prazo_a_receber():
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_a_prazo = [venda for venda in vendas if venda['tipo_venda'] == 'A prazo']

    if vendas_a_prazo:
        print('Vendas a prazo a receber:')
        for venda in vendas_a_prazo:
            print(venda)
    else:
        print('Nenhuma venda a prazo registrada.')

def saldo_total_mes():
    mes = int(input('Digite o mês (1-12): '))
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_mes = [venda for venda in vendas if venda['data'].month == mes]
    saldo_total = sum(venda['preco_venda'] for venda in vendas_mes)
    print(f'Saldo total do mês {mes}: R$ {saldo_total:.2f}')

def saldo_total_dia():
    dia = input('Digite a data (AAAA-MM-DD): ')
    vendas = carregar_dados('vendas.xlsx', 'Vendas')
    vendas_dia = [venda for venda in vendas if venda['data'].strftime('%Y-%m-%d') == dia]
    saldo_total = sum(venda['preco_venda'] for venda in vendas_dia)
    print(f'Saldo total do dia {dia}: R$ {saldo_total:.2f}')

def clientes_com_dividas():
    clientes = carregar_dados('clientes.xlsx', 'Clientes')
    clientes_dividas = [cliente for cliente in clientes if cliente['divida'] > 0]

    if clientes_dividas:
        print('Clientes com dívidas:')
        for cliente in clientes_dividas:
            print(f'Nome: {cliente["nome"]}, Dívida: R$ {cliente["divida"]:.2f}')
    else:
        print('Nenhum cliente com dívida registrada.')

def produtos_com_estoque_baixo():
    produtos = carregar_dados('produtos.xlsx', 'Produtos')
    quantidade_limite = int(input('Digite a quantidade limite para considerar estoque baixo: '))
    produtos_estoque_baixo = [produto for produto in produtos if produto['quantidade'] <= quantidade_limite]

    if produtos_estoque_baixo:
        print('Produtos com estoque baixo:')
        for produto in produtos_estoque_baixo:
            print(f'Nome: {produto["nome"]}, Quantidade: {produto["quantidade"]}')
    else:
        print('Nenhum produto com estoque baixo registrado.')

# Menu Principal
def menu_principal():
    while True:
        print('\n==== Menu Principal ====')
        print('1. Clientes')
        print('2. Produtos')
        print('3. Vendas')
        print('4. Relatórios')
        print('0. Sair')
        opcao = input('Escolha uma opção: ').strip()

        if opcao == '1':
            menu_clientes()
        elif opcao == '2':
            menu_produtos()
        elif opcao == '3':
            menu_vendas()
        elif opcao == '4':
            menu_relatorios()
        elif opcao == '0':
            print('Saindo do sistema...')
            break
        else:
            print('Opção inválida. Tente novamente.')

# Menu Clientes
def menu_clientes():
    while True:
        print('\n==== Menu Clientes ====')
        print('1. Cadastrar Cliente')
        print('2. Editar Cliente')
        print('3. Remover Cliente')
        print('4. Acessar Histórico do Cliente')
        print('5. Abater Dívida de Cliente')
        print('6. Consultar Dívida de Cliente')
        print('7. Consultar Lista de Clientes')
        print('0. Voltar ao Menu Principal')
        opcao = input('Escolha uma opção: ').strip()

        if opcao == '1':
            cadastrar_cliente()
        elif opcao == '2':
            editar_cliente()
        elif opcao == '3':
            remover_cliente()
        elif opcao == '4':
            acessar_historico_cliente()
        elif opcao == '5':
            abater_divida_cliente()
        elif opcao == '6':
            consultar_divida_cliente()
        elif opcao == '7':
            consultar_lista_clientes()
        elif opcao == '0':
            break
        else:
            print('Opção inválida. Tente novamente.')

# Menu Produtos
def menu_produtos():
    while True:
        print('\n==== Menu Produtos ====')
        print('1. Cadastrar Produto')
        print('2. Editar Produto')
        print('3. Remover Produto')
        print('4. Consultar Lista de Produtos')
        print('5. Consultar Produto Específico')
        print('0. Voltar ao Menu Principal')
        opcao = input('Escolha uma opção: ').strip()

        if opcao == '1':
            cadastrar_produto()
        elif opcao == '2':
            editar_produto()
        elif opcao == '3':
            remover_produto()
        elif opcao == '4':
            consultar_lista_produtos()
        elif opcao == '5':
            consultar_produto()
        elif opcao == '0':
            break
        else:
            print('Opção inválida. Tente novamente.')

# Menu Vendas
def menu_vendas():
    while True:
        print('\n==== Menu Vendas ====')
        print('1. Vender Produto')
        print('0. Voltar ao Menu Principal')
        opcao = input('Escolha uma opção: ').strip()

        if opcao == '1':
            vender_produto()
        elif opcao == '0':
            break
        else:
            print('Opção inválida. Tente novamente.')

# Menu Relatórios
def menu_relatorios():
    while True:
        print('\n==== Menu Relatórios ====')
        print('1. Mostrar Tudo Vendido no Mês')
        print('2. Histórico Mensal de Vendas')
        print('3. Vendas Diárias')
        print('4. Vendas a Prazo a Receber')
        print('5. Saldo Total do Mês')
        print('6. Saldo Total do Dia')
        print('7. Clientes com Dívidas')
        print('8. Produtos com Estoque Baixo')
        print('0. Voltar ao Menu Principal')
        opcao = input('Escolha uma opção: ').strip()

        if opcao == '1':
            mostrar_tudo_vendido_no_mes()
        elif opcao == '2':
            historico_mensal_vendas()
        elif opcao == '3':
            vendas_diarias()
        elif opcao == '4':
            vendas_prazo_a_receber()
        elif opcao == '5':
            saldo_total_mes()
        elif opcao == '6':
            saldo_total_dia()
        elif opcao == '7':
            clientes_com_dividas()
        elif opcao == '8':
            produtos_com_estoque_baixo()
        elif opcao == '0':
            break
        else:
            print('Opção inválida. Tente novamente.')

# Executar o menu principal
menu_principal()

