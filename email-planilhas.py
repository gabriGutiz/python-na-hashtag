import pandas as pd
tabela_vendas = pd.read_excel('vendas.xlsx')
# trazer tabela, usar r e aspas duplas para caminho do computador

tabela_faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
# pegar apenas as colunas especificadas, groupby = agrupar as infromações com base na coluna desejada e somar a outra
tabela_faturamento = tabela_faturamento.sort_values(by='Valor Final', ascending=False)
# sort_values = ordenar valores com base na coluna 'Valor Final', ascending = False -- decrescente

tabela_quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
# pegar apenas o desejado e agrupar
tabela_quantidade = tabela_quantidade.sort_values(by='Quantidade', ascending=False)
# ordenar valores

ticket_medio = (tabela_faturamento['Valor Final'] / tabela_quantidade['Quantidade']).to_frame()
# ticket medio = faturamento / quantidade vendida
# operação usando as tabelas desejadas e transformando em tabela com to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
# a coluna ticket medio sai com nome '0', renomear para 'Ticket Medio'
ticket_medio = ticket_medio.sort_values(by='Ticket Medio', ascending=False)
# ordenar valores


def enviar_email(nome_da_loja, tabela):
    import smtplib
    import email.message
    server = smtplib.SMTP('smtp.gmail.com:587')
    corpo_email = f"""
  <p> Prezados, segue reltório de vendas. </p>
  <p> Contate-me para qualquer dúvida. </p>
  {tabela.to_html}
  """  # EDITAR corpo do email -- .to_html() deixa a tabela bonitinha

    msg = email.message.Message()
    msg['Subject'] = f"Relatório de Vendas - {nome_da_loja}"  # EDITAR assunto do email

    # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
    # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
    # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha

    msg['From'] = 'gabri.cesar243@gmail.com'  # EDITAR remetente
    msg['To'] = 'gabri.faculs@gmail.com'  # EDITAR destinatário
    password = "gabriel1008@"  # EDITAR senha do email
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')
# definir enviar_email(informações que mudam) como toda essa operação
# usar f antes da mensagem que vai utilizar variável -- {tabela.to_html()}
# usar """ mensagem """ para usar mais de uma linha


tabela_completa = tabela_faturamento.join(tabela_quantidade).join(ticket_medio)
# join junta as tabelas com oq elas não tem em comum usando como referencia oq tem em comum
tabela_completa = tabela_completa.sort_values(by='Ticket Medio', ascending=False)
# tabela_completa = tabela_completa.to_excel('Venda.xlsx')
'enviar_email("Diretoria", tabela_completa)'
print(tabela_completa)

lista_lojas = tabela_vendas["ID Loja"].unique()
# pegar informações de tabela_vendas com unique() = sem q se repitem pegar uma unica vez

'''for loja in lista_lojas:
    #tabela_loja = tabela_completa.loc[tabela_completa["ID Loja"] == loja, ["Valor Final", "Quantidade",  "Ticket Medio"]]
    #loc = localizar [linha, coluna]
    #enviar_email(loja, tabela_loja)

    #OU FAZER ASSIM:
    tabela_loja = tabela_vendas.loc[tabela_vendas["ID Loja"] == loja, ["ID Loja", "Quantidade", "Valor Final"]]
    tabela_loja = tabela_loja.groupby("ID Loja").sum()
    tabela_loja["Ticket Medio"] = tabela_loja["Valor Final"] / tabela_loja["Quantidade"]
    enviar_email(loja, tabela_loja)'''
