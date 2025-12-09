# Automação de Certificados com Google Forms, Sheets, Slides e Apps Script

Este projeto oferece uma solução completa para automatizar a geração, envio e controle de certificados emitidos a partir das respostas enviadas por um Google Forms.  
A automação integra diversos serviços do Google Workspace, reduzindo trabalho manual, evitando erros e garantindo rastreabilidade total.

---

## Sumário
1. Introdução  
2. Arquitetura do Sistema  
3. Estrutura da Planilha  
4. Template do Certificado  
5. Pasta de Certificados no Drive  
6. Script do Apps Script  
7. Gatilhos Necessários  
8. Fluxo Completo  
9. Requisitos  
10. Instalação  
11. Licença  

---

# 1. Introdução

Este projeto automatiza:

- Processamento de respostas do formulário  
- Cálculo de pontuação e aprovação  
- Geração do certificado em PDF  
- Personalização dos dados no Slides  
- Armazenamento em uma pasta específica no Drive  
- Envio automático de e-mail ao participante  
- Registro de token, link e status na planilha  

---

# 2. Arquitetura do Sistema

Fluxo utilizado:

```
Google Forms → Google Sheets → Apps Script → Google Slides → Google Drive → Gmail
```

O processo inicia quando o usuário envia um formulário.  
A planilha recebe os dados e o Apps Script dispara automaticamente a geração e entrega do certificado.

---

# 3. Estrutura da Planilha

A aba obrigatória deve se chamar:

```
Form_Responses
```

E conter as seguintes colunas (1 a 24):

| Coluna | Cabeçalho |
|--------|-----------|
| A | Carimbo de data/hora |
| B | Endereço de e-mail |
| C | Pontuação |
| D | Nome completo |
| E | CNPJ |
| F | Analista |
| G–P | Perguntas do quiz |
| Q | Curso |
| R | Data_Inicio_Curso |
| S | Data_Termino_Curso |
| T | Status_Aprovacao |
| U | Certificado_Gerado |
| V | Link_Certificado |
| W | Token |
| X | Validade_Certificado |

### Fórmula de aprovação (coluna T)

```excel
=ArrayFormula(
  SE(
    C2:C = "";
    ;
    SE( VALOR(REGEXEXTRACT(TO_TEXT(C2:C); "^\d+")) >= 7;
        "Aprovado";
        "Reprovado"
    )
  )
)
```

### Fórmula de validade (coluna X)

```excel
=ArrayFormula(
  SE(
    R2:R = "";
    ;
    R2:R + 365
  )
)
```

---

# 4. Template do Certificado

O template no Google Slides deve conter os placeholders:

```
<<NOME>>
<<DATA>>
<<INSTRUTOR>>
```

Esses textos serão substituídos dinamicamente pelo Apps Script.

---

# 5. Pasta de Certificados no Google Drive

Crie uma pasta e copie o ID após `/folders/`.

A configuração deve ser inserida no script no campo:

```
FOLDER_ID
```

---

# 6. Script do Apps Script

O script completo está no arquivo:

```
code.gs
```

Ele contém:

- Função principal `onFormSubmit`
- Geração de token único
- Geração do certificado em PDF usando Slides
- Atualização da planilha
- Envio de e-mail ao participante

---

# 7. Gatilhos Necessários

No Apps Script:

1. Vá para **Gatilhos**  
2. Adicione um novo gatilho com as configurações:

```
Função: onFormSubmit
Origem do evento: Da planilha
Tipo de evento: Ao enviar formulário
```

Permissões necessárias: Drive, Slides, Gmail, Sheets.

---

# 8. Fluxo Completo

1. Usuário envia o formulário  
2. Resposta é gravada na planilha  
3. Fórmulas calculam status e validade  
4. Script verifica se o participante está aprovado  
5. Certificado é criado como PDF  
6. PDF é salvo no Drive  
7. Planilha é atualizada com token, link e status  
8. E-mail é enviado automaticamente  

---

# 9. Requisitos

- Conta Google com permissões de Drive, Slides, Gmail e Apps Script  
- Template configurado  
- IDs configurados corretamente no script  
- Planilha seguindo o layout oficial  

---

# 10. Instalação

1. Crie a planilha e insira os cabeçalhos  
2. Crie o Forms vinculado à planilha  
3. Insira as fórmulas  
4. Configure o template no Slides  
5. Crie a pasta de certificados no Drive  
6. Cole o script no Apps Script  
7. Configure os gatilhos  
8. Teste o fluxo  

---

# 11. Licença

Este projeto está licenciado sob os termos da licença incluída no arquivo `LICENSE.txt`.

