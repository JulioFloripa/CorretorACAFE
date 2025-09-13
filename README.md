# 📊 Gerador de Boletins - Simulado ACAFE

Sistema automatizado para correção de simulados ACAFE e geração de boletins individuais em PDF.

## 🚀 Funcionalidades

- ✅ Correção automática de respostas baseada no gabarito
- 📊 Geração de ranking da turma
- 📈 Cálculo de médias por disciplina
- 📄 Boletins individuais em PDF com gráficos
- 📦 Download de todos os boletins em arquivo ZIP
- 🔍 Validação completa dos dados de entrada
- 📱 Interface responsiva e intuitiva

## 📋 Formato do Arquivo Excel

O arquivo deve conter **duas abas obrigatórias**:

### Aba "RESPOSTAS"
| ID | Nome | Q1 | Q2 | Q3 | ... |
|----|------|----|----|----|----|
| 1 | João Silva | A | B | C | ... |
| 2 | Maria Santos | B | A | D | ... |

**Colunas obrigatórias:**
- `ID`: Identificador único do aluno (número)
- `Nome`: Nome completo do aluno
- `Q1`, `Q2`, `Q3`, etc.: Respostas do aluno (A, B, C, D ou E)

### Aba "GABARITO"
| Questão | Resposta | Disciplina |
|---------|----------|------------|
| 1 | A | Matemática |
| 2 | B | Português |
| 3 | C | História |

**Colunas obrigatórias:**
- `Questão`: Número da questão
- `Resposta`: Resposta correta (A, B, C, D ou E)
- `Disciplina`: Nome da disciplina

## 🛠️ Instalação e Execução

### Pré-requisitos
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Instalação Local

1. **Clone o repositório:**
```bash
git clone https://github.com/JulioFloripa/CorretorACAFE.git
cd CorretorACAFE
```

2. **Instale as dependências:**
```bash
pip install -r requirements.txt
```

3. **Execute a aplicação:**
```bash
streamlit run app.py
```

4. **Acesse no navegador:**
```
http://localhost:8501
```

### Deploy no Render

1. **Fork este repositório**
2. **Conecte sua conta do Render ao GitHub**
3. **Crie um novo Web Service no Render:**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`

## 📊 Relatórios Gerados

Cada boletim individual contém:

### 📈 Gráficos
- **Radar**: Comparação do desempenho por disciplina
- **Barras**: Notas individuais vs média da turma
- **Distribuição**: Posição na distribuição geral das notas
- **Ranking**: Posição no ranking da turma

### 📋 Tabela de Resultados
- Acertos por disciplina
- Percentual de acertos
- Comparação com a média da turma
- Diferença em relação à média

### 📊 Informações Gerais
- Posição no ranking
- Nota individual
- Média da turma
- Diferença em relação à média

## 🔧 Melhorias Implementadas

### ✅ Validação de Dados
- Verificação da existência das abas obrigatórias
- Validação das colunas necessárias
- Detecção de questões duplicadas no gabarito
- Verificação de dados nulos ou incompletos

### 🎨 Interface Melhorada
- Barra de progresso durante o processamento
- Preview dos dados antes do processamento
- Sidebar com instruções e estatísticas
- Mensagens de erro mais claras e específicas

### 📊 Gráficos Aprimorados
- Melhor qualidade visual (DPI 150)
- Cores e estilos mais profissionais
- Grid e legendas melhoradas
- Tratamento de erros na geração

### 🔒 Robustez
- Tratamento completo de exceções
- Logs detalhados de erros
- Validação de tipos de dados
- Configuração adequada do matplotlib

### 📱 Experiência do Usuário
- Estatísticas em tempo real
- Feedback visual do progresso
- Instruções claras na sidebar
- Mensagens de sucesso e erro

## 🐛 Solução de Problemas

### Erro: "Aba não encontrada"
- Verifique se o arquivo Excel contém as abas "RESPOSTAS" e "GABARITO"
- Certifique-se de que os nomes estão exatamente como especificado

### Erro: "Coluna não encontrada"
- Verifique se todas as colunas obrigatórias estão presentes
- Certifique-se de que não há espaços extras nos nomes das colunas

### Erro na geração de PDF
- Verifique se há caracteres especiais nos nomes dos alunos
- Certifique-se de que há dados suficientes para gerar os gráficos

### Aplicação lenta
- Para arquivos grandes (>500 alunos), o processamento pode demorar alguns minutos
- Verifique a conexão de internet se estiver usando o deploy online

## 📞 Suporte

Para reportar bugs ou solicitar melhorias:
1. Abra uma issue no GitHub
2. Descreva o problema detalhadamente
3. Inclua o arquivo Excel (sem dados pessoais) se possível

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

## 🤝 Contribuições

Contribuições são bem-vindas! Por favor:
1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

---

**Desenvolvido para o Colégio Fleming** 🎓

