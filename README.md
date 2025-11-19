# 🔄 Gerador de Fluxograma Interativo

Aplicação web Flask que gera fluxogramas interativos a partir de arquivos Excel, permitindo arrastar e reorganizar elementos com ajuste automático das conexões usando Drawflow.

## 🚀 Funcionalidades

- ✨ **Upload de Excel**: Importa dados estruturados de planilhas .xlsx
- 🎨 **Renderização Interativa**: Utiliza Drawflow para criar fluxogramas editáveis
- 🖱️ **Arrastar e Soltar**: Mova nós livremente com reconexão automática das setas
- 🔍 **Zoom Controles**: Amplie, reduza ou redefina a visualização
- 💾 **Exportação**: Salve o fluxograma em múltiplos formatos (PNG, PDF, SVG, JSON)
- 📱 **Interface Responsiva**: Design moderno e adaptável para desktop e mobile

## 📋 Formato do Excel

O arquivo Excel deve conter as seguintes colunas obrigatórias:

| Coluna | Descrição |
|--------|-----------|
| **NOME PROCESSO** | Nome do processo a ser mapeado |
| **ATIVIDADE INÍCIO** | Indica se é ponto inicial (SIM/NÃO) |
| **ATIVIDADE ORIGEM** | Atividade de origem no fluxo |
| **PROCEDIMENTO** | Descrição do procedimento executado |
| **ATIVIDADE DESTINO** | Atividade de destino no fluxo |

⚠️ **Formato aceito**: Apenas arquivos **.xlsx**

## 🛠️ Instalação Local

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/gerar_fluxograma.git
cd gerar_fluxograma

# Crie um ambiente virtual
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
python app.py
```

Acesse: `http://localhost:5000`

## 🐳 Deploy com Docker

```bash
# Build da imagem
docker build -t gerar-fluxograma .

# Execute o container
docker run -p 5000:5000 gerar-fluxograma
```

Acesse: `http://localhost:5000`

## ☁️ Deploy no Render

### Passo a Passo

1. Faça push do código para o GitHub
2. Acesse [render.com](https://render.com) e faça login
3. Clique em **New +** → **Web Service**
4. Conecte seu repositório GitHub
5. Configure o serviço:
   - **Name**: gerar-fluxograma
   - **Environment**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
6. Clique em **Create Web Service**

O deploy será automático a cada push no GitHub! ✨

## 🎨 Personalização de Cores

As cores do fluxograma podem ser modificadas em `ProcFluxograma.py`:

```python
# Cores customizáveis
COLOR_ACTIVITY_FILL = "#00AE9D"  # Cor de fundo das atividades
COLOR_PROC_FILL = "#E8E8E8"      # Cor de fundo dos procedimentos
COLOR_START_FILL = "#87C2BC"     # Cor de fundo início/fim
COLOR_EDGE = "#00796B"           # Cor das setas/conexões
```

## 📦 Tecnologias Utilizadas

- **Backend**: 
  - Flask 3.x
  - Pandas
  - Graphviz
  - OpenPyXL

- **Frontend**: 
  - HTML5, CSS3, JavaScript
  - Drawflow (biblioteca de fluxogramas interativos)

- **Deploy**: 
  - Docker
  - Render.com
  - Gunicorn

## 🎯 Estrutura do Projeto

```
gerar_fluxograma/
├── app.py                    # Aplicação Flask principal
├── ProcFluxograma.py        # Processamento e geração de fluxogramas
├── requirements.txt         # Dependências Python
├── Dockerfile              # Configuração Docker
├── .gitignore              # Arquivos ignorados pelo Git
├── README.md               # Documentação
└── templates/
    └── index.html          # Interface web com Drawflow
```

## 🔧 Desenvolvimento

### Adicionar Novas Funcionalidades

1. Modifique `ProcFluxograma.py` para alterar a lógica de processamento
2. Edite `templates/index.html` para customizar a interface
3. Atualize `app.py` para adicionar novos endpoints

### Testes Locais

```bash
# Ative o ambiente virtual
source venv/bin/activate  # Linux/Mac

# Execute em modo debug
python app.py

# Acesse http://localhost:5000
```

## 📝 Licença

Este projeto está sob a licença MIT.

## 👤 Autor

**Guilherme Martins**
- GitHub: [@guivmartins](https://github.com/guivmartins)

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fork o projeto
2. Criar uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abrir um Pull Request

## 📞 Suporte

Se encontrar problemas ou tiver dúvidas:

1. Abra uma [Issue](https://github.com/seu-usuario/gerar_fluxograma/issues)
2. Verifique se já existe uma issue similar
3. Forneça detalhes sobre o problema e como reproduzi-lo
