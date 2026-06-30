# 🏢 Workflow openBIM & Gestão da Informação

Este repositório contém a demonstração funcional e interativa do workflow **openBIM** aplicado à Gestão da Informação em projetos de construção civil, utilizando o **Google Antigravity** como assistente de modelagem e validação.

O objetivo do protótipo é auditar modelos tridimensionais no formato neutro **IFC (Industry Foundation Classes)** confrontando-os com requisitos contratuais computáveis descritos no formato **IDS (Information Delivery Specification)**.

---

## 🚀 Como Testar a Aplicação

Siga o passo a passo abaixo para rodar a plataforma de validação localmente ou no Streamlit Cloud:

1. **Selecione o Modelo IFC** (no painel esquerdo):
   * Escolha um dos modelos já pré-carregados (ex: `TIPO1-ARQ-I4D_R03.ifc` para arquitetura ou `TIPO1-SCO-I4D_R03.ifc` para estrutura).
   * Ou faça o upload do seu próprio arquivo `.ifc`.
2. **Selecione o Arquivo de Regras IDS**:
   * Escolha o arquivo padrão de auditoria (`requisitos.ids`).
   * Ou envie um arquivo `.ids` customizado de acordo com suas regras.
3. **Execute a Auditoria**:
   * Clique em **"🚀 Executar Validação openBIM"** e aguarde o processamento.
4. **Explore os Resultados**:
   * **📊 Dashboard Geral**: Veja as métricas globais e o gráfico circular com o índice de conformidade.
   * **🎯 Validação Detalhada**: Analise elemento por elemento quais regras falharam e quais passaram.
   * **📋 Tabelas do PIR**: Inspecione as tabelas dedicadas a cada um dos Requisitos do Projeto.
   * **📑 Resposta Técnica (BEP)**: Leia o documento de requisitos do PIR e baixe a matriz de mapeamento de modelagem no formato Excel.
   * **💾 Exportar**: Baixe o relatório completo de conformidade.

---

## 📋 Requisitos de Informação do Projeto (PIR)

Esta demonstração foca na auditoria de 3 requisitos fundamentais descritos no documento contratual:

* **PIR-01: Compartimentação de Espaços (`IfcSpace`)**
  * **Regra**: Cada espaço modelado deve conter o preenchimento obrigatório de seu atributo identificador descritivo (**`LongName`**) para garantir o correto cadastro de ambientes.
* **PIR-02: Volume de Elementos Estruturais (`IfcColumn`, `IfcBeam`, `IfcSlab`, `IfcFooting`)**
  * **Regra**: Cada elemento de concreto ou metal deve apresentar o Property Set de quantidades básicas preenchido contendo um volume estrutural líquido válido (**`NetVolume` > 0**).
* **PIR-03: Resistência ao Fogo de Paredes (`IfcWall`)**
  * **Regra**: Paredes devem apresentar classificação de segurança contra incêndios através do parâmetro de resistência ao fogo (**`FireRating`** no conjunto `Pset_WallCommon`).

---

## 📑 Resposta Técnica - Matriz BEP (BIM Execution Plan)

Para orientar a equipe de projetistas a exportar os modelos IFC corretamente de acordo com os requisitos contratuais, a aba **"Resposta Técnica (BEP)"** disponibiliza o download de uma planilha formatada (`Anexo_BEP.xlsx`).
Esta planilha atua como um guia técnico instruindo:
* As classes de entidade IFC correspondentes a cada requisito.
* Os Property Sets e propriedades esperados no arquivo.
* A configuração correta dos botões de exportação no Revit (ex: habilitar "Exportar Quantidades Básicas" e parâmetros comuns de projeto).

---

## 🛠️ Tecnologias Utilizadas

* **Python 3**
* **Streamlit** (Interface e visualização interativa)
* **ifcopenshell** (Motor openBIM para leitura de dados IFC)
* **openpyxl** (Geração e estilização de planilhas de relatórios Excel)
* **Plotly** (Gráficos interativos)
* **buildingSMART IDS 1.0 XML** (Esquema de especificação de requisitos de informação)
