# Requisitos de Informação de Projeto (PIR)
**Projeto**: Demonstração de Workflow openBIM e Gestão de Informação  
**Autor**: Contratante (Simulação de Gestão de Ativos)  
**Versão**: 1.0  
**Data**: 28 de Junho de 2026  

---

## 1. Introdução

Este documento estabelece os **Requisitos de Informação de Projeto (PIR - Project Information Requirements)** para a simulação do nosso fluxo de trabalho openBIM. Ele define as informações críticas que devem ser fornecidas pela equipe de projeto (Contratada) nos modelos de dados em formato IFC (Industry Foundation Classes) para viabilizar auditorias automatizadas e planejamento de obras/gestão de ativos.

O cumprimento destes requisitos será verificado de forma objetiva por meio de um arquivo de validação computável **IDS (Information Delivery Specification)** e do respectivo relatório de auditoria gerado pelo validador.

---

## 2. Requisitos de Qualidade da Informação

Ficam definidos três requisitos principais de qualidade de dados aos quais todos os modelos de disciplinas estruturais e arquitetônicas submetidos devem obedecer:

### PIR-01: Verificação de Compartimentação e Cadastro de Espaços
*   **Objetivo**: Garantir que o modelo arquitetônico possui os limites espaciais definidos e que cada espaço possua seu nome descritivo longo (atributo `LongName`) preenchido para fins de identificação de ambientes.
*   **Critério de Aceitação**: O modelo deve conter instâncias da entidade `IfcSpace` e o atributo `LongName` deve estar definido com um valor de texto válido. A presença de espaços com o atributo `LongName` ausente ou vazio será considerada uma não conformidade.
*   **Estrutura de Saída Esperada**: O resultado da validação deste requisito deve ser consolidado em uma tabela com a seguinte estrutura:
    *   `ID do Espaço`: O identificador único global (`GlobalId`) do elemento `IfcSpace`.
    *   `LongName`: O nome descritivo completo atribuído ao espaço (atributo `LongName`).
    *   `Área do Espaço`: Área informada do espaço (para fins informativos, obtida de `GrossArea` ou `NetArea` no conjunto `Qto_SpaceBaseQuantities`, se disponível).
    *   `Nível (andar)`: O nível de referência no qual o espaço está associado (`IfcBuildingStorey`).


### PIR-02: Consistência de Volumes para Elementos Estruturais
*   **Objetivo**: Assegurar que os elementos que compõem a estrutura principal da edificação possuam informações quantitativas de volume preenchidas e válidas para subsidiar orçamentos e planejamento de concretagem.
*   **Elementos Alvo**: Pilares (`IfcColumn`), Vigas (`IfcBeam`), Lajes (`IfcSlab`) e Fundações (`IfcFooting`).
*   **Critério de Aceitação**: Para cada um dos elementos alvos presentes no modelo, a propriedade de volume líquida (normalmente `NetVolume` nos conjuntos de quantidades padrão, como `Qto_ColumnBaseQuantities`, `Qto_BeamBaseQuantities`, etc.) deve existir e possuir um valor estritamente maior que zero (Volume > 0).
*   **Estrutura de Saída Esperada**: Uma tabela detalhando:
    *   `ID do Objeto`: O `GlobalId` do elemento estrutural.
    *   `Name`: Nome do elemento.
    *   `Volume`: Valor da propriedade de volume líquida encontrada no modelo.
    *   `Nível (andar)`: O nível de referência no qual o elemento está contido (`IfcBuildingStorey`).
    *   `Status`: Indicação se o elemento "Atende" (Volume > 0) ou "Não atende" (Volume ausente ou igual a zero).

### PIR-03: Resistência ao Fogo de Elementos de Vedação (Paredes)
*   **Objetivo**: Garantir o preenchimento de parâmetros cruciais de segurança contra incêndio nos elementos de vedação vertical (paredes), necessários para a aprovação em órgãos reguladores e planos de fuga.
*   **Elementos Alvo**: Paredes (`IfcWall` e `IfcWallStandardCase`).
*   **Critério de Aceitação**: Toda parede modelada deve possuir a propriedade `FireRating` preenchida com um valor textual válido (ex: "60 min", "120 min", "Corta-Fogo 60"). Esta propriedade deve ser associada no conjunto padrão de propriedades `Pset_WallCommon`. Elementos com a propriedade ausente ou vazia serão considerados não conformes.
*   **Estrutura de Saída Esperada**: Uma tabela detalhando:
    *   `ID do Objeto`: O `GlobalId` da parede.
    *   `Name`: Nome do elemento.
    *   `Nível (andar)`: O nível de referência da parede (`IfcBuildingStorey`).
    *   `Fire Rating`: O valor da propriedade `FireRating` encontrado no modelo.
    *   `Status`: Indicação se o elemento "Atende" (propriedade preenchida com valor válido) ou "Não atende" (propriedade ausente ou vazia).

---

## 3. Fluxo de Validação de Conformidade

1.  A **Contratada** desenvolve os modelos BIM em conformidade com as diretrizes e gera a exportação para o formato openBIM `IFC`.
2.  Os modelos `IFC` são submetidos à plataforma de validação juntamente com o arquivo `requisitos.ids` fornecido pela contratante.
3.  O motor de validação processa as regras computáveis e emite o relatório visual e planilhado, identificando o nível de conformidade do projeto antes da aceitação do modelo pela contratante.
