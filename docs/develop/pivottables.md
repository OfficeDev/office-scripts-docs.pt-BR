---
title: Trabalhar com Tabelas Dinâmicas Office Scripts
description: Saiba mais sobre o modelo de objeto para Tabelas Dinâmicas na API JavaScript Office Scripts.
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 579f94140214674912c9610e707123924e4aef18
ms.sourcegitcommit: 4e3d3aa25fe4e604b806fbe72310b7a84ee72624
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/27/2022
ms.locfileid: "65077050"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Trabalhar com Tabelas Dinâmicas Office Scripts

As Tabelas Dinâmicas permitem que você analise grandes coleções de dados rapidamente. Com o poder deles vem a complexidade. As apIs Office scripts permitem que você personalize uma Tabela Dinâmica para atender às suas necessidades, mas o escopo do conjunto de APIs torna a introdução um desafio. Este artigo demonstra como executar tarefas comuns de Tabela Dinâmica e explica classes e métodos importantes.

> [!NOTE]
> Para entender melhor o contexto dos termos usados pelas APIs, leia Excel documentação da Tabela Dinâmica. Comece com [Criar uma Tabela Dinâmica para analisar dados da planilha](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## <a name="object-model"></a>Modelo de objetos

:::image type="content" source="../images/pivottable-object-model.png" alt-text="Uma imagem simplificada das classes, métodos e propriedades usadas ao trabalhar com Tabelas Dinâmicas.":::

A [Tabela Dinâmica é](/javascript/api/office-scripts/excelscript/excelscript.pivottable) o objeto central para Tabelas Dinâmicas na API Office Scripts.

- O [objeto Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) tem uma coleção de todas as [Tabelas Dinâmicas](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Cada [Planilha também](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contém uma coleção de Tabela Dinâmica que é local para essa planilha.
- Uma [Tabela Dinâmica](/javascript/api/office-scripts/excelscript/excelscript.pivottable) contém [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). Uma hierarquia pode ser considerada uma coluna em uma tabela.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) podem ser adicionados como linhas ou colunas ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), dados ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) ou filtros ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- Cada [PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) contém exatamente um [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). Estruturas de Tabela Dinâmica fora Excel podem conter vários campos por hierarquia, portanto, esse design existe para dar suporte a opções futuras. Para Office scripts, campos e hierarquias são mapeados para as mesmas informações.
- Um [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) contém vários [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Cada PivotItem é um valor exclusivo no campo. Pense em cada item como um valor na coluna da tabela. Os itens também podem ser valores agregados, como somas, se o campo estiver sendo usado para dados.
- O [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) define como os [PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) e [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem) são exibidos.
- [Os PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) filtram dados da [Tabela Dinâmica](/javascript/api/office-scripts/excelscript/excelscript.pivottable) usando critérios diferentes.

Veja como essas relações funcionam na prática. Os dados a seguir descrevem as vendas de frutas de várias fazendas. É a base para todos os exemplos neste artigo. Use <a href="pivottable-sample.xlsx">pivottable-sample.xlsx</a> para acompanhar.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="Uma coleção de vendas de frutas de diferentes tipos de fazendas diferentes.":::

## <a name="create-a-pivottable-with-fields"></a>Criar uma Tabela Dinâmica com campos

As Tabelas Dinâmicas são criadas com referências aos dados existentes. Os intervalos e tabelas podem ser a origem de uma Tabela Dinâmica. Eles também precisam de um lugar para existir na pasta de trabalho. Como o tamanho de uma Tabela Dinâmica é dinâmico, somente o canto superior esquerdo do intervalo de destino é especificado.

O snippet de código a seguir cria uma Tabela Dinâmica com base em um intervalo de dados. A Tabela Dinâmica não tem hierarquias, portanto, os dados ainda não estão agrupados de forma alguma.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="Uma Tabela Dinâmica chamada 'Farm Pivot' sem hierarquias.":::

### <a name="hierarchies-and-fields"></a>Hierarquias e campos

As Tabelas Dinâmicas são organizadas por meio de hierarquias. Essas hierarquias são usadas para dinamização de dados quando adicionadas como um tipo específico de hierarquia. Há quatro tipos de hierarquias.

- **Linha**: exibe itens em linhas horizontais.
- **Coluna**: exibe itens em colunas verticais.
- **Dados**: exibe agregações de valores com base nas linhas e colunas.
- **Filtro**: adiciona ou remove itens da Tabela Dinâmica.

Uma Tabela Dinâmica pode ter tantos ou poucos campos atribuídos a essas hierarquias específicas. Uma Tabela Dinâmica precisa de pelo menos uma hierarquia de dados para mostrar dados numéricos resumidos e pelo menos uma linha ou coluna para dinamizar esse resumo. O snippet de código a seguir adiciona duas hierarquias de linha e duas hierarquias de dados.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="Uma Tabela Dinâmica mostrando o total de vendas de frutas diferentes com base na fazenda de onde vieram.":::

## <a name="layout-ranges"></a>Intervalos de layout

Cada parte da Tabela Dinâmica é mapeada para um intervalo. Isso permite que o script obtenha dados da Tabela Dinâmica para uso posterior no script ou para serem retornados em um fluxo [Power Automate dados](power-automate-integration.md). Esses intervalos são acessados por meio [do objeto PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) adquirido de `PivotTable.getLayout()`. O diagrama a seguir mostra os intervalos retornados pelos métodos em `PivotLayout`.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="Um diagrama que mostra quais seções de uma Tabela Dinâmica são retornadas pelas funções de intervalo get do layout.":::

## <a name="filters-and-slicers"></a>Filtros e segmentações

Há três maneiras de filtrar uma Tabela Dinâmica.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` adicione uma hierarquia adicional para filtrar cada linha de dados. Qualquer linha com um item filtrado é excluída da Tabela Dinâmica e seus resumos. Como esses filtros são baseados em itens, eles só funcionam em valores discretos. Se "Classificação" for uma hierarquia de filtro em nossa amostra, os usuários poderão selecionar os valores "Orgânico" e "Convencional" para o filtro. Da mesma forma, se "Crates Sold Wholesale" for selecionado, as opções de filtro serão os números individuais, como 120 e 150, em vez de intervalos numéricos.

`FilterPivotHierarchies` são criados com todos os valores selecionados. Isso significa que nada é filtrado até que o usuário interaja manualmente com o controle de filtro ou um `PivotManualFilter` seja definido no campo pertencente a `FilterPivotHierarchy`.

O snippet de código a seguir adiciona "Classificação" como uma hierarquia de filtro.

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="Um controle de filtro que usa 'Classificação' para uma Tabela Dinâmica.":::

### <a name="pivotfilters"></a>PivotFilters

O `PivotFilters` objeto é uma coleção de filtros aplicados a um único campo. Como cada hierarquia tem exatamente um campo, você sempre deve usar o primeiro campo ao `PivotHierarchy.getFields()` aplicar filtros. Há quatro tipos de filtro.

- **Filtro de data**: filtragem baseada em data do calendário.
- **Filtro de rótulo**: filtragem de comparação de texto.
- **Filtro manual**: filtragem de entrada personalizada.
- **Filtro de valor**: filtragem de comparação de números. Isso compara os itens na hierarquia associada com valores em uma hierarquia de dados especificada.

Normalmente, apenas um dos quatro tipos de filtros é criado e aplicado ao campo. Se o script tentar usar filtros incompatíveis, um erro será gerado com o texto "O argumento é inválido ou está ausente ou tem um formato incorreto".

O snippet de código a seguir adiciona dois filtros. O primeiro é um filtro manual que seleciona itens em uma hierarquia de filtros de "Classificação" existente. O segundo filtro remove todos os farms que têm menos de 300 "Crates Vendidos no Atacado". Observe que isso filtra a "Soma" desses farms, não as linhas individuais dos dados originais.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="Uma Tabela Dinâmica depois que o filtro de valor e o filtro manual foram aplicados.":::

### <a name="slicers"></a>Segmentações de dados

[Segmentações filtram](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) dados em uma Tabela Dinâmica (ou tabela padrão). Eles são objetos movêveis na planilha que permitem seleções de filtragem rápida. Uma segmentação opera de maneira semelhante ao filtro manual e `PivotFilterHierarchy`. Os itens da tabela `PivotField` dinâmica são alternados para incluí-los ou excluí-los.

O snippet de código a seguir adiciona uma segmentação de dados para o campo "Tipo". Ele define os itens selecionados como "Lemon" e "Lime" e move a segmentação de dados 400 pixels para a esquerda.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="Uma segmentação de dados de filtragem em uma Tabela Dinâmica.":::

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
