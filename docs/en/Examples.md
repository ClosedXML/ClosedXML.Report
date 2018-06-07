---
title: Примеры шаблонов
---

### Простой шаблон
![simple](https://user-images.githubusercontent.com/1150085/41095320-b5a5d990-6a59-11e8-8145-245de5d174b5.png)

Вы можете применять к ячейкам любое форматирование, включая условные форматы.

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Simple.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Simple.xlsx)

### Сортировка списка
![tlists1_sort](https://user-images.githubusercontent.com/1150085/41095319-b5852ace-6a59-11e8-8cbc-1cb9c77ef614.png)

Вы можете отсортировать список по столбцам. Просто укажите тэг `<<sort>>` в ячейках опций соответствующих столбцов. Чтобы отсортировать в порядке убывания, добавьте опцию «Desc» в параметр тэга сортировки (`<<sort desc>>`). 

Подробнее смотрите раздел [Сортировка данных](Sorting)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists1_sort.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists1_sort.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists1_sort.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists1_sort.xlsx)

### Итоги по столбцам
![tlists2_sum](https://user-images.githubusercontent.com/1150085/41095318-b556cbac-6a59-11e8-9467-aab8da0f46fa.png)

Вы можете получить итоговые значения для столбца диапазона, указав соответствующий параметр в ячейке опции столбца. 
В этом примере мы использовали тэг `<<sum>>` в строке опций списка для колонки Amount paid.

Подробнее смотрите раздел [Подитоги по столбцам](Totals-in-a-column).

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists2_sum.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists2_sum.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists2_sum.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists2_sum.xlsx)

### Опции списка и колонок
![tlists3_options](https://user-images.githubusercontent.com/1150085/41095317-b534b968-6a59-11e8-8bd9-f1b3e06d9052.png)

Помимо передачи данных из набора данных в диапазон, ClosedXML.Report может сортировать диапазон, составлять итоговые значения, группировать диапазон и так далее. Эти действия выполняются ClosedXML.Report, если он находит тэги диапазона и тэги столбца в соответствующих ячейках. 

Подробнее смотрите раздел [Плоские таблицы](Flat-tables)

В этом примере мы выравняли колонки по содержимому, добавили фильтры в заголовок таблицы, заменили формулы на значения и защитили колонку Amount paid от изменения. Для этого мы использовали тэги: <<AutoFilter>>, <<ColsFit>>, <<OnlyValues>> и <<Protected>>.

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists3_options.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists3_options.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists3_options.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists3_options.xlsx)

### Сложный диапазон
![tlists4_complexrange](https://user-images.githubusercontent.com/1150085/41095316-b511fa22-6a59-11e8-90ee-11983c29f228.png)

ClosedXML.Report способен использовать многострочные шаблоны для строки таблицы. Вы можете произвольным образом форматировать ячейки, объединять их, использовать условное форматирование, формулы Excel.

Подробнее смотрите раздел [Плоские таблицы](Flat-tables)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists4_complexrange.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tlists4_complexrange.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists4_complexrange.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tlists4_complexrange.xlsx)

### Группировка списка
![grouptagtests_simple](https://user-images.githubusercontent.com/1150085/41095313-b4931464-6a59-11e8-93d5-502642425bb4.png)

С тэгом `<<group>>` могут использоваться все тэги суммирования. Укажите опцию `<<group>>` в ячейках опций столбцов, где вы хотите получить промежуточные итоги.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_simple.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_simple.xlsx)

### Схлопнутые группы
![grouptagtests_collapse](https://user-images.githubusercontent.com/1150085/41095309-b426c12e-6a59-11e8-800b-e65d46ff5d8f.png)

Используйте параметр collapse тэга group (`<<group collapse>>`), если вы хотите быстро отображать только строки, которые содержат сводки или заголовки для разделов вашего листа.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_collapse.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_collapse.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_collapse.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_collapse.xlsx)

### Заголовки над данными
![grouptagtests_summaryabove](https://user-images.githubusercontent.com/1150085/41095314-b4caf14a-6a59-11e8-9f36-8051306a52ad.png)

ClosedXML.Report реализует тэг SUMMARYABOVE. Он помещает строку итогов над сгруппированными данными. 

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_summaryabove.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_summaryabove.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_summaryabove.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_summaryabove.xlsx)

### Объединение заголовков (вариант 1)
![grouptagtests_mergelabels](https://user-images.githubusercontent.com/1150085/41095305-b3c1891c-6a59-11e8-8b64-295ae87c7df1.png)

Тэг `<<group>>` позволяет объединять ячейки в сгруппированном столбце. Эта функция доступна с помощью параметра 'mergelabels' тэга `<<group>>`.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_mergelabels.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_mergelabels.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_mergelabels.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_mergelabels.xlsx)

### Объединение заголовков (вариант 2)
![grouptagtests_mergelabels2](https://user-images.githubusercontent.com/1150085/41095307-b3e50b26-6a59-11e8-8f57-536997dcd3cd.png)

Тэг `<<group>>` позволяет группировать данные без вставки строки заголовка группы. Эта функция доступна с помощью параметра 'MergeLabels=Merge2' тэга `<<group>>`. Ячейки, содержащие сгруппированные данные, объединяются, а ячейка результата содержит сгруппированное значение.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_mergelabels2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_mergelabels2.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_mergelabels2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_mergelabels2.xlsx)

### Вложенные группы
![grouptagtests_nestedgroups](https://user-images.githubusercontent.com/1150085/41095310-b44b4968-6a59-11e8-94e6-4c7e13477a38.png)

Списки могут быть сгруппированы с любым уровнем вложенности.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_nestedgroups.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_nestedgroups.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_nestedgroups.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_nestedgroups.xlsx)

### Отключенное схлопывание групп
![grouptagtests_disableoutline](https://user-images.githubusercontent.com/1150085/41095303-b39ac3ea-6a59-11e8-8548-802c8369cf1a.png)

Используйте параметр `disableoutline` тэга `<<group>>` чтобы отключить схлопывание групп. В этом примере диапазон группируется по колонке Company и Payment method. Схлопывание групп столбца Payment method отключено.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_disableoutline.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_disableoutline.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_disableoutline.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_disableoutline.xlsx)

### Размещение заголовка группы
![grouptagtests_placetocolumn](https://user-images.githubusercontent.com/1150085/41095311-b46ef08e-6a59-11e8-8f3c-7bfb8e63f338.png)

Тэг `<<group>>` позволяет размещать заголовок группы в любой столбец сгруппированного диапазона с помощью параметра `PLACETOCOLUMN=n`, где n - номер столбца в диапазоне (начиная с 1). Так же ClosedXML.Report содержит тэг `<<delete>>`, позволяющий удалить столбец или строку. В примере группировка выполняется по колонке Company с использованием параметра `mergelabels`. Заголовок группы столбца Company помещается во второй столбец диапазона (параметр `PLACETOCOLUMN=2`). Затем удаляется столбец Company.

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_placetocolumn.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_placetocolumn.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_placetocolumn.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_placetocolumn.xlsx)

### Группировка с заголовками
![grouptagtests_withheader](https://user-images.githubusercontent.com/1150085/41095315-b4ec70a4-6a59-11e8-99e3-327f123570a9.png)

Вы можете настроить отображение заголовка группы с помощью параметра `WITHHEADER` тэга `<<group>>`. В этом случае заголовок группы помещается над сгруппированными данными и сводной строкой ниже данных. Параметр `SUMMARYABOVE` не влияет на это размещение. 

Подробнее смотрите раздел [Группировка](Grouping)

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_withheader.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/grouptagtests_withheader.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_withheader.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/grouptagtests_withheader.xlsx)

### Вложенные области
![subranges_simple_tmd1](https://user-images.githubusercontent.com/1150085/41095301-b34dcc84-6a59-11e8-8252-b45e23419e1d.png)

Вы можете вложить диапазон в другой диапазон, отражающий таким образом подчиненные отношения ваших данных. В примере диапазон Items вложен в Orders, а последний - в Customers. Все три диапазона имеют собственную строку параметров плюс ту же левую границу и ту же ширину.

Подробнее смотрите раздел [Вложенные области: отчет с детализацией](Nested-ranges_-Master-detail-reports).

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_simple_tmd1.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_simple_tmd1.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_simple_tmd1.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_simple_tmd1.xlsx)

### Вложенные области с итогами
![subranges_withsubtotals_tmd2](https://user-images.githubusercontent.com/1150085/41095302-b37596e2-6a59-11e8-83ff-e9d5859ab86f.png)

Вы можете использовать тэги суммирования на каждом уровне вложенности в отчете master-detail. В примере тэг `<<sum>>` в ячейке I9 будет суммировать ячейки `{{item.Discount}}` по заказу, в то время как тот же тэг в ячейке I10 суммирует эти ячейки по поставщику.

Подробнее смотрите раздел [Вложенные области: отчет с детализацией](Nested-ranges_-Master-detail-reports).

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_withsubtotals_tmd2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_withsubtotals_tmd2.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_withsubtotals_tmd2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_withsubtotals_tmd2.xlsx)

### Вложенные области с внутренней сортировкой
![subranges_withsort_tmd3](https://user-images.githubusercontent.com/1150085/41095300-b310d2c0-6a59-11e8-8753-937159d9f8fd.png)

Вы можете использовать тэг `<<sort>>` в самом внутреннем диапазоне.

Подробнее смотрите раздел [Вложенные области: отчет с детализацией](Nested-ranges_-Master-detail-reports).

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_withsort_tmd3.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/subranges_withsort_tmd3.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_withsort_tmd3.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/subranges_withsort_tmd3.xlsx)

### Сводный отчёт
![tpivot5_static](https://user-images.githubusercontent.com/1150085/41095299-b2ee2946-6a59-11e8-849c-ad38bec29b28.png)

ClosedXML.Report поддерживает мощное средство работы со сводными таблицами. Вы можете разместить одну или несколько сводных таблиц прямо в шаблоне отчета, воспользовавшись удобством мастера сводных таблиц Excel и практически всеми возможностями в их оформлении и структурировании.

Подробнее смотрите раздел [Сводные таблицы](Pivot-tables).

Шаблон: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tpivot5_static.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tpivot5_static.xlsx)

Результат: [https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tpivot5_static.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tpivot5_static.xlsx)
