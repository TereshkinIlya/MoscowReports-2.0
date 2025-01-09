using Core.Abstracts;

namespace Core.Tables
{
    public class SmallStreamsTable : Table<Grid>
    {
        public override Grid Headline { get => _headline; }
        private static readonly Grid _headline = new()
        {
            Cells = [new Grid("Дата последнего обследования"),
                new Grid("Вид обследования"),
                new Grid("Положение МТ по отношению к ППРР (выше/ниже)"),
                new Grid("Разделение на пойму и русло"){
                    Cells = [new Grid("Наличие отклонений ПВП в русле."){
                                Cells = [new Grid("недостаточное заглубление, м"),
                                         new Grid("в т.ч. оголения, м"),
                                         new Grid("в т.ч. провис, м")]},
                             new Grid("Наличие отклонений ПВП в пойме."){
                                Cells = [new Grid("недостаточное заглубление, м"),
                                         new Grid("в т.ч. оголения, м"),
                                         new Grid("в т.ч. провис, м")]}
                             ]},
                new Grid("информация о ремонтах выявленных отклонений"),
                new Grid("Широта на момент обследования"),
                new Grid("Долгота на момент обследования"),
                new Grid("ID"),
                new Grid("Наличие отклонений ПВП в русле"){
                    Cells = [new Grid("недостаточное заглубление, м"),
                        new Grid("min толщина защитного слоя"),
                        new Grid("в т.ч. оголения, м"),
                        new Grid("max глубина оголения, м"),
                        new Grid("в т.ч. провис, м"),
                        new Grid("длина единичного участка с max протяженностью"),]
                }
            ]
        };
    }
}
