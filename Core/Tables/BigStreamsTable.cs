﻿using Core.Abstracts;

namespace Core.Tables
{
    public class BigStreamsTable : Table<Grid>
    {
        public override Grid Headline { get => _headline; }
        private static readonly Grid _headline = new()
        {
            Cells = [new Grid("Состояние (в работе/отключена)"),
                new Grid("Дата последнего обследования"),
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
                new Grid("Русловые процессы"){
                    Cells = [new Grid("Ср"),
                             new Grid("Ам"),
                             new Grid("Вм"),
                             new Grid("характер")]},
                new Grid("Марка стали"),
                new Grid("Дата  последнего ВТД"),
                new Grid("Наличие дефектов"){
                    Cells = [new Grid($"с предельным сроком эксплуатации {DateTime.Now.Year}"){
                                Cells = [new Grid("пойма"),
                                         new Grid("русло")]},
                             new Grid($"с предельным сроком эксплуатации {DateTime.Now.Year + 1}"){
                                Cells = [new Grid("пойма"),
                                         new Grid("русло")]},
                             new Grid($"с предельным сроком эксплуатации {DateTime.Now.Year + 2}"){
                                Cells = [new Grid("пойма"),
                                         new Grid("русло")]}
                             ]},
                new Grid("Срок безопасной эксплуатации по отчету ОТС"),
                new Grid("Дата выдачи отчета по ОТС"),
                new Grid("Номер ОТС"),
                new Grid("Cрок безопасной эксплуатации по параметрам"){
                    Cells = [new Grid("Проведение ВТД"),
                             new Grid("Неустраненный дефект"),
                             new Grid("Диагностика спиралешовных труб (CDS)"),
                             new Grid("Диагностирование объектов не подлежащих ВТД"),
                             new Grid("Диагностика объектов с ограниченными возможностями ВТД"),
                             new Grid("Диагностика ВРК"),
                             new Grid("Диагностик запорной арматуры"),
                             new Grid("Диагностика приварных"),
                             new Grid("Диагностика соед. деталей"),
                             new Grid("Диагностика КПП СОД"),
                             new Grid("Диагностика дренажных емкостей"),
                             new Grid("Обследование ПВП и русловых процессов"),
                             new Grid("Обследование корр. сост.")]},
                new Grid("Организация, выдавшая заключение о сроке безопасной эксплуатации"),
                new Grid("Наличие мероприятий по приведению"),
                new Grid("Координаты"){
                    Cells = [new Grid("Левый берег"){
                                Cells = [new Grid("Широта, на момент обследования"),
                                         new Grid("Долгота, на момент обследования")]},
                             new Grid("Правый берег"){
                                Cells = [new Grid("Широта, на момент обследования"),
                                         new Grid("Долгота, на момент обследования")]}
                             ]},
                new Grid("ID"),
                new Grid("Расход воды"){
                    Cells = [new Grid("При максимальных значениях заданной обеспеченности, %") {
                                Cells = [new Grid("1"),
                                         new Grid("10")]},
                             new Grid("При среднемеженном уровне воды (Н меж)")
                            ]},
                new Grid("Максимальная скорость течения воды в русле"){
                    Cells = [new Grid("При максимальных значениях заданной обеспеченности, %") {
                                Cells = [new Grid("1"),
                                         new Grid("10")]},
                             new Grid("При среднемеженном уровне воды (Н меж)")
                            ]}
                ]
        };
    }
}
