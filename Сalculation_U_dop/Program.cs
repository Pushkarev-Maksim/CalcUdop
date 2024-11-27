using ASTRALib;
using LibraryEnergizingPowerLine;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Сalculation_U_dop
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Создаем экземпляр Rastr.
            IRastr rastr = new Rastr();
            // Путь до файла.
            string patch =
                @"C:\Users\maks_\OneDrive\Рабочий стол\УЧЁБА ТПУ\Магистратура\диплом\2024\СМЗУ\05022024_12_02_36\mdp_debug_1";
            // Загружаем файл с режимом.
            rastr.Load(RG_KOD.RG_REPL, patch, "");
            // Обращение к таблице ветви.
            ITable tableVetv = (ITable)rastr.Tables.Item("vetv");
            // Обращение к таблице узлы.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Тип ветви.
            ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");
            // Обращение к колонке Состояние ветви.
            ICol staVetv = (ICol)tableVetv.Cols.Item("sta");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");

            // Лист ЛЭП, для которых необходимо посчитать допустимое напряжение
            // перед постановкой ЛЭП под напряжения.
            var listLine = new List<Line>()
            {
                // добавить ветвь выключателя начала
                new Line(1106, "ВЛ 500 кВ Алтай - Итатская", 6125, 555, 0),
                new Line(580, "ВЛ 500 кВ Ангара - Камала", 8004, 6153, 0),
                new Line(551, "ВЛ 500 кВ Барнаульская - Рубцовская", 546, 548, 0),
                new Line(560, "ВЛ 500 кВ Братский ПП - Ново-Зиминская", 6110, 137, 0),
                new Line(526, "ВЛ 500 кВ Итатская - Томская", 6129, 6317, 0),
                new Line(541, "ВЛ 500 кВ Саяно-Шушенская ГЭС - Новокузнецкая №1", 6293, 6236, 1)
            };

            // Лист со значениями напряжений, используемых для проведения оперативного расчёта
            // максимально допустимых напряжений в зависимости от состава работающих ШР.
            var listUdopLine = new List<Line>()
            {
                // Алтай -  Итатская
                // Добавить индексы, чтобы не было строк.
                new Line("ВЛ 500 кВ Алтай - Итатская", "",
                "ПС Итатская и ПС Алтай - ШР отключены", 568, 568),
                new Line("ВЛ 500 кВ Алтай - Итатская", "",
                "ПС Итатская - 4РШ или 5РШ, ПС Алтай - ШР отключены", 568, 568),
                new Line("ВЛ 500 кВ Алтай - Итатская", "",
                "ПС Итатская - 4РШ и 5РШ, ПС Алтай - ШР отключены", 568, 568),

                // Ангара - Камала
                new Line("ВЛ 500 кВ Ангара - Камала", "",
                "УШР-580 включен или отключен", 568, 568),

                // Барнаульская - Рубцовская
                new Line("ВЛ 500 кВ Барнаульская - Рубцовская", "",
                "ШР отключены с двух сторон", 575, 575),

                // Братский ПП - Ново-Зиминская
                new Line("ВЛ 500 кВ Братский ПП - Ново-Зиминская", "УПК Тыреть 500 кВ расшунтирован",
                "Все ШР отключены", 551, 551),

                // Итатская - Томская
                new Line("ВЛ 500 кВ Итатская - Томская", "Ремонт ВЛ 500 кВ Ново-Анжерская-Томская",
                "Все ШР отключены", 568, 568),
                new Line("ВЛ 500 кВ Итатская - Томская", "Ремонт ВЛ 500 кВ Ново-Анжерская-Томская",
                "ПС Томская - ШР-500, ПС Итатская - все отключены", 568, 568),

                // Саяно-Шушенская ГЭС - Новокузнецкая №1. Не все
                // new Line("ВЛ 500 кВ Саяно-Шушенская ГЭС - Новокузнецкая №1", )
            };

            // Лист КП по напряжению со значениями границ графика напряжения.
            var listVoltageControlNode = new List<VoltageLowerLimit>()
            {
                new VoltageLowerLimit(563, "ПС 500 кВ Барнаульская", 505, 525),
                new VoltageLowerLimit(6267, "ПС 500 кВ Рубцовская", 495, 525),
                new VoltageLowerLimit(6272, "ПС 500 кВ Барабинская", 510, 525),
                new VoltageLowerLimit(6230, "ПС 500 кВ Юрга", 500, 525),
                new VoltageLowerLimit(6233, "ПС 500 кВ Новокузнецкая", 500, 525),
                new VoltageLowerLimit(210, "ПС 500 кВ Ново-Анжерская", 505, 525),
                new VoltageLowerLimit(294, "ПС 500 кВ Беловская ГРЭС", 500, 525),
                new VoltageLowerLimit(6140, "ПС 500 кВ Абаканская", 500, 525),
                new VoltageLowerLimit(6283, "ПС 500 кВ Саяно-Шушенская ГЭС", 520, 525),
                new VoltageLowerLimit(6117, "ПС 1150 кВ Итатская", 505, 525),
                new VoltageLowerLimit(40, "Назаровская ГРЭС", 510, 525),
                new VoltageLowerLimit(23, "ПС 500 кВ Красноярская", 500, 525),
                new VoltageLowerLimit(8000, "ПС 500 кВ Ангара", 510, 525),
                new VoltageLowerLimit(6148, "ПС 500 кВ Камала-1", 510, 525),
                new VoltageLowerLimit(6366, "Богучанская ГЭС", 520, 525),
                new VoltageLowerLimit(6220, "ПС 500 кВ Тайшет", 515, 525),
                new VoltageLowerLimit(2841, "ПС 500 кВ Братский ПП", 495, 525),
                new VoltageLowerLimit(805, "Братская ГЭС", 495, 525),
                new VoltageLowerLimit(6000, "Усть – Илимская ГЭС", 500, 525),
                new VoltageLowerLimit(133, "ПС 500 кВ Иркутская", 494, 525)
            };
            
            //rastr.rgm("");
            //Console.WriteLine("Напряжение по концам ЛЭП в исходном режиме");
            //Console.WriteLine(EnergizingPowerLine.GetVoltageNode(rastr, listLine, "ВЛ 500 кВ Барнаульская - Рубцовская"));
            
            //// Одностороннее включение ЛЭП
            //EnergizingPowerLine.OneWayLineComutation(rastr, listLine, 551, 1, 1);
            //rastr.rgm("");
            //Console.WriteLine("Напряжение по концам ЛЭП при одностороннем включении");
            //Console.WriteLine(EnergizingPowerLine.GetVoltageNode(rastr, listLine, "ВЛ 500 кВ Барнаульская - Рубцовская"));
            //string patch1 =
            //    @"C:\Users\maks_\OneDrive\Рабочий стол\УЧЁБА ТПУ\Магистратура\диплом\2024\СМЗУ\05022024_12_02_36\regim1";
            //rastr.Save(patch1, "");

            //// Включение ЛЭП с двух сторон
            //EnergizingPowerLine.TwoWayLineComutation(rastr, listLine, 551, 1);
            //rastr.rgm("");
            //Console.WriteLine("Напряжение по концам ЛЭП при включении с двух сторон");
            //Console.WriteLine(EnergizingPowerLine.GetVoltageNode(rastr, listLine, "ВЛ 500 кВ Барнаульская - Рубцовская"));
            //string patch2 =
            //    @"C:\Users\maks_\OneDrive\Рабочий стол\УЧЁБА ТПУ\Магистратура\диплом\2024\СМЗУ\05022024_12_02_36\regim2";
            //rastr.Save(patch2, "");

            //// Отключение ЛЭП с двух сторон
            //EnergizingPowerLine.TwoWayLineAnComutation(rastr, listLine, 551, 0);
            //rastr.rgm("");
            //Console.WriteLine("Напряжение по концам ЛЭП при отключении с двух сторон");
            //Console.WriteLine(EnergizingPowerLine.GetVoltageNode(rastr, listLine, "ВЛ 500 кВ Барнаульская - Рубцовская"));
            //string patch3 =
            //    @"C:\Users\maks_\OneDrive\Рабочий стол\УЧЁБА ТПУ\Магистратура\диплом\2024\СМЗУ\05022024_12_02_36\regim3";
            //rastr.Save(patch3, "");

            List<int> switchMCVRubtsovskaya = new List<int>()
            {
                // Р-1 ПС 500 кВ Барнаульская
                9,
                // Р-2 ПС 500 кВ Барнаульская
                10,
                // Р-3 ПС 500 кВ Барнаульская
                11,
                // Р-1 ПС 500 кВ Рубцовская
                7,
                // Р-2 ПС 500 кВ Рубцовская
                8,
                // Р-5537 ЕЭК
                40,
                // Р1-1106 ПС 1150 кВ Алтай
                14,
                // Р2-1106 ПС 1150 кВ Алтай
                15,
                // Р3-1104 ПС 1150 кВ Алтай
                12,
                // Р4-1104 ПС 1150 кВ Алтай
                13,
                // Р-Л-5370 ПС 500 кВ Семей
                41,
                // Р-Л-5384 ПС 500 кВ Семей
                42,
                // Р-5577 Экибастузская ГРЭС-1
                44,
                // БСК-1-220 ПС 220 кВ Светлая
                125,
                // БСК-2-220 ПС 220 кВ Светлая
                126,
            };

            List<int> switchMCVBarnaulskaya = new List<int>()
            {
                // Р-1 ПС 500 кВ Барнаульская
                9,
                // Р-2 ПС 500 кВ Барнаульская
                10,
                // Р-3 ПС 500 кВ Барнаульская
                11,
                // Р1-1106 ПС 1150 кВ Алтай
                14,
                // Р2-1106 ПС 1150 кВ Алтай
                15,
                // Р3-1104 ПС 1150 кВ Алтай
                12,
                // Р4-1104 ПС 1150 кВ Алтай
                13,
                // Р-532 ПС 500 кВ Заря
                6,
                // Р-1 ПС 500 кВ Рубцовская
                7,
                // Р-2 ПС 500 кВ Рубцовская
                8,
                // Р-541 ПС 500 кВ Новокузнецкая
                19,
                // Р-542 ПС 500 кВ Новокузнецкая
                20,
                // БСК-1-220 ПС 220 кВ Светлая
                125,
                // БСК-2-220 ПС 220 кВ Светлая
                126,
                // ПС 1150 кВ Итатская 1РШ
                23,
                // ПС 1150 кВ Итатская 2РШ
                24,
                // ПС 1150 кВ Итатская 3РШ
                25,
                // ПС 1150 кВ Итатская 4РШ
                26,
                // ПС 1150 кВ Итатская 5РШ
                27,
                // Р-1 ПС 500 кВ Ново-Анжерская
                16,
                // Р-2 ПС 500 кВ Ново-Анжерская
                17,
                // Р-534 ПС 500 кВ Барабинская
                5,
                // ШР-500 ПС 500 кВ Томская
                18,
            };


#if false
            Console.WriteLine("КП ПС 500 кВ Рубцовская");

            Parallel.ForEach<int>(
                  switchMCVRubtsovskaya,
                  Efficiency
            );

            void Efficiency(int MCV)
            {
                IRastr rastr = new Rastr();
                rastr.Load(RG_KOD.RG_REPL, patch, "");
                Console.WriteLine($"{EfficiencyMeansControlVoltage.CalculationEfficiencySwitchedMCV(
                    rastr, listVoltageControlNode, "ПС 500 кВ Рубцовская", MCV)}");
                Marshal.FinalReleaseComObject(rastr);
            }
#else
            // Расчет эффективности коммутируемых СРН для КП ПС 500 кВ Рубцовская 
            Console.WriteLine("КП ПС 500 кВ Рубцовская");

            foreach (int MCV in switchMCVRubtsovskaya)
            {
                rastr.Load(RG_KOD.RG_REPL, patch, "");
                Console.WriteLine($"{MCV}");
                Console.WriteLine($"{EfficiencyMeansControlVoltage.CalculationEfficiencySwitchedMCV(
                    rastr, listVoltageControlNode, "ПС 500 кВ Рубцовская", MCV)}");
            }

            // Расчет эффективности коммутируемых СРН для КП ПС 500 кВ Барнаульская
            Console.WriteLine("КП ПС 500 кВ Барнаульская");

            foreach (int MCV in switchMCVBarnaulskaya)
            {
                rastr.Load(RG_KOD.RG_REPL, patch, "");
                Console.WriteLine($"{MCV}");
                Console.WriteLine($"{EfficiencyMeansControlVoltage.CalculationEfficiencySwitchedMCV(
                    rastr, listVoltageControlNode, "ПС 500 кВ Барнаульская", MCV)}");
            }
#endif
        }
    }
}