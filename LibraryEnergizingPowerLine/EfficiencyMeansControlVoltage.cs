using ASTRALib;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace LibraryEnergizingPowerLine
{
    public class EfficiencyMeansControlVoltage
    {
        /// <summary>
        /// Метод для определения эффективности коммутируемых СРН.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listVoltageControlNode">Лист КП.</param>
        /// <param name="nameVoltageControlNode">Название КП.</param>
        /// <param name="mCV">Номер СРН.</param>
        /// <returns>Строка со значениями эффективности СРН.</returns>
        public static string CalculationEfficiencySwitchedMCV(
            IRastr rastr, 
            List<VoltageLowerLimit> listVoltageControlNode,
            string nameVoltageControlNode, 
            int switchMCV)
        {
            // Обращение к таблице Реакторы.
            ITable tableReactor = (ITable)rastr.Tables.Item("Reactors");
            // Обращение к колонке Состояние реактора.
            ICol staReactor = (ICol)tableReactor.Cols.Item("sta");
            // Обращение к колонке Название реактора.
            ICol nameReactor = (ICol)tableReactor.Cols.Item("Name");

            // Перебор строк в листе listVoltageControlNode и определение
            // эффективности СРН в КП.
            int nodeNumber = 0;
            for (int i = 0; i < listVoltageControlNode.Count; i++)
            {
                if (listVoltageControlNode[i].Name == nameVoltageControlNode)
                {
                    nodeNumber = listVoltageControlNode[i].Number;
                }
            }

            double U = GetVoltageNode(rastr, nodeNumber);

            var setSelNodeR = "Id=" + switchMCV;
            tableReactor.SetSel(setSelNodeR);
            var indexReactor = tableReactor.FindNextSel[-1];
            string nameR = (string)nameReactor.ZN[indexReactor];

            while (indexReactor >= 0)
            {
                if ((bool)staReactor.Z[indexReactor] == false)
                {
                    return $"Реактор <<{nameR}>> уже в работе";
                }

                else if ((bool)staReactor.Z[indexReactor] == true)
                {
                    staReactor.Z[indexReactor] = 0;
                    indexReactor = tableReactor.FindNextSel[indexReactor];
                }
                break;
            }

            rastr.rgm("");

            double U1 = GetVoltageNode(rastr, nodeNumber);

            double eFF = Math.Round(Math.Abs(U1 - U), 3);

            ////Если понадобится выполнять параллельный расчет, то освобождаем также
            ////в методе все таблицы и колонки
            //Marshal.FinalReleaseComObject(tableReactor);
            //Marshal.FinalReleaseComObject(staReactor);
            //Marshal.FinalReleaseComObject(nameReactor);

            return $"Реактор - {nameR}\n" 
                + $"Напряжение до - {U} кВ\n" 
                + $"Напряжение после - {U1} кВ\n" 
                + $"Эффективность СРН - {eFF} кВ";
        }

        /// <summary>
        /// Метод для определения напряжения в указанном узле.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="nodeNumber">Номер узла.</param>
        /// <returns>Напряжение в узле.</returns>
        public static double GetVoltageNode(IRastr rastr, int nodeNumber)
        {
            // Обращение к таблице узлы.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");

            var setSelNodeVras = "ny=" + nodeNumber;
            tableNode.SetSel(setSelNodeVras);
            var nodeNumberVras = tableNode.FindNextSel[-1];
            double Unode = Math.Round((double)vrasNode.Z[nodeNumberVras], 3);

            ////Если понадобится выполнять параллельный расчет, то освобождаем также
            ////в методе все таблицы и колонки
            //Marshal.FinalReleaseComObject(tableNode);
            //Marshal.FinalReleaseComObject(vrasNode);

            return Unode;
        }

        public static string CalculationEfficiencyControlledMCV(
            IRastr rastr, 
            List<VoltageLowerLimit> listVoltageControlNode,
            string nameVoltageControlNode, 
            List<int> controllMCV)
        {
            // Обращение к таблице Генераторы УР.
            ITable tableGenerator = (ITable)rastr.Tables.Item("Generator");
            // Обращение к колонке Название генератора УР.
            ICol nameGenerator = (ICol)tableGenerator.Cols.Item("Name");
            // Обращение к колонке Номер узла генератора УР.
            ICol nodeGenerator = (ICol)tableGenerator.Cols.Item("Node");
            // Обращение к колонке Активная мощность генератора УР.
            ICol PGenerator = (ICol)tableGenerator.Cols.Item("P");
            // Обращение к колонке Реактивная мощность генератора УР.
            ICol QGenerator = (ICol)tableGenerator.Cols.Item("Q");
            // Обращение к колонке Минимальная реактивная мощность генератора УР.
            ICol QminGenerator = (ICol)tableGenerator.Cols.Item("Qmin");
            // Обращение к колонке Максимальная реактивная мощность генератора УР.
            ICol QmaxGenerator = (ICol)tableGenerator.Cols.Item("Qmax");

            // Обращение к таблице узлы.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");
            // Обращение к колонке Заданный модуль напряжения в узле.
            ICol vzdNode = (ICol)tableNode.Cols.Item("vzd");
            // Обращение к колонке Мощность генерации P в узле.
            ICol PNode = (ICol)tableNode.Cols.Item("pg");
            // Обращение к колонке Мощность генерации Q в узле.
            ICol QNode = (ICol)tableNode.Cols.Item("qg");
            // Обращение к колонке Мощность генерации Qmin в узле.
            ICol QminNode = (ICol)tableNode.Cols.Item("qmin");
            // Обращение к колонке Мощность генерации Qmax в узле.
            ICol QmaxNode = (ICol)tableNode.Cols.Item("qmax");

            // Перебор строк в листе listVoltageControlNode и определение
            // эффективности СРН в КП.
            int nodeNumber = 0;
            for (int i = 0; i < listVoltageControlNode.Count; i++)
            {
                if (listVoltageControlNode[i].Name == nameVoltageControlNode)
                {
                    nodeNumber = listVoltageControlNode[i].Number;
                }
            }
            
            // Напряжение в КП
            double U = GetVoltageNode(rastr, nodeNumber);

            double totalDeltaQ = 0;

            foreach (var genNumber in controllMCV)
            {
                var setSelParametersGenerator = "Num=" + genNumber;
                tableGenerator.SetSel(setSelParametersGenerator);

                if (tableGenerator.Count == 0)
                {
                    var setSelNodeGenerator = "ny=" + genNumber;
                    tableNode.SetSel(setSelNodeGenerator);
                    var generatorNode = tableNode.FindNextSel[-1];
                    double vzd = (double)vzdNode.Z[generatorNode];
                    double Vzad = 0;
                    vzdNode.set_ZN(generatorNode, Vzad);
                    double vzd1 = (double)vzdNode.Z[generatorNode];

                    double Qnach = (double)QNode.Z[generatorNode];
                    double Qmin = (double)QminNode.Z[generatorNode];
                    QNode.set_ZN(generatorNode, Qmin);
                    double Q = (double)QNode.Z[generatorNode];

                    double deltaQ = Math.Round(Math.Abs(Qnach - Q), 3);

                    totalDeltaQ += deltaQ;
                }
                else
                {
                    var nodeNumberGenerator = tableGenerator.FindNextSel[-1];
                    int nodeGen = (int)nodeGenerator.Z[nodeNumberGenerator];

                    var setSelNodeGenerator = "ny=" + nodeGen;
                    tableNode.SetSel(setSelNodeGenerator);
                    var generatorNode = tableNode.FindNextSel[-1];
                    double vzd = (double)vzdNode.Z[generatorNode];
                    double Vzad = 0;
                    vzdNode.set_ZN(generatorNode, Vzad);
                    double vzd1 = (double)vzdNode.Z[generatorNode];

                    double Qnach = (double)QGenerator.Z[nodeNumberGenerator];
                    double Qmin = (double)QminGenerator.Z[nodeNumberGenerator];
                    QGenerator.set_ZN(nodeNumberGenerator, Qmin);
                    double Q = (double)QGenerator.Z[nodeNumberGenerator];

                    double deltaQ = Math.Round(Math.Abs(Qnach - Q), 3);

                    totalDeltaQ += deltaQ;
                }
            }

            rastr.rgm("");

            // Напряжение в КП
            double U1 = GetVoltageNode(rastr, nodeNumber);

            double deltaU = Math.Round(Math.Abs(U - U1), 3);

            double eFF;
            
            if (deltaU == 0 & totalDeltaQ == 0)
            {
                return "Управляемое СРН загружено под минимум";
            }
            else
            {
                eFF = Math.Round(Math.Abs(totalDeltaQ / deltaU), 3);
            }

            return $"Эффективность управляемого СРН - {eFF} МВар/кВ";
        }

        /// <summary>
        /// Метод для получения параметров в указанном Генераторе.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="nodeNumber">Номер узла.</param>
        /// <returns>Напряжение в узле.</returns>
        public static string GetParametersGenerator(IRastr rastr, int genNumber)
        {
            // Обращение к таблице Генераторы УР.
            ITable tableGenerator = (ITable)rastr.Tables.Item("Generator");
            // Обращение к колонке Название генератора УР.
            ICol nameGenerator = (ICol)tableGenerator.Cols.Item("Name");
            // Обращение к колонке Номер узла генератора УР.
            ICol nodeGenerator = (ICol)tableGenerator.Cols.Item("Node");
            // Обращение к колонке Активная мощность генератора УР.
            ICol PGenerator = (ICol)tableGenerator.Cols.Item("P");
            // Обращение к колонке Реактивная мощность генератора УР.
            ICol QGenerator = (ICol)tableGenerator.Cols.Item("Q");
            // Обращение к колонке Минимальная реактивная мощность генератора УР.
            ICol QminGenerator = (ICol)tableGenerator.Cols.Item("Qmin");
            // Обращение к колонке Максимальная реактивная мощность генератора УР.
            ICol QmaxGenerator = (ICol)tableGenerator.Cols.Item("Qmax");

            var setSelParametersGenerator = "Num=" + genNumber;
            tableGenerator.SetSel(setSelParametersGenerator);
            var nodeNumberGenerator = tableGenerator.FindNextSel[-1];
            string nameGen = (string)nameGenerator.Z[nodeNumberGenerator];
            int nodeGen = (int)nodeGenerator.Z[nodeNumberGenerator];
            double PGen = Math.Round((double)PGenerator.Z[nodeNumberGenerator], 3);
            double QGen = Math.Round((double)QGenerator.Z[nodeNumberGenerator], 3);
            double QminGen = Math.Round((double)QminGenerator.Z[nodeNumberGenerator], 3);
            double QmaxGen = Math.Round((double)QmaxGenerator.Z[nodeNumberGenerator], 3);

            ////Если понадобится выполнять параллельный расчет, то освобождаем также
            ////в методе все таблицы и колонки
            //Marshal.FinalReleaseComObject(tableGenerator);
            //Marshal.FinalReleaseComObject(nameGenerator);
            //Marshal.FinalReleaseComObject(nodeGenerator);
            //Marshal.FinalReleaseComObject(PGenerator);
            //Marshal.FinalReleaseComObject(QGenerator);
            //Marshal.FinalReleaseComObject(QminGenerator);
            //Marshal.FinalReleaseComObject(QmaxGenerator);

            return $"Название генератора - {nameGen}\n"
                + $"Узел подключения генератора - {nodeGen}\n"
                + $"Активная мощность генерации - {PGen} МВт\n"
                + $"Реактивная мощность генерации - {QGen} МВар\n"
                + $"Минимальная реактивная мощность генерации - {QminGen} МВар\n"
                + $"Максимальная реактивная мощность генерации - {QmaxGen} МВар";
        }
    }
}