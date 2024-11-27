using ASTRALib;
using System.Runtime.InteropServices;

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
        public static string CalculationEfficiencySwitchedMCV(IRastr rastr, List<VoltageLowerLimit> listVoltageControlNode,
            string nameVoltageControlNode, int mCV)
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

            var setSelNodeR = "Id=" + mCV;
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

           //Если понадобится выполнять параллельный расчет, то освобождаем также
           //в методе все таблицы и колонки - Marshal.FinalReleaseComObject(tableNode);

            return $"Реактор - {nameR}\n" 
                + $"Напряжение до - {U} кВ\n" 
                + $"Напряжение после - {U1} кВ\n" 
                + $"Эффективность СРН - {eFF}";
        }

        public static double GetVoltageNode(IRastr rastr, int nodeNumber)
        {
            // Обращение к таблице ветви.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");

            var setSelNodeVras = "ny=" + nodeNumber;
            tableNode.SetSel(setSelNodeVras);
            var nodeNumberVras = tableNode.FindNextSel[-1];
            double Unode = Math.Round((double)vrasNode.Z[nodeNumberVras], 3);
            return Unode;
        }
    }
}