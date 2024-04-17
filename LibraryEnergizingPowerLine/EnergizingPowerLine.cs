using System.Collections.Generic;
using ASTRALib;

namespace LibraryEnergizingPowerLine
{
    public class EnergizingPowerLine
    {
        /// <summary>
        /// Метод для одностороннего включения или отключения ЛЭП.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        /// <param name="ip_iq">Начало или конец линии 0/1.</param>
        public static void Commutation(IRastr rastr, List<Line> listLine,
            int numberLine, int _sta, int ip_iq)
        {
            // Обращение к таблице ветви.
            ITable tableVetv = (ITable)rastr.Tables.Item("vetv");
            // Обращение к таблице Реакторы.
            ITable tableReactor = (ITable)rastr.Tables.Item("Reactors");
            // Обращение к колонке Тип ветви.
            ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");
            // Обращение к колонке Состояние ветви.
            ICol staVetv = (ICol)tableVetv.Cols.Item("sta");

            int numberNode = 0;
            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    // Начало линии
                    if (ip_iq == 0)
                    {
                        numberNode = listLine[i].Ip;
                    }
                    // Конец линии
                    else if (ip_iq == 1)
                    {
                        numberNode = listLine[i].Iq;
                    }
                }
            }

            var setSelNode = "ip=" + numberNode + "|" + "iq=" + numberNode;
            tableVetv.SetSel(setSelNode);
            var indexVetv = tableVetv.FindNextSel[-1];

            //ICol n_Con = (ICol)tableVetv.Cols.Item("iq");
            //ICol n_Nach = (ICol)tableVetv.Cols.Item("ip");
            //int con = (int)n_Con.get_ZN(indexVetv);
            //int nach = (int)n_Nach.get_ZN(indexVetv);
            //var setSelReactor = "id1=" + con + "|" + "id1=" + nach;
            //tableReactor.SetSel(setSelReactor);
            //var indexReactor = tableReactor.FindNextSel[-1];

            while (indexVetv >= 0)
            {
                if ((int)tipVetv.Z[indexVetv] == 2/* && indexReactor == -1*/)
                {
                    switch (_sta)
                    {
                        // Выкл ветвь
                        case 0:
                            {
                                staVetv.Z[indexVetv] = 1;
                            }
                            break;
                        // Вкл ветвь
                        case 1:
                            {
                                staVetv.Z[indexVetv] = 0;
                            }
                            break;
                    }
                }
                indexVetv = tableVetv.FindNextSel[indexVetv];
            }
        }

        /// <summary>
        /// Метод для включения или отключения ЛЭП с двух сторон.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        public static void Commutation(IRastr rastr, List<Line> listLine,
            int numberLine, int _sta)
        {
            // Обращение к таблице ветви.
            ITable tableVetv = (ITable)rastr.Tables.Item("vetv");
            // Обращение к колонке Состояние ветви.
            ICol staVetv = (ICol)tableVetv.Cols.Item("sta");

            int ip = 0; int iq = 0;

            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    ip = listLine[i].Ip;
                    iq = listLine[i].Iq;
                }
            }

            var setSelNode = "ip=" + ip + "&" + "iq=" + iq;
            tableVetv.SetSel(setSelNode);
            var indexVetv = tableVetv.FindNextSel[-1];

            while (indexVetv >= 0)
            {
                switch (_sta)
                {
                    // Выкл ветвь
                    case 0:
                        {
                            staVetv.Z[indexVetv] = 1;
                        }
                        break;
                    // Вкл ветвь
                    case 1:
                        {
                            staVetv.Z[indexVetv] = 0;
                        }
                        break;
                }
                indexVetv = tableVetv.FindNextSel[indexVetv];
            }
        }

        /// <summary>
        /// Метод для получения напряжения в узле начала и конца ЛЭП.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <returns>Строка со значениями напряжений в узле начала и конца ЛЭП.</returns>
        public static string GetVoltageNode(IRastr rastr, List<Line> listLine, int numberLine)
        {
            // Обращение к таблице узлы.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");

            int ip = 0; int iq = 0;

            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    ip = listLine[i].Ip;
                    iq = listLine[i].Iq;
                }
            }

            var setSelNodeBegin = "ny=" + ip;
            tableNode.SetSel(setSelNodeBegin);
            var nodeNumberBegin = tableNode.FindNextSel[-1];
            var UBegin = vrasNode.Z[nodeNumberBegin];

            var setSelNodeEnd = "ny=" + iq;
            tableNode.SetSel(setSelNodeEnd);
            var nodeNumberEnd = tableNode.FindNextSel[-1];
            var UEnd = vrasNode.Z[nodeNumberEnd];

            return $"ЛЭП №{numberLine}:\n" +
                $"Напряжение в узле начала ({ip}) - {UBegin} кВ\n" +
                $"Напряжение в узле конца ({iq}) - {UEnd} кВ\n";
        }
    }
}