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
            // Перебор строк в листе listLine и определение с какой стороны ЛЭП
            // моделировать одностороннее включение или отключение.
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
            
            // Выборка в таблице ветви по узлу начала или конца.
            var setSelNode = "ip=" + numberNode + "|" + "iq=" + numberNode;
            tableVetv.SetSel(setSelNode);
            var indexVetv = tableVetv.FindNextSel[-1];
            
            // Обращение к колонке Конца ветви.
            ICol n_Con = (ICol)tableVetv.Cols.Item("iq");
            // Обращение к колонке Начала ветви.
            ICol n_Nach = (ICol)tableVetv.Cols.Item("ip");

            while (indexVetv >= 0)
            {
                if ((int)tipVetv.Z[indexVetv] == 2)
                {
                    int con = (int)n_Con.get_ZN(indexVetv);
                    int nach = (int)n_Nach.get_ZN(indexVetv);

                    int nodeReactor = -1;

                    if (con != numberNode)
                    {
                        nodeReactor = con;
                    }
                    else if (nach != numberNode)
                    {
                        nodeReactor = nach;
                    }

                    if (nodeReactor != -1)
                    {
                        tableReactor.SetSel("Id1=" + nodeReactor + "|" + "Id2=" + nodeReactor);
                        var indexReactor = tableReactor.FindNextSel[-1];
                        if (indexReactor == -1)
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
        public static string GetVoltageNode(IRastr rastr, List<Line> listLine, string nameLine)
        {
            // Обращение к таблице узлы.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к колонке Напряжение в узле.
            ICol vrasNode = (ICol)tableNode.Cols.Item("vras");

            int ip = 0; int iq = 0;

            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Name == nameLine)
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

            return $"{nameLine}:\n" +
                $"Напряжение в узле начала ({ip}) - {UBegin} кВ\n" +
                $"Напряжение в узле конца ({iq}) - {UEnd} кВ\n";
        }

        /// <summary>
        /// Метод для одностороннего включения ЛЭП.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        /// <param name="ip_iq">Начало или конец линии 0/1.</param>
        public static void OneWayLineComutation(IRastr rastr, List<Line> listLine,
            int numberLine, int _sta, int ip_iq)
        {
            // Обращение к таблице ветви.
            ITable tableVetv = (ITable)rastr.Tables.Item("vetv");
            // Обращение к таблице ветви.
            ITable tableNode = (ITable)rastr.Tables.Item("node");
            // Обращение к таблице Реакторы.
            ITable tableReactor = (ITable)rastr.Tables.Item("Reactors");
            // Обращение к колонке Тип ветви.
            ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");
            // Обращение к колонке Состояние ветви.
            ICol staVetv = (ICol)tableVetv.Cols.Item("sta");
            // Обращение к колонке Состояние узла.
            ICol staNode = (ICol)tableNode.Cols.Item("sta");
            // Перебор строк в листе listLine и определение с какой стороны ЛЭП
            // моделировать одностороннее включение или отключение.
            int numberNode = 0;
            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    switch (ip_iq)
                    {
                        // Начало линии
                        case 0:
                            {
                                numberNode = listLine[i].Ip;
                            }
                            break;
                        // Конец линии
                        case 1:
                            {
                                numberNode = listLine[i].Iq;
                            }
                            break;
                    }
                }
            }

            int ip = 0; int iq = 0;
            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    ip = listLine[i].Ip;
                    iq = listLine[i].Iq;
                }
            }

            // Выборка в таблице ветви по узлу начала или конца.
            var setSelNode = "ip=" + numberNode + "|" + "iq=" + numberNode;
            tableVetv.SetSel(setSelNode);
            var indexVetv = tableVetv.FindNextSel[-1];

            // Выборка в таблице узлы узел начала и конца ЛЭП.
            var setSelNode1 = "ny=" + ip + "|" + "ny=" + iq;
            tableNode.SetSel(setSelNode1);
            var indexNode = tableNode.FindNextSel[-1];

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

            while (indexNode >= 0)
            {
                switch (_sta)
                {
                    // Выкл ветвь
                    case 0:
                        {
                            staNode.Z[indexNode] = 1;
                        }
                        break;
                    // Вкл ветвь
                    case 1:
                        {
                            staNode.Z[indexNode] = 0;
                        }
                        break;
                }
                indexNode = tableNode.FindNextSel[indexNode];
            }
        }

        /// <summary>
        /// Метод для включения или отключения ЛЭП с двух сторон.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        public static void TwoWayLineComutation(IRastr rastr, List<Line> listLine,
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

            var setSelNode = "ip=" + ip + "|" + "iq=" + iq;
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
        /// Метод для включения или отключения ЛЭП с двух сторон.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        public static void TwoWayLineAnComutation(IRastr rastr, List<Line> listLine,
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
    }
}