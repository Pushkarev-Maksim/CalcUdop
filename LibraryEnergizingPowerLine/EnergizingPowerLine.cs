﻿using System.Collections.Generic;
using ASTRALib;

namespace LibraryEnergizingPowerLine
{
    public class EnergizingPowerLine
    {
        Rastr rastr = new Rastr();

        /// <summary>
        /// Метод для одностороннего включения или отключения ЛЭП.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        /// <param name="ip_iq">Начало или конец линии 0/1.</param>
        public static void Commutation(Rastr rastr, List<Line> listLine,
            int numberLine, int _sta, int ip_iq)
        {
            var tables = rastr.Tables;
            var vetv = tables.Item("vetv");
            var tip = vetv.Cols.Item("tip");
            var sta = vetv.Cols.Item("sta");

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
            vetv.SetSel(setSelNode);
            var indexVetv = vetv.FindNextSel(-1);

            while (indexVetv >= 0)
            {
                if (tip.Z[indexVetv] == 2)
                {
                    switch (_sta)
                    {
                        // Выкл ветвь
                        case 0:
                            {
                                sta.Z[indexVetv] = 1;
                            }
                            break;
                        // Вкл ветвь
                        case 1:
                            {
                                sta.Z[indexVetv] = 0;
                            }
                            break;
                    }
                }
                indexVetv = vetv.FindNextSel(indexVetv);
            }
        }

        /// <summary>
        /// Метод для вклюяения или отключения ЛЭП с двух сторон.
        /// </summary>
        /// <param name="rastr">Экземпляр Rastr.</param>
        /// <param name="listLine">Лист ЛЭП.</param>
        /// <param name="numberLine">Номер линии.</param>
        /// <param name="_sta">Состояние ВКЛ/ВЫКЛ.</param>
        public static void Commutation(Rastr rastr, List<Line> listLine,
            int numberLine, int _sta)
        {
            var tables = rastr.Tables;
            var vetv = tables.Item("vetv");
            var sta = vetv.Cols.Item("sta");

            int ip = 0; int iq = 0;

            for (int i = 0; i < listLine.Count; i++)
            {
                if (listLine[i].Number == numberLine)
                {
                    ip = listLine[i].Ip;
                    iq = listLine[i].Iq;
                }
            }

            var setSelNode = "ip=" + ip + "|" + "ip=" + iq;
            vetv.SetSel(setSelNode);
            var indexVetv = vetv.FindNextSel(-1);

            while (indexVetv >= 0)
            {
                switch (_sta)
                {
                    // Выкл ветвь
                    case 0:
                        {
                            sta.Z[indexVetv] = 1;
                        }
                        break;
                    // Вкл ветвь
                    case 1:
                        {
                            sta.Z[indexVetv] = 0;
                        }
                        break;
                }
                indexVetv = vetv.FindNextSel(indexVetv);
            }
        }
    }
}