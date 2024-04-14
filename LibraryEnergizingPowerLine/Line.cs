namespace LibraryEnergizingPowerLine
{
    public class Line
    {
        /// <summary>
        /// Номер линии.
        /// </summary>
        private int _number;

        /// <summary>
        /// Наименование линии.
        /// </summary>
        private string _name;

        /// <summary>
        /// Начало линии.
        /// </summary>
        private int _ip;

        /// <summary>
        /// Конец линии.
        /// </summary>
        private int _iq;

        /// <summary>
        /// Номер параллельности.
        /// </summary>
        private int _np;

        /// <summary>
        /// Схема сети.
        /// </summary>
        private string _systemDiagram;

        /// <summary>
        /// Состав шунтирующих реакторов по концам ЛЭП.
        /// </summary>
        private string _compositionShuntReactor;

        /// <summary>
        /// Допустимое напряжение
        /// в узле начала линии.
        /// </summary>
        private int _uDop1;

        /// <summary>
        /// Допустимое напряжение
        /// в узле конца линии.
        /// </summary>
        private int _uDop2;

        /// <summary>
        /// Конструктор класса Line.
        /// </summary>
        /// <param name="ip">Начало линии.</param>
        /// <param name="iq">Конец линии.</param>
        /// <param name="np">Номер параллельности.</param>
        public Line(int number, string name, int ip, int iq, int np)
        {
            _number = number;
            _name = name;
            _ip = ip;
            _iq = iq;
            _np = np;
        }

        /// <summary>
        /// Конструктор класса Line.
        /// </summary>
        /// <param name="ip">Начало линии.</param>
        /// <param name="iq">Конец линии.</param>
        /// <param name="np">Номер параллельности.</param>
        public Line(string name, string systemDiagram,
            string compositionShuntReactor, int uDop1, int uDop2)
        {
            _name = name;
            _systemDiagram = systemDiagram;
            _compositionShuntReactor = compositionShuntReactor;
            _uDop1 = uDop1;
            _uDop2 = uDop2;
        }

        /// <summary>
        /// Задание номера ветви.
        /// </summary>
        public int Number
        {
            get { return _number; }
            set { _number = value; }
        }

        /// <summary>
        /// Задание наименования линии
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <summary>
        /// Задание начала ветви.
        /// </summary>
        public int Ip
        {
            get { return _ip; }
            set { _ip = value; }
        }

        /// <summary>
        /// Задание конца ветви.
        /// </summary>
        public int Iq
        {
            get { return _iq; }
            set { _iq = value; }
        }

        /// <summary>
        /// Задание номера параллельности.
        /// </summary>
        public int Np
        {
            get { return _np; }
            set { _np = value; }
        }

        /// <summary>
        /// Задание схемы сети.
        /// </summary>
        public string SystemDiagram
        {
            get { return _systemDiagram; }
            set { _systemDiagram = value; }
        }

        /// <summary>
        /// Задание состава работающих ШР.
        /// </summary>
        public string CompositionShuntReactor
        {
            get { return _compositionShuntReactor; }
            set { _compositionShuntReactor = value; }
        }

        /// <summary>
        /// Задание допустимого напряжения
        /// в узле начала линии.
        /// </summary>
        public int UDop1
        {
            get { return _uDop1; }
            set { _uDop1 = value; }
        }

        /// <summary>
        /// Задание допустимого напряжения
        /// в узле конца линии.
        /// </summary>
        public int UDop2
        {
            get { return _uDop2; }
            set { _uDop2 = value; }
        }
    }
}