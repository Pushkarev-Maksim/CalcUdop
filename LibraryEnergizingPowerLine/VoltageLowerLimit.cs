namespace LibraryEnergizingPowerLine
{
    public class VoltageLowerLimit
    {
        /// <summary>
        /// Номер узла.
        /// </summary>
        private int _number;

        /// <summary>
        /// Название КП.
        /// </summary>
        private string _name;

        /// <summary>
        /// Нижняя граница графика напряжения.
        /// </summary>
        private double _ng;

        /// <summary>
        /// Верхняя граница графика напряжения.
        /// </summary>
        private double _vg;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="number"></param>
        /// <param name="ng"></param>
        /// <param name="vg"></param>
        public VoltageLowerLimit(int number, string name, double ng, double vg)
        {
            _number = number;
            _name = name;
            _ng = ng;
            _vg = vg;
        }

        /// <summary>
        /// Задание номера узла.
        /// </summary>
        public int Number
        {
            get { return _number; }
            set { _number = value; }
        }

        /// <summary>
        /// Задание названия КП.
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <summary>
        /// Задание нижней границы графика напряжения.
        /// </summary>
        public double Ng
        {
            get { return _ng; }
            set { _ng = value; }
        }

        /// <summary>
        /// Задание верхней границы графика напряжения.
        /// </summary>
        public double Vg
        {
            get { return _vg; }
            set { _vg = value; }
        }

        /// <summary>
        /// Получение информации о КП.
        /// </summary>
        /// <returns>Строка с данными полей объекта класса VoltageLowerLimit.</returns>
        public string GetInfo()
        {
            return $"Узел: {Number}, Имя: {Name}, Нижняя граница: {Ng}, Верхняя граница: {Vg}\n";
        }
    }
}