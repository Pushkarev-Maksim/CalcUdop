using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryEnergizingPowerLine
{
    public class Node
    {
        /// <summary>
        /// Номер узла.
        /// </summary>
        private int _number;

        /// <summary>
        /// Название узла.
        /// </summary>
        private string _name;

        /// <summary>
        /// Расчетный модуль напряжения.
        /// </summary>
        private double _voltage;

        /// <summary>
        /// Конструктор класса Node/
        /// </summary>
        /// <param name="number">Номер узла.</param>
        /// <param name="name">Название узла.</param>
        /// <param name="voltage">Расчетный модель напряжения.</param>
        public Node(int number, string name, double voltage)
        {
            _number = number;
            _name = name;
            _voltage = voltage;
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
        /// Задание названия узла.
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <summary>
        /// Задание расчетного напряжения.
        /// </summary>
        public double Voltage
        {
            get { return _voltage; }
            set { _voltage = value; }
        }
    }
}