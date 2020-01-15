using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GYM
{
    public class ClassEmployee
    {
        /// <summary>
        /// Идсотрудника
        /// </summary>
        public int id { get; set; }
        public string Surname { get; set; }
        public string Name { get; set; }
        /// <summary>
        /// Отчество
        /// </summary>
        public string Patrionymic { get; set; }
        /// <summary>
        /// Должность
        /// </summary>
        public string Position { get; set; }
        /// <summary>
        /// Телефон
        /// </summary>
        public string Phone { get; set; }
        public DateTime Date { get; set; }
        /// <summary>
        /// Пароль
        /// </summary>
        public int Passw { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Photo { get; set; }
        public int IdMoney { get; set; }
        public ClassEmployee()
        {
            this.id = 0;
            this.IdMoney = 0;
            this.Surname = "";
            this.Name = "";
            this.Patrionymic = "";
            this.Position = "";
            this.Phone = "";
            this.Passw = 0;
            this.Photo = "";
            this.Date = default(DateTime);
        }
        /// <summary>
        /// Инициализирует новый экземпляр класса Type

        public ClassEmployee(int id, string Surname, string Name, string Patrionymic, string Position, string Phone, int Passw, string Photo, DateTime Date, int IdMoney)
        {
            this.id = id;
            this.IdMoney = IdMoney;
            this.Surname = Surname;
            this.Name = Name;
            this.Patrionymic = Patrionymic;
            this.Position = Position;
            this.Phone = Phone;
            this.Passw = Passw;
            this.Photo = Photo;
            this.Date = Date;

        }
    } }