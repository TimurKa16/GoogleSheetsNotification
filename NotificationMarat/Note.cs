using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificationMarat
{
    // Запись заявки
    public class Note
    {
        public int rowNumber { get; set; }
        public string customer { get; set; }
        public string country { get; set; }
        public string driver { get; set; }
        public DateTime date { get; set; }
        public NoteStatus status { get; set; }
        public string paymentSign { get; set; }
    }

    public enum NoteStatus
    {
        White = 0,
        Green = 1,
        Yellow = 2,
        Red = 3

    }
}
