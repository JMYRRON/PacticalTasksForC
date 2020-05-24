using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Collector
{
    class HFSession
    {
        private string post = "";
        public string Post { private get { return post; } set { post = value; } }

        private string assignmentType = "";
        public string AssignmentType { private get { return assignmentType; } set { assignmentType = value; } }

        private string data = "";
        public string Data { private get { return data; } set { data = value; } }

        private string time = "";
        public string Time { private get { return time; } set { time = value; } }

        private string receiver = "";
        public string Receiver { private get { return receiver; } set { receiver = value; } }

        private string transmitter = "";
        public string Transmitter { private get { return transmitter; } set { transmitter = value; } }

        private string frequency = "";
        public string Frequency { private get { return frequency; } set { frequency = value; } }

        private string peleng = "";
        public string Peleng { private get { return peleng; } set { peleng = value; } }

        private string text = "";
        public string Text { private get { return text; } set { text = value; } }


        public bool CheckAllData()
        {
            bool flag = false;
            if (Data.Trim() != "")
            {
                if (
                Receiver.Trim() != "" ||
                Transmitter.Trim() != "" ||
                Text.Trim() != "")
                {
                    flag = true;
                }
            }

                return flag;
        }

        public string GetHFSessionDescription()
        {
            string result = "Пост: " + post + "\n" +
                "Тип завдання: " + assignmentType + "\n" +
                "Дата: " + data + "\n" +
                "Час: " + time + "\n" +
                "Кому: " + receiver + "\n" +
                "Від кого: " + transmitter + "\n" +
                "Частота: " + frequency + "\n" +
                "Зміст: " + text + "\n" +
                "пеленг: " + peleng;

            return result;
        }

        public override string ToString()
        {
            return Post + " | " + AssignmentType + " | " + Data + " | " + Time + " | " + Receiver + " | " + Transmitter + " | " + Frequency + " | " +
                Frequency + " | " + Peleng + " | " + Text;
        }
    }
}
