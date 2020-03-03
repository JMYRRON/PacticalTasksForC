using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.IO.Compression;
using System.IO;
using Application = Microsoft.Office.Interop.Word.Application;
using Microsoft.Win32;
using System.Globalization;
using Esri.ArcGISRuntime.Mapping;
using System.Reflection;

namespace TestProject
{
    public partial class Form1 : Form
    {
        
        
        static Application app = Globals.ThisAddIn.Application;
        static Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;        
        static bool mapCounter = true;
        
        public Form1()
        {
            
            
            InitializeComponent();
            
                        
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 2;
            loadStates();
            checkedListBox1.Sorted = true;

           

            //Додавання GUID
            if (!isGUID())
            {
                doc.Variables.Add("guid", Guid.NewGuid().ToString());
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            const string userRoot = "HKEY_CURRENT_USER";
            const string subkey = "MyRegistry";
            const string keyName = userRoot + "\\" + subkey;

            if (checkObligatoryWindows())
            {
                //Додавання дати завантаженння
                if (isDates())
                {
                    addDate();
                }
                else
                {
                    doc.Variables.Add("dates", DateTime.Now);
                }

                //Додавання поля "Від"                
                checkTextBox(textBox1, "from");
                Registry.SetValue(keyName, "from", textBox1.Text);
                

                //Додавання поля "хто завантажив"
                checkTextBox(textBox2, "loadedBy");
                Registry.SetValue(keyName, "loadedBy", textBox2.Text);

                //Додавання поля Важливість            
                doc.Fields.Update();
                checkComboBox(comboBox1, "credibility");

                //Додавання поля Достовірність
                checkComboBox(comboBox2, "reliability");

                //Додавання поля Номер завдання
                checkTextBox(textBox7, "code");

                //Додавання поля ОР
                checkTextBox(textBox4, "object");

                //Додавання поля Координати
                checkTextBox(textBox5, "coordinates");

                //Додавання поля Заголовок
                checkTextBox(textBox3, "title");

                //Додавання поля Теги
                checkTextBox(textBox6, "tags");

                //Додавання поля Текст
                bool textFlag = true;
                foreach (Variable var in doc.Variables)
                {
                    if (var.Name.Equals("text"))
                    {
                        var.Value.Replace(var.Value, doc.Content.Text);
                        textFlag = false;
                        break;
                    }
                }
                if (textFlag)
                {
                    doc.Variables.Add("text", doc.Content.Text);
                }

                //Додавання поля Країни
                addStates(checkedListBox1, "states");

                //Додавання поля Дата отримання
                checkDateTimePicker(dateTimePicker1, "date");

                //Відкриття вікну підтвердження реєстрації та закриття форми
                MessageBox.Show("Реєстрація успішна");
                this.Close();
            }

            


        }

        private bool isGUID ()
        {            
            
            bool flag = false;
            foreach (Variable var in doc.Variables)
            {
                if (var.Name.Equals("guid"))
                {
                    flag = true;
                }
            }
            return flag;
        }

        private bool isDates()
        {
            bool flag = false;
            foreach (Variable var in doc.Variables)
            {
                if (var.Name.Equals("dates"))
                {
                    flag = true;
                }
            }
            return flag;
        }

        private void addDate()
        {
            string dates = "";
            foreach (Variable var in doc.Variables)
            {
                if (var.Name.Equals("dates"))
                {
                    dates = var.Value;
                    dates += "; " + DateTime.Now;
                    var.Value = dates;
                        
                }
            }

        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        //private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        //{
            //char number = e.KeyChar;
            //if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8  && number != 32 && number != 59)
            //{
                //e.Handled = true;
            //}
        //}

        private void checkTextBox (TextBox box, string varName)
        {
            bool flag = true;
            if (box.Text.Length != 0)
            {
                foreach (Variable var in doc.Variables)
                {
                    if (var.Name.Equals(varName))
                    {
                        doc.Variables[varName].Value = box.Text;
                        flag = false;
                        break;                        
                    }
                }
            }
            if (flag)
            {
                doc.Variables.Add(varName, box.Text);
            }
        }

        private void checkComboBox(ComboBox box, string varName)
        {
            bool flag = true;
            if (box.Text.Length != 0)
            {
                foreach (Variable var in doc.Variables)
                {
                    if (var.Name.Equals(varName))
                    {
                        doc.Variables[varName].Value = box.Text;
                        flag = false;
                        break;
                    }
                }
            }
            if (flag)
            {
                doc.Variables.Add(varName, box.Text);
            }
        }

        private void addStates(CheckedListBox box, string varName)
        {
            string result = "";
            if (box.CheckedItems.Count > 0)
            {
                foreach (string state in box.CheckedItems)
                {
                    string stateISO = getISO(state);
                    result += stateISO + "; ";
                }
                bool flag = true;
                foreach (Variable var in doc.Variables)
                {
                    if (var.Name.Equals(varName))
                    {
                        doc.Variables[varName].Value = result;
                        flag = false;
                        break;
                    }
                }
                if (flag)
                {
                    doc.Variables.Add(varName, result);
                }
            }
                
        }

        private void checkDateTimePicker(DateTimePicker box, string varName)
        {
            bool flag = true;
            foreach (Variable var in doc.Variables)
            {
                if (var.Name.Equals(varName))
                {
                    doc.Variables[varName].Value = box.Value + "";
                    flag = false;
                    break;
                }
            }
            if (flag)
            {
                doc.Variables.Add(varName, box.Value + "");
            }
        }

        private void loadStates()
        {
            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures & ~CultureTypes.NeutralCultures);
            foreach (CultureInfo culture in cultures)
            {
                try
                {
                    RegionInfo region = new RegionInfo(culture.LCID);
                    if (!(checkedListBox1.Items.Contains(region.DisplayName)))
                    {
                        checkedListBox1.Items.Add(region.DisplayName);
                    }
                    
                } catch (ArgumentException)
                {
                    continue;
                }             
                
            }

        }

        private string getISO (string state)
        {
            string result = "";
            List<string> cultureList = new List<string>();

            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures & ~CultureTypes.NeutralCultures);
                        
            foreach (CultureInfo culture in cultures)
            {
                try
                {
                    RegionInfo region = new RegionInfo(culture.LCID);

                    if (region.DisplayName.Equals(state))
                    {
                        result = region.ThreeLetterISORegionName;
                    }                        
                } 
                catch (ArgumentException)
                {
                    continue;
                }
                
            }
            
            return result;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 2;

            foreach (int checkedItemIndex in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemChecked(checkedItemIndex, false);
            }

        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {            
            char letter = e.KeyChar;
            string result = textBox8.Text + letter;
            int index = checkedListBox1.FindString(result);
            try
            {
                checkedListBox1.SetSelected(index, true);
            } catch (Exception)
            {

            }
            
        }

        private void textBox5_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (mapCounter)
            {
                Coordinates coordinates = new Coordinates(this);
                coordinates.Show();
                mapCounter = false;

            }
            else
            {
                MessageBox.Show("Карта відкрита");
            }
        }


            public void setCoords (string coords)
        {        
            textBox5.Text = coords;
        }

        private bool checkObligatoryWindows()
        {
            bool flag = true;
            
            if (textBox1.Text.Length == 0)
            {
                MessageBox.Show("Заповніть поле 'Від'");
                flag = false;
            } 
            else if (textBox2.Text.Length == 0)
            {
                MessageBox.Show("Заповніть поле 'Хто'");
                flag = false;
            }
            else if (textBox3.Text.Length == 0)
            {
                MessageBox.Show("Заповніть поле 'Заголовок'");
                flag = false;
            }


            return flag;
        }

        public static void changeMapCounter()
        {
            Form1.mapCounter = true;
        }

        
    }
}
