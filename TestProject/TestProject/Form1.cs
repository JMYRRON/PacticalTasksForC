using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.IO.Compression;
using System.IO;
using Application = Microsoft.Office.Interop.Word.Application;
using Microsoft.Win32;
using System.Globalization;

namespace TestProject
{
    public partial class Form1 : Form
    {
        static Application app = Globals.ThisAddIn.Application;
        static Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

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
            if (textBox1.Text.Length != 0)
            {
                checkTextBox(textBox1, "from");
                Registry.SetValue(keyName, "from", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Fill the field 'Від'");
            }

            //Додавання поля "хто завантажив"
            if (textBox2.Text.Length != 0)
            {
                checkTextBox(textBox2, "loadedBy");
                Registry.SetValue(keyName, "loadedBy", textBox2.Text);
            }
            else
            {
                MessageBox.Show("Fill the field 'Хто'");
            }

            //Додавання поля Важливість
            
            doc.Fields.Update();
            //checkComboBox(comboBox1, "importance");

            //Додавання поля Достовірність
            checkComboBox(comboBox2, "validity");

            //Додавання поля Номер завдання
            checkTextBox(textBox7, "code");

            //Додавання поля ОР
            checkTextBox(textBox4, "object");

            //Додавання поля Координати
            checkTextBox(textBox5, "coordinates");

            //Додавання поля Заголовок
            if (textBox3.Text.Length != 0)
            {
                checkTextBox(textBox3, "title");
            }
            else
            {
                MessageBox.Show("Fill the field 'Заголовок'");
            }

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

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 59 && number != 32
                )
            {
                e.Handled = true;
            }
        }

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
                    
                } catch (ArgumentException ex)
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
                catch (ArgumentException ex)
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
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {            
            char letter = e.KeyChar;
            string result = textBox8.Text + letter;
            int index = checkedListBox1.FindString(result);
            checkedListBox1.SetSelected(index, true);
        }
    }
}
