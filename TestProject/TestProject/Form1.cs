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
using Esri.ArcGISRuntime.UI;
using Esri.ArcGISRuntime.Geometry;
using System.Text.RegularExpressions;

namespace TestProject
{
    public partial class Form1 : Form
    {
        
        
        static Application app = Globals.ThisAddIn.Application;
        static Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;        
        static bool mapCounter = true;
        string checkedStates;
        

        public Form1()
        {
            
            
            InitializeComponent();
            


        }


        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 2;
            this.checkAndLoadVariables();
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
            //const string userRoot = "HKEY_CURRENT_USER";
            //const string subkey = "MyRegistry";
            //const string keyName = userRoot + "\\" + subkey;

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
                Registry.CurrentUser.SetValue("from", textBox1.Text);
                

                //Додавання поля "хто завантажив"
                checkTextBox(textBox2, "loadedBy");
                Registry.CurrentUser.SetValue("loadedBy", textBox2.Text);

                //Додавання поля Важливість            
                doc.Fields.Update();
                checkComboBox(comboBox1, "reliability");

                //Додавання поля Достовірність
                checkComboBox(comboBox2, "sourceReliability");

                //Додавання поля Номер завдання
                checkTextBox(textBox7, "code");

                //Додавання поля ОР
                checkTextBox(textBox4, "objects");

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
            string[] priviousStatesISO = new string[1];
            bool flag = false;
            if (checkedStates!=null)
            {
                flag = true;
                priviousStatesISO = checkedStates.Split(';');                
            }
            


            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures & ~CultureTypes.NeutralCultures);
            foreach (CultureInfo culture in cultures)
            {
                try
                {
                    RegionInfo region = new RegionInfo(culture.LCID);
                    if (!(checkedListBox1.Items.Contains(region.DisplayName)))
                    {
                        checkedListBox1.Items.Add(region.DisplayName);
                        if (flag)
                        {
                            foreach (string element in priviousStatesISO)
                            {                                
                                if (element.Trim().Equals(region.ThreeLetterISORegionName))
                                {
                                    checkedListBox1.SetItemChecked(checkedListBox1.Items.Count - 1, true);
                                }
                            }
                        }
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
            //textBox1.Text = "";
            //textBox2.Text = "";
            //textBox3.Text = "";
            //textBox4.Text = "";
            //textBox5.Text = "";
            //textBox6.Text = "";
            //textBox7.Text = "";

            //comboBox1.SelectedIndex = 0;
            //comboBox2.SelectedIndex = 2;

            //foreach (int checkedItemIndex in checkedListBox1.CheckedIndices)
            //{
            //    checkedListBox1.SetItemChecked(checkedItemIndex, false);
            //}

            this.Close();

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
                closeCoordsPanel();
                Coordinates coordinates = new Coordinates(this);
                coordinates.setCoordsText(textBox5.Text);
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

        public string getCoords()
        {
            return textBox5.Text;
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

        private void checkAndLoadVariables()
        {
            if (Registry.CurrentUser.GetValue("from")!=null)
            {
                textBox1.Text = (string)Registry.CurrentUser.GetValue("from");
            }
            if (Registry.CurrentUser.GetValue("loadedBy") != null)
            {
                textBox2.Text = (string)Registry.CurrentUser.GetValue("loadedBy");
            }
            foreach (Variable var in doc.Variables)
            {
                if (var.Name.Equals("title"))
                {
                    textBox3.Text = var.Value;
                }
                else if (var.Name.Equals("objects"))
                {
                    textBox4.Text = var.Value;
                }
                else if (var.Name.Equals("coordinates"))
                {
                    textBox5.Text = var.Value;
                }
                else if (var.Name.Equals("tags"))
                {
                    textBox6.Text = var.Value;
                }
                else if (var.Name.Equals("code"))
                {
                    textBox7.Text = var.Value;
                }
                else if (var.Name.Equals("date"))
                {
                    dateTimePicker1.Value = Convert.ToDateTime(var.Value);
                }
                else if (var.Name.Equals("reliability"))
                {
                    comboBox1.SelectedItem = var.Value;
                }
                else if (var.Name.Equals("sourceReliability"))
                {
                    comboBox2.SelectedItem = var.Value;
                }
                else if (var.Name.Equals("states"))
                {
                    checkedStates = var.Value;
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            if (button3.Text.Equals(">"))
            {
                
                panel1.Visible = true;
                button3.Text = "<";
                convertCoordinates(textBox5.Text);
            } else
            {
                closeCoordsPanel();
            }            
        }

        private void closeCoordsPanel()
        {
            textBox9.Text = "";
            panel1.Visible = false;
            button3.Text = ">";
        }

        private void convertCoordinates(string coordinates)
        {
            try
            {
                string[] coordsArray = coordinates.Split('\b');
                try
                {
                    int pointIndex = 1;
                    int lineIndex = 1;
                    int polygonIndex = 1;
                    foreach (string coord in coordsArray)
                    {                        
                        if (coord.Contains(@"x"))
                        {
                            string point = "Point №" + pointIndex + Environment.NewLine;
                            Regex regex = new Regex(@"\-?\d+\.{1}\d+");
                            MatchCollection matches = regex.Matches(coord);
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                    point+= match.Value + ";" + Environment.NewLine;

                            }
                            pointIndex++;
                            textBox9.Text += point + Environment.NewLine + Environment.NewLine;
                        }
                        else if (coord.Contains("paths"))
                        {
                            string line = "Line №" + lineIndex + Environment.NewLine;
                            int pointsCounter = 1;
                            int pointNumber = 1;
                            Regex regex = new Regex(@"\-?\d+\.{1}\d+");
                            MatchCollection matches = regex.Matches(coord);
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                    if (pointsCounter == 1)
                                    {
                                        line += ">>Line point №" + pointNumber + Environment.NewLine;
                                        line += ">>" + match.Value + Environment.NewLine;
                                        pointNumber++;
                                        pointsCounter = 2;
                                    } else
                                    {
                                        line += ">>" + match.Value + Environment.NewLine;
                                        pointsCounter = 1;
                                    }
                            }
                            lineIndex++;
                            textBox9.Text += line + Environment.NewLine + Environment.NewLine;
                        }
                        else if (coord.Contains("rings"))
                        {
                            string polygon = "Polygon №" + polygonIndex + Environment.NewLine;
                            int pointsCounter = 1;
                            int pointNumber = 1;
                            Regex regex = new Regex(@"\-?\d+\.{1}\d+");
                            MatchCollection matches = regex.Matches(coord);
                            if (matches.Count > 0)
                            {
                                foreach (Match match in matches)
                                    if (pointsCounter == 1)
                                    {
                                        polygon += ">> Polygon point №" + pointNumber + Environment.NewLine;
                                        polygon += ">>" + match.Value + Environment.NewLine;
                                        pointNumber++;
                                        pointsCounter = 2;
                                    }
                                    else
                                    {
                                        polygon += ">>" + match.Value + Environment.NewLine;
                                        pointsCounter = 1;
                                    }
                            }
                            polygonIndex++;
                            textBox9.Text += polygon + Environment.NewLine + Environment.NewLine;
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
            catch (Exception)
            {

            }

        }
    }
}
