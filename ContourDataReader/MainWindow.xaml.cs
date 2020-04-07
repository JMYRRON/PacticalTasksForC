using BarcodeLib.BarcodeReader;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using Spire;
using System.Drawing.Imaging;
using Spire.Pdf;
using Microsoft.Win32;

namespace ContourDataReader
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string path = AppDomain.CurrentDomain.BaseDirectory + "QRB.bmp";
        static OpenFileDialog openFileDialog = new OpenFileDialog();
        public MainWindow()
        {
            InitializeComponent();
            OpenFileDialogSettings();
            openFileDialog.FileOk += OpenFileDialog_FileOk;
        }

        private void OpenFileDialogSettings()
        {
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
        }

        private void OpenFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            TextBox2.Text = openFileDialog.FileName;
        }


        private void scanPDF()
        {
            if (TextBox2.Text.Length > 0)
            {
                try
                {
                    PdfDocument pdfDoc = new PdfDocument();
                    pdfDoc.LoadFromFile(TextBox2.Text);
                    TextBox1.Text = "";
                    for (int i = 0; i < pdfDoc.Pages.Count; i++)
                    {
                        System.Drawing.Image bmp = pdfDoc.SaveAsImage(i);
                        string fileName = string.Format("Page-{0}.png", i + 1);
                        bmp.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);

                        bool flag = false;

                        while (!flag)
                        {
                            string[] results = BarcodeLib.BarcodeReader.BarcodeReader.read(AppDomain.CurrentDomain.BaseDirectory + fileName, BarcodeLib.BarcodeReader.BarcodeReader.QRCODE);

                            foreach (string result in results)
                            {
                                try
                                {
                                    byte[] data = Convert.FromBase64String(result);
                                    string decodedString = Encoding.UTF8.GetString(data);
                                    if (decodedString.Contains("�"))
                                    {
                                        continue;
                                    } 
                                    else
                                    {
                                        flag = true;
                                        setTextBox(decodedString);
                                    }
                                }
                                catch (Exception) { }
                            }
                        }
                        
                    }
                } catch (Exception)
                {
                    MessageBox.Show("QR code not found");
                }
            }            
        }

        private void loadPDFDoc()
        {

        }

        private void setTextBox(string str)
        {
            string[] strArray = str.Split('|');
            foreach(string element in strArray)
            {
                TextBox1.Text += element + Environment.NewLine;
            }
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            scanPDF();
        }

        private void Button1_Click_1(object sender, RoutedEventArgs e)
        {
                     
            openFileDialog.ShowDialog();

        }

        private void Button4_Click(object sender, RoutedEventArgs e)
        {
            TextBox1.Text = "";
            TextBox2.Text = "";
        }
    }
}
