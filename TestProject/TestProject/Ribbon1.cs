using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.IO.Compression;
using System.IO;

namespace TestProject
{
    public partial class Ribbon1
    {

        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
                
        }

        public void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            int folderId = GenerateId();

            if (isVariables(doc.Variables))
            {
                updateVariables(doc.Variables);
            }
                       
            try
            {
                doc.Variables.Add("account", app.UserName);
                doc.Variables.Add("date", DateTime.Now+"");
                doc.Variables.Add("folderId", folderId);
                doc.Variables.Add("docName", doc.Name);
            } catch (Exception)
            {
                Console.WriteLine("Somethig went wrong");
            }

            

            try
            {

                doc.Save();

                
            } catch (Exception)
            {
                Console.WriteLine("Somethig went wrong");
            }
            string directoryPath = doc.Path + @"\" + folderId;

            Directory.CreateDirectory(directoryPath);

            string fileName = doc.Name;
            string sourcePath = doc.Path;
            string targetPath = directoryPath;

            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);

            File.Copy(sourceFile, destFile);

            string sourceFolder = targetPath;
            string targetZip = sourcePath + @"\" + folderId + ".zip";

            ZipFile.CreateFromDirectory(sourceFolder, targetZip);

            Directory.Delete(directoryPath, true);
        }

        private int GenerateId()
        {
            Random rand = new Random();
            return rand.Next(100000, 110000);        

        }

        private bool isVariables(Variables vars)
        {
            bool flag = false;
             foreach (Variable var in vars)
            {
                if (var.Name.Equals("account"))
                {
                    flag = true;
                    break;
                }
                else if (var.Name.Equals("date"))
                {
                    flag = true;
                    break;
                }
                else if (var.Name.Equals("folderId"))
                {
                    flag = true;
                    break;
                }
                else if (var.Name.Equals("docName"))
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        private void updateVariables(Variables vars)
        {
            Application app = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            
            foreach (Variable var in vars)
            {
                if (var.Name.Equals("account") && var.Value.Equals(app.UserName))
                {
                    var.Value.Replace(var.Value,app.UserName);
                }
                else if (var.Name.Equals("date") && var.Value.Equals(DateTime.Now))
                {
                    var.Value.Replace(var.Value, DateTime.Now+"");
                }
                else if (var.Name.Equals("docName") && var.Value.Equals(doc.Name))
                {
                    var.Value.Replace(var.Value, doc.Name);
                }
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
        }


    }
}
