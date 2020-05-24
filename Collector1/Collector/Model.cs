using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace Collector
{
    class Model
    {
        private WordprocessingDocument Doc { get; set; }
        private string pathToEXE = "";

        private string errorsString = "";
        public string ErrorsString { get { return errorsString; } set{ errorsString = value; } }
        private string DocPath { get; set; }
        private string DocFullName { get; set; }
        private string DocName { get; set; }
        private string Post { get; set; }
        private string AssignmentType { get; set; }
        private string Path { get; set; }

        private Table Table;

        private List<HFSession> Done = new List<HFSession>();

        public void SetPath(string path)
        {
            Path = path.Trim().ToString();
            SetPathToEXE();
        }
        public void SetPathToEXE()
        {
            string[] pathArray = Assembly.GetExecutingAssembly().Location.Split('\\');
            for (int i = 0; i < pathArray.Length-1;i++)
            {
                pathToEXE += pathArray[i] + @"\";
            }
        }
        public void LoadDoc()
        {
            if (Path.Length>6)
            {
                try
                {
                    CheckAndConvertDocToDocx(Path);
                    Doc = WordprocessingDocument.Open(Path, false);
                    DocPath = GetDirectoryPath(Path);
                    if (Doc != null)
                    {
                        Console.WriteLine("Doc was found");
                    }
                    else
                    {
                        Console.WriteLine("Doc was not found");
                    }
                } catch (Exception)
                {
                    Console.WriteLine("Wrong path 1");
                }
                
            } else
            {
                Console.WriteLine("Wrong path 2");
            }
            
        }
        private string GetDirectoryPath(string filepath)
        {
            string[] pathArray = filepath.Split('\\');
            string DirectoryPath = "";
            DocFullName = pathArray[pathArray.Length-1];
            for (int i = 0; i < pathArray.Length-1;i++)
            {
                DirectoryPath += pathArray[i] + @"\";
            }
            string[] nameArray = pathArray[pathArray.Length-1].Split('.');
            DocName = nameArray[0];
            getPostAndAssignmentType(DocName);
            return DirectoryPath;
        }
        public void getPostAndAssignmentType(string docName)
        {
            string pattern2 = @"[^0-9\s]+";
            string pattern3 = @"[^a-zA-Zа-яА-Я]+";
            if (docName.Contains(" "))
            {
                string[] postTypeArray = docName.Split(' ');
                Post = postTypeArray[0];
                AssignmentType = postTypeArray[1];
                //Console.WriteLine("Both: " + Post + " " + AssignmentType);
            }
            else if (Regex.IsMatch(docName, pattern2, RegexOptions.IgnoreCase)) {
                AssignmentType = docName;
                //Console.WriteLine("Only type: " + AssignmentType);
            }
            else if (Regex.IsMatch(docName, pattern3, RegexOptions.IgnoreCase))
            {
                Post = docName;
                //Console.WriteLine("Only post: " + Post);
            } 
        }
        public void LoadTable()
        {
            if (Doc!=null)
            {
                try
                {
                    Table = Doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                    Console.WriteLine("Table was found");
                }
                catch (Exception)
                {
                    Console.WriteLine("Table was not found");
                }
            }                                  
        }

        public void ScanTable()
        {
            if (Table!=null)
            {
                Console.WriteLine("Scanning...");
                IEnumerable<TableRow> rows = Table.Elements<TableRow>();
                int rowIndex = 0;
                foreach (TableRow row in rows)
                {
                    if (rowIndex!=0 && row.Descendants<TableCell>().Count() == 7)
                    {
                        HFSession hFSession = new HFSession();
                        hFSession.Post = Post;
                        hFSession.AssignmentType = AssignmentType;
                        bool dataFlag = false;
                        for (int i = 0; i < row.Descendants<TableCell>().Count(); i++)
                        {
                            string content = row.Descendants<TableCell>().ElementAt(i).InnerText;
                            switch (i) {
                                case 0:
                                    {
                                        if (content.Trim().Length > 0)
                                        {
                                            hFSession.Data = content;
                                            dataFlag = true;
                                        }
                                        break;
                                    }
                                case 1: hFSession.Time = content; break;
                                case 2:
                                    {
                                        if (dataFlag && content.Trim().Length > 0)
                                        {
                                            hFSession.Receiver = content;
                                        } else
                                        {
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Кому]\n";
                                        }
                                        break;
                                    }
                                case 3:
                                    {
                                        if (dataFlag && content.Trim().Length > 0)
                                        {
                                            hFSession.Transmitter = content;
                                        }
                                        else
                                        {
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Хто]\n";
                                        }
                                        break;
                                    }
                                case 4: hFSession.Frequency = content; break;
                                case 5:
                                    {
                                        if (dataFlag && content.Trim().Length > 0)
                                        {
                                            hFSession.Text = content;
                                        }
                                        else
                                        {
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Зміст]\n";
                                        }
                                        break;
                                    }
                                case 6: hFSession.Peleng = content; break;
                            }
                        }
                        if (dataFlag)
                        {
                            Done.Add(hFSession);
                        }
                    }
                    if (rowIndex != 0 && row.Descendants<TableCell>().Count() == 6)
                    {
                        HFSession hFSession = new HFSession();
                        hFSession.Post = Post;
                        hFSession.AssignmentType = AssignmentType;
                        bool dataFlag = false;
                        for (int i = 0; i < row.Descendants<TableCell>().Count(); i++)
                        {
                            string content = row.Descendants<TableCell>().ElementAt(i).InnerText;
                            switch (i)
                            {
                                case 0:
                                    {
                                        if (content.Trim().Length > 0)
                                        {
                                            hFSession.Data = content;
                                            dataFlag = true;
                                        }
                                        break;
                                    }
                                case 1: hFSession.Time = content; break;
                                case 2:
                                    {
                                        if (dataFlag && content.Trim().Length > 0)
                                        {
                                            hFSession.Receiver = content;
                                            hFSession.Transmitter = content;
                                        }
                                        else
                                        {
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Кому]\n";
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Хто]\n";
                                        }
                                        break;
                                    }
                                case 3: hFSession.Frequency = content; break;
                                case 4:
                                    {
                                        if (dataFlag && content.Trim().Length > 0)
                                        {
                                            hFSession.Text = content;
                                        }
                                        else
                                        {
                                            errorsString += "Error in " + (rowIndex + 1) + " row; and " + (i + 1) + " cell -> Missed [Зміст]\n";
                                        }
                                        break;
                                    }
                                case 5: hFSession.Peleng = content; break;
                            }
                        }
                        if (dataFlag)
                        {
                            Done.Add(hFSession);
                        }
                    }
                    rowIndex++;
                }
                if (errorsString.Trim().Length > 0)
                {
                    AddToError();                    
                } else
                {
                    AddToDone();
                }
            }
            
        }

        private void AddToError()
        {
            Console.WriteLine(pathToEXE);
            if (Directory.Exists(pathToEXE + @"Error"))
            {
                Doc.SaveAs(pathToEXE + @"Error\" + DocFullName);
            } else
            {
                Directory.CreateDirectory(pathToEXE + @"Error");
                Doc.SaveAs(pathToEXE + @"Error\" + DocFullName);
            }
            using (StreamWriter sw = File.CreateText(pathToEXE + @"Error\" + DocName + " Errors Description.txt"))
            {
                sw.WriteLine(errorsString);
            }
            Console.WriteLine("Errors were found");
        }

        private void AddToDone()
        {
            if (Directory.Exists(pathToEXE + @"Done"))
            {
                Doc.SaveAs(pathToEXE + @"Done\" + DocFullName);
            }
            else
            {
                Directory.CreateDirectory(pathToEXE + @"Done");
                Doc.SaveAs(pathToEXE + @"Done\" + DocFullName);
            }
            Directory.CreateDirectory(DocPath + DocName);
            int dataIndex = 1;
            foreach (HFSession element in Done)
            {
                using (StreamWriter sw = File.CreateText(DocPath + DocName + @"\Data #" + dataIndex + ".txt"))
                {
                    sw.WriteLine(element.GetHFSessionDescription());
                }
                dataIndex++;
            }

            Console.WriteLine("Scanned successfully");
        }

        private void CheckAndConvertDocToDocx(string path)
        {
            Application word = new Application();

            if (path.ToLower().EndsWith(".doc"))
            {
                Console.WriteLine("Formatting .doc -> .docx");
                var sourceFile = new FileInfo(path);
                var document = word.Documents.Open(sourceFile.FullName);

                string newFileName = sourceFile.FullName.Replace(".doc", ".docx");
                Path = newFileName;
                document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                 CompatibilityMode: WdCompatibilityMode.wdWord2010);

                word.ActiveDocument.Close();
                word.Quit();

                File.Delete(path);
                Console.WriteLine("Formated successfuly");
            }
        }
    }
}
