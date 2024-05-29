using System;
using System.Xml;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Globalization;


namespace ExcelPasswordRemover
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        public void RemoveTagFromFile(string fileName,string NodeName)
        {
            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(fileName);

                XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
                namespaceManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                XmlNodeList sheetProtectionNodes = xmlDocument.SelectNodes("//ns:"+NodeName, namespaceManager);

                foreach (XmlNode node in sheetProtectionNodes)
                {
                    node.ParentNode.RemoveChild(node);
                }

                xmlDocument.Save(fileName);
                LogBox.Items.Add("Remove Protection From " + Path.GetFileNameWithoutExtension(fileName));
            }
            catch (Exception ex)
            {
                LogBox.Items.Add("Error occurred while removing Protection from Sheet: " + ex.Message);
            }
        }
        public void ExtractZipFile(string zipFilePath)
        {
            string tempFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp");
            try
            {
                if (Directory.Exists(tempFolderPath))
                {
                    Directory.Delete(tempFolderPath, true);
                }
                else
                {
                    Directory.CreateDirectory(tempFolderPath);
                }
                try
                {
                    ZipFile.ExtractToDirectory(zipFilePath, tempFolderPath);
                    LogBox.Items.Add("Let See whats inside");

                }
                catch (Exception ex)
                {
                    LogBox.Items.Add("Can not see: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                LogBox.Items.Add("dont mess with me: " + ex.Message);
            }



        }
        public void CompressDirectory(string directoryPath)
        {
            string zipFilePath = string.Empty;

            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel File |*.xlsx";
                    saveFileDialog.Title = "Save Excel File";
                    saveFileDialog.CheckFileExists = false;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        zipFilePath = saveFileDialog.FileName;

                        if (File.Exists(zipFilePath))
                        {
                            File.Delete(zipFilePath);
                        }

                        ZipFile.CreateFromDirectory(directoryPath, zipFilePath);
                        LogBox.Items.Add("Well Done Boy :)");
                    }
                    else
                    {
                        LogBox.Items.Add("WTF...");
                    }
                }
            }
            catch (Exception ex)
            {
                LogBox.Items.Add("WTF: " + ex.Message);
            }
            /*            finally
                        {
                            if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
                            {
                                // In case of any exception, delete the partially created compressed file
                                File.Delete(zipFilePath);
                            }
                        }*/
        }
        public  void ProcessXmlSheetFiles(string directoryPath)
        {
            try
            {
                string[] xmlFiles = Directory.GetFiles(directoryPath, "*.xml");

                foreach (string xmlFile in xmlFiles)
                {
                    Console.WriteLine($"Processing XML file: {xmlFile}");

                    // Perform your desired operations on the XML file here
                    // For example, you can call the RemoveTagFromFile function from the previous response

                    // Uncomment the following line if you want to remove the 'sheetProtection' tag from each XML file
                    RemoveTagFromFile(xmlFile, "sheetProtection");
                }

                Console.WriteLine("XML files processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred while processing XML files: " + ex.Message);
            }
        }
        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LogBox.Items.Add("Opening...");
                string selectedFilePath = openFileDialog.FileName;
                LogBox.Items.Add("Going Deep...");
                ExtractZipFile(selectedFilePath);
                LogBox.Items.Add("I am inside");
                try
                {
                    string sheetFilesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp\\xl\\worksheets");
                    LogBox.Items.Add("Try Remove Protection From Sheets");
                    ProcessXmlSheetFiles(sheetFilesPath);
                    LogBox.Items.Add("Sheets Finished");
                    LogBox.Items.Add("Try remove Protection From Workbook");
                    RemoveTagFromFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp\\xl\\workbook.xml"), "workbookProtection");
                    LogBox.Items.Add("Finished");
                    LogBox.Items.Add("ReGenerate Excel File...");
                    CompressDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp"));

                }
                catch (Exception)
                {

                    throw;
                }
            }
            string tempFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp");
            if (Directory.Exists(tempFolderPath))
            {
                Directory.Delete(tempFolderPath, true);
            }

        }
        private void LogBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LogBox.Items.Clear();
                btnOpenFile.Enabled = true;
                btnOpenFile.Visible = true;
        }
    }
}
