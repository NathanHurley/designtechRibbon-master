using Autodesk.Revit.DB;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DesignTechRibbonPaid.Revit.EssentialTools.ExportCategoryParametersForm
{
    public partial class ExportCategoryParametersForm : System.Windows.Forms.Form
    {
        #region Variables
        Document localDoc;
        
        public FilteredElementCollector coll;
        public SortedList<string, Category> myCategory;
        public string currentName;
        List<Element> found  = new List<Element>();
        
        List<string> sortInstanceHeader = new List<string>();
        List<string> sortTypeHeader = new List<string>();

        List<Tuple<string, ElementId>> sortedInstanceParameters = new List<Tuple<string, ElementId>>();
        List<Tuple<string, ElementId>> sortedTypeParameters = new List<Tuple<string, ElementId>>();
        FamilySymbol fam;

        List<Tuple<string, Element>> typeList = new List<Tuple<string, Element>>();
        List<Element> sortedTypeElements = new List<Element>();

        GetSetData Variables = new GetSetData();
        string folderLocation = "";
        object templateWorkBook;
        bool run = false;
        #endregion

        public ExportCategoryParametersForm(Document doc)
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            localDoc = doc;

            coll = new FilteredElementCollector(localDoc).WhereElementIsNotElementType();

            GetListOfElements();

            cancelButton.Enabled = false;
            ExportToExcel.Enabled = false;
        }

        public class GetSetData
        {
            public Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            public Microsoft.Office.Interop.Excel.Worksheet worksheet;
            public Microsoft.Office.Interop.Excel.Worksheet worksheet2;

            public Excel.Worksheet newWorksheet;
            public Excel.Workbook xlWorkBookTemplate;
            
            public Excel.Worksheet xlWorkSheetTemplate;

            public bool dataWasExtracted = false;
            
            public string fileLocation = "";
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
            double i = 0;
            double max = coll.Count();

            foreach (var element in coll)
            {
                i += 100 / max;

                if (i <= 100)
                {
                    backgroundWorker1.ReportProgress((int)i);
                }
                else
                {
                    i = 100;
                    backgroundWorker1.ReportProgress((int)i);
                }
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                foreach (var category in myCategory)
                {
                    if (category.Value.Name == currentName)
                    {
                        if ((element.Category != null))
                        {
                            if ((category.Value.Name == element.Category.Name))
                            {
                                found.Add(element);
                                ElementId typeID = element.GetTypeId();
                                Element E = localDoc.GetElement(typeID); 

                                fam = localDoc.GetElement(typeID) as FamilySymbol; 

                                if (fam != null)
                                {
                                    typeList.Add(new Tuple<string, Element>(fam.UniqueId, fam));
                                    sortedTypeParameters.Add(new Tuple<string, ElementId>(fam.UniqueId, fam.Id));
                                }
                                else
                                {
                                    try
                                    {
                                        typeList.Add(new Tuple<string, Element>(E.UniqueId, E));
                                        sortedTypeParameters.Add(new Tuple<string, ElementId>(E.UniqueId, E.Id));
                                    }
                                    catch
                                    {

                                    }

                                }
                            }
                        }
                    }
                }
            }

            foreach (Element item in found)
            {
                ElementId id = item.GetTypeId();
                ElementType type = localDoc.GetElement(id) as ElementType;

                item.UniqueId.OrderBy(x => x).Distinct().ToList();

                sortedInstanceParameters.Add(new Tuple<string, ElementId>(item.UniqueId, item.Id));

                foreach (Parameter a in item.Parameters)
                {
                    sortInstanceHeader.Add(a.Definition.Name);
                }
                try
                {
                    foreach (Parameter t in type.Parameters)
                    {
                        sortTypeHeader.Add(t.Definition.Name);
                    }
                }
                catch
                {

                }

            }

            // Collects names of all the headers
            sortInstanceHeader = sortInstanceHeader.Distinct().OrderBy(x => x).ToList();
            sortTypeHeader = sortTypeHeader.Distinct().OrderBy(x => x).ToList();

            // Collect parameters
            sortedInstanceParameters = sortedInstanceParameters.GroupBy(x => x.Item1).Select(y => y.First()).ToList();
            sortedTypeParameters = sortedTypeParameters.GroupBy(x => x.Item1).Select(y => y.First()).ToList();

            // Grab the list of types and orders them depending on the UniqueId
            typeList = typeList.GroupBy(x => x.Item1).Select(y => y.First()).ToList();

            foreach (var item in typeList)
            {
                sortedTypeElements.Add(item.Item2);
            }
            Export();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The Task Has Been Cancelled", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Variables.dataWasExtracted = false;
                progressBar1.Value = 0;
                ExportToExcel.Text = "Export to Excel";
                CloseButton.Enabled = true;
                listBox1.Enabled = true;
            }
            else if (e.Error != null)
            {
                MessageBox.Show("The Task Has Caused An Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Variables.dataWasExtracted = true;
                progressBar1.Value = 0;
                ExportToExcel.Text = "Export to Excel";
                CloseButton.Enabled = true;
                listBox1.Enabled = true;
            }
            else
            {
                MessageBox.Show("Category (" + currentName + ") parameter information has been exported to Excel.",
                    "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar1.Value = 0;
                
                Variables.xlWorkBookTemplate.Close(0);                
                Variables.dataWasExtracted = true;
                ExportToExcel.Text = "Export to Excel";
                CloseButton.Enabled = true;
                listBox1.Enabled = true;
            }

            ExportToExcel.Enabled = true;
            cancelButton.Enabled = false;
        }

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            ExportToExcel.Enabled = false;
            cancelButton.Enabled = true;
            ExportToExcel.Text = "Please wait...";
            CloseButton.Enabled = false;
            listBox1.Enabled = false;
            ClearLists();

            //FolderBrowserDialog browseFolders = new FolderBrowserDialog();
            SaveFileDialog browseFolders = new SaveFileDialog();
            browseFolders.Filter = "Excel files (*.xlsx)| *.xlsx";
            browseFolders.OverwritePrompt = false;
            DialogResult result = browseFolders.ShowDialog();


            //currentName = comboListCategory.Text;
            currentName = listBox1.Text;
            
            if (result == DialogResult.OK)
            {
                //folderLocation = browseFolders.SelectedPath;
                folderLocation = browseFolders.FileName;

                templateWorkBook = System.Reflection.Missing.Value;
                Variables.xlWorkBookTemplate = Variables.xlApp.Workbooks.Add(templateWorkBook);
                try
                {
                    run = true;

                    Variables.xlWorkBookTemplate.SaveAs(folderLocation);
                    ExportToExcel.Text = "Please wait...";
                }
                catch (Exception ex)
                {
                    Variables.xlWorkBookTemplate.Close(0);
                    result = DialogResult.Cancel;
                    ExportToExcel.Text = "Export to Excel";
                    CloseButton.Enabled = true;
                    listBox1.Enabled = true;
                    MessageBox.Show(ex.Message);
                }
            }
            if (result == DialogResult.Cancel)
            {
                //folderLocation = "";
                ExportToExcel.Enabled = true;
                cancelButton.Enabled = false;
                run = false;
                ExportToExcel.Text = "Export to Excel";
                CloseButton.Enabled = true;
                listBox1.Enabled = true;
            }

            if (run == true)
            {
                if (!this.backgroundWorker1.IsBusy)
                {
                    this.backgroundWorker1.RunWorkerAsync();
                }
            }
            run = false;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();       
            progressBar1.Value = 0;
            Variables.dataWasExtracted = false;
            Variables.xlWorkBookTemplate.Close(0);

            ExportToExcel.Text = "Export to Excel";
            CloseButton.Enabled = true;
            listBox1.Enabled = true;
        }

        private void GetListOfElements()
        {
            Categories cat = localDoc.Settings.Categories;
            myCategory = new SortedList<string, Category>();

            foreach (Category item in cat)
            {
                if (item.CategoryType.ToString() != "Model" || item.CanAddSubcategory == true)
                {
                    myCategory.Add(item.Name, item);
                }

            }
            foreach (var item in myCategory)
            {
                //comboListCategory.Items.Add(item.Value.Name);
                listBox1.Items.Add(item.Value.Name);
            }
        }

        private void CleanUpFiles()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (Variables.dataWasExtracted)
            {
                Marshal.ReleaseComObject(Variables.xlWorkSheetTemplate);
                Marshal.ReleaseComObject(Variables.worksheet);
                Marshal.ReleaseComObject(Variables.worksheet2);
                Marshal.ReleaseComObject(Variables.newWorksheet);
                Marshal.ReleaseComObject(Variables.xlWorkBookTemplate);

                Variables.xlWorkSheetTemplate = null;
                Variables.worksheet = null;
                Variables.worksheet2 = null;
                Variables.newWorksheet = null;
                Variables.xlWorkBookTemplate = null;
            }

            Variables.xlApp.Quit();
            Marshal.ReleaseComObject(Variables.xlApp);
        }

        private void ClearLists()
        {
            found.Clear();
            sortInstanceHeader.Clear();
            sortTypeHeader.Clear();
            typeList.Clear();
            sortedTypeElements.Clear();
            sortedInstanceParameters.Clear();
            sortedTypeParameters.Clear();
        }

        private void Export()
        {
            int column = 0, row = 0;
            string TypeS, TypeS2, InstanceS, InstanceS2;
            
            #region Type Parameter
            Variables.xlWorkSheetTemplate = (Excel.Worksheet)Variables.xlWorkBookTemplate.Worksheets.get_Item(1);

            Variables.xlWorkSheetTemplate.Cells[1, 1] = "GUID";
            Variables.xlWorkSheetTemplate.Cells[1, 2] = "Revit ID";

            for (int i = 1; i <= sortTypeHeader.Count; i++)
            {
                Variables.xlWorkSheetTemplate.Cells[1, i + 2] = sortTypeHeader[i - 1].ToString();
            }

            foreach (var item in sortedTypeParameters)
            {
                Variables.xlWorkSheetTemplate.Cells[row + 2, column + 1] = item.Item1;
                Variables.xlWorkSheetTemplate.Cells[row + 2, column + 2] = item.Item2.ToString();
                row++;
            }

            row = 0;
            
            foreach (Element item in sortedTypeElements)
            {
                foreach (Parameter P in item.Parameters)
                {
                    TypeS = P.AsValueString();
                    TypeS2 = P.AsString();
                    
                    for (int i = 0; i < sortTypeHeader.Count; i++)
                    {
                        if (sortTypeHeader[i] == P.Definition.Name)
                        {                            
                            if (TypeS != null)
                            {
                                Variables.xlWorkSheetTemplate.Cells[row + 2, i + 3] = TypeS;
                                break;
                            }
                            else
                            {
                                Variables.xlWorkSheetTemplate.Cells[row + 2, i + 3] = TypeS2;
                                break;
                            }
                        }
                    }
                }
                row++;
            }

            Variables.worksheet = (Excel.Worksheet)Variables.xlApp.Worksheets["Sheet1"];
            Variables.worksheet.Name = "Type Parameter";
            #endregion

            #region Instance Parameter
            column = 0;
            row = 0;
            
            Variables.newWorksheet = (Excel.Worksheet)Variables.xlWorkBookTemplate.Worksheets.Add(After: Variables.xlApp.Worksheets[Variables.xlApp.Worksheets.Count]);
            Variables.xlWorkSheetTemplate = (Excel.Worksheet)Variables.xlWorkBookTemplate.Worksheets.get_Item(2);

            Variables.xlWorkSheetTemplate.Cells[1, 1] = "GUID";
            Variables.xlWorkSheetTemplate.Cells[1, 2] = "Revit ID";

            for (int i = 1; i <= sortInstanceHeader.Count; i++)
            {
                Variables.xlWorkSheetTemplate.Cells[1, i + 2] = sortInstanceHeader[i - 1].ToString();
            }
            
            foreach (var item in sortedInstanceParameters)
            {
                Variables.xlWorkSheetTemplate.Cells[row + 2, column + 1] = item.Item1;
                Variables.xlWorkSheetTemplate.Cells[row + 2, column + 2] = item.Item2.ToString();
                row++;
            }

            row = 0;
            
            foreach (Element item in found)
            {
                foreach (Parameter a in item.Parameters)
                {
                    InstanceS = a.AsValueString();
                    InstanceS2 = a.AsString();
                    
                    for (int i = 0; i < sortInstanceHeader.Count; i++)
                    {
                        if (sortInstanceHeader[i].ToString() == a.Definition.Name)
                        {
                            if (InstanceS != null)
                            {
                                Variables.xlWorkSheetTemplate.Cells[row + 2, i + 3] = InstanceS;
                                break;
                            }
                            else
                            {
                                Variables.xlWorkSheetTemplate.Cells[row + 2, i + 3] = InstanceS2;
                                break;
                            }
                        }
                    }
                }
                row++;
            }

            Variables.worksheet2 = (Excel.Worksheet)Variables.xlApp.Worksheets["Sheet2"];
            Variables.worksheet2.Name = "Instance Parameter";
            #endregion

            try
            {
                Variables.xlWorkBookTemplate.Save();
                //MessageBox.Show("Template was Saved At" + folderLocation, "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Template was not saved", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportCategoryParametersForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            backgroundWorker1.CancelAsync();
            CleanUpFiles();
        }

        //private void comboListCategory_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if(comboListCategory.SelectedText != null)
        //    {
        //        ExportToExcel.Enabled = true;
        //    }
        //}

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ExportToExcel.Enabled == false)
            {
                ExportToExcel.Enabled = true;
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
            //CleanUpFiles();
            this.Close();
        }
    }
}