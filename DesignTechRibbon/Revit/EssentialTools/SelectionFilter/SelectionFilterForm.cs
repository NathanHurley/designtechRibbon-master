using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DesignTechRibbonPaid.Revit.EssentialTools.SelectionFilter
{
    public partial class SelectionFilterForm : System.Windows.Forms.Form
    {
        Document localDoc;
        UIDocument localUidoc;
        List<string> categoryList = new List<string>();
        List<Tuple<string, Category>> typeList = new List<Tuple<string, Category>>();
        Element E;
        List<string> S = new List<string>();


        public SelectionFilterForm(Document doc, UIDocument uidoc)
        {
            InitializeComponent();

            localDoc = doc;
            localUidoc = uidoc;

            GetListOfSelectedCategories();
        }

        private void GetListOfSelectedCategories()
        {
            Selection selection = localUidoc.Selection;
            ICollection<ElementId> selectedIds = localUidoc.Selection.GetElementIds();

            foreach (ElementId item in selectedIds)
            {
                E = localDoc.GetElement(item); // Converts ids into elements
                categoryList.Add(E.Category.Name);
                typeList.Add(new Tuple<string, Category>(E.Category.Name, E.Category));
            }

            categoryList = categoryList.Distinct().OrderBy(x => x).ToList();

            for (int i = 0; i < categoryList.Count; i++)
            {
                checkedListBox1.Items.Add(categoryList[i]);
                checkedListBox1.SetItemChecked(i, true);
            }

            foreach (Tuple<string, Category> item in typeList)
            {
                S.Add(item.Item1);
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                for (int j = 0; j < typeList.Count; j++)
                {
                    if (checkedListBox1.GetItemChecked(i) && S[j].Equals((string)checkedListBox1.Items[i]))
                    {
                        checkedListBox2.Items.Add(S[j]);
                    }
                }
            }


            //for (int i = 0; i < checkedListBox1.Items.Count; i++)
            //{
            //    if (checkedListBox1.GetItemChecked(i))
            //    {
            //        foreach (var item in typeList)
            //        {
            //            if (item.Item1 == (string)checkedListBox1.Items[i])
            //            {
            //                checkedListBox2.Items.Add(item.Item1);
            //            }
            //        }
            //    }
            //}

            //FilteredElementCollector coll = new FilteredElementCollector(localDoc).WhereElementIsNotElementType();



            //Categories cat = localDoc.Settings.Categories;
            //SortedList<string, Category> myCategory = new SortedList<string, Category>();

            //foreach (Category item in cat)
            //{
            //    myCategory.Add(item.Name, item);
            //}            
        }
    }
}