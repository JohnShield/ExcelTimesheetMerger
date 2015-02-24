/* Author: John Shield
 * Date: 2014 October
 */

using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Windows;
using Microsoft.Win32;
using XlsxReadWrite.Properties;
using System.Diagnostics;
using System.Windows.Controls;

namespace XlsxReadWrite
{
    public partial class MainWindow : Window
    {
        private readonly DataTable data = new DataTable(Settings.Default.DataTableName, Settings.Default.DataTableNamespace);
        private readonly string tempDir = Settings.Default.TemporaryDirectory;
        private int TabControlLastSelection = 0;

        public MainWindow()
        {
            InitializeComponent();

            this.data.Columns.Add(new DataColumn("A"));
            this.data.Columns.Add(new DataColumn("B"));
            this.data.Columns.Add(new DataColumn("C"));
            this.data.Columns.Add(new DataColumn("D"));
            this.data.Columns.Add(new DataColumn("E"));
            this.data.Columns.Add(new DataColumn("F"));
            this.data.Columns.Add(new DataColumn("G"));
            this.data.Columns.Add(new DataColumn("H"));
            this.data.Columns.Add(new DataColumn("I"));
            this.data.Columns.Add(new DataColumn("J"));
            this.data.Columns.Add(new DataColumn("K"));
            this.data.Columns.Add(new DataColumn("L"));
            this.data.Columns.Add(new DataColumn("M"));
            this.data.Columns.Add(new DataColumn("N"));
            this.data.Columns.Add(new DataColumn("O"));
            this.data.Columns.Add(new DataColumn("P"));
            this.data.Columns.Add(new DataColumn("Q"));
            this.data.Columns.Add(new DataColumn("R"));
            this.data.Columns.Add(new DataColumn("S"));
            this.data.Columns.Add(new DataColumn("T"));
            this.data.Columns.Add(new DataColumn("U"));
            this.data.Columns.Add(new DataColumn("V"));
            this.data.Columns.Add(new DataColumn("W"));
            this.data.Columns.Add(new DataColumn("X"));
            this.data.Columns.Add(new DataColumn("Y"));
            this.data.Columns.Add(new DataColumn("Z"));


            this.dataGrid.ItemsSource = this.data.AsDataView();

        }

        private void ShowOpenFileDialog(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                FileName = Settings.Default.InputFileName,
                DefaultExt = "*.xlsx",
                Filter = "Excel Workbook (.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog(this) == true)
                Settings.Default.InputFileName = openFileDialog.FileName;
        }

        private void ShowOpenFileDialog2(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                FileName = Settings.Default.InputFileName2,
                DefaultExt = "*.xlsx",
                Filter = "Excel Workbook (.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog(this) == true)
                Settings.Default.InputFileName2 = openFileDialog.FileName;
        }

        private void ShowSaveFileDialog(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog()
            {
                FileName = Settings.Default.OutputFileName,
                DefaultExt = "*.xlsx",
                Filter = "Excel Workbook (.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog(this) == true)
                Settings.Default.OutputFileName = saveFileDialog.FileName;
        }

        private void ReadInput(object sender, RoutedEventArgs e)
        {
            ReadInput_sub();
        }

        private bool ReadInput_sub() {

            // Get the input file name from the text box.
            var fileName = this.inputTextBox.Text;

            // Delete contents of the temporary directory.
            XlsxRW.DeleteDirectoryContents(tempDir);

            if (!(File.Exists(fileName))) {
                System.Windows.Forms.MessageBox.Show("Invalid File Chosen (State): " + fileName,
                            "Invalid File",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation,
                            System.Windows.Forms.MessageBoxDefaultButton.Button1);
                return false;
            }

            // Unzip input XLSX file to the temporary directory.
            bool ret = XlsxRW.UnzipFile(fileName, tempDir);
            if (ret == false)
                return false;

            IList<string> stringTable;
            // Open XML file with table of all unique strings used in the workbook..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"),
                FileMode.Open, FileAccess.Read))
                // ..and call helper method that parses that XML and returns an array of strings.
                stringTable = XlsxRW.ReadStringTable(stream);

            // Open XML file with worksheet data..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\worksheets\sheet1.xml"),
                FileMode.Open, FileAccess.Read))
                // ..and call helper method that parses that XML and fills DataTable with values.
                XlsxRW.ReadWorksheet(stream, stringTable, this.data);

            return true;
        }

        private void ReadInput2(object sender, RoutedEventArgs e)
        {
            ReadInput2_sub();
        }

        private bool ReadInput2_sub()
        {
        
            // Get the input file name from the text box.
            var fileName = this.inputTextBox2.Text;

            // Delete contents of the temporary directory.
            XlsxRW.DeleteDirectoryContents(tempDir);

            if (!(File.Exists(fileName)))
            {
                System.Windows.Forms.MessageBox.Show("Invalid File Chosen (Divison): " + fileName,
                            "Invalid File",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation,
                            System.Windows.Forms.MessageBoxDefaultButton.Button1);
                return false;
            }

            // Unzip input XLSX file to the temporary directory.
            bool ret = XlsxRW.UnzipFile(fileName, tempDir);
            if (ret == false)
                return false;

            IList<string> stringTable;
            // Open XML file with table of all unique strings used in the workbook..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"),
                FileMode.Open, FileAccess.Read))
                // ..and call helper method that parses that XML and returns an array of strings.
                stringTable = XlsxRW.ReadStringTable(stream);

            // Open XML file with worksheet data..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\worksheets\sheet1.xml"),
                FileMode.Open, FileAccess.Read))
                // ..and call helper method that parses that XML and fills DataTable with values.
                XlsxRW.ReadWorksheet2(stream, stringTable, this.data);
            return true;
        }

        private void WriteOutput(object sender, RoutedEventArgs e)
        {
            // Get the output file name from the text box.
            string fileName = this.outputTextBox.Text;

            if (ReadInput2_sub() == false)
                return;

            if (ReadInput_sub() == false)
                return;

            // Delete contents of the temporary directory.
            XlsxRW.DeleteDirectoryContents(tempDir);

            // Unzip AMC Roster XLSX file to the temporary directory.
            bool ret = XlsxRW.UnzipFile(this.inputTextBox2.Text, tempDir);
            if (ret == false)
                return;
            // We will need two string tables; a lookup IDictionary<string, int> for fast searching and 
            // an ordinary IList<string> where items are sorted by their index.
            IDictionary<string, int> lookupTable;

            // Call helper methods which creates both tables from input data.
            var stringTable = XlsxRW.CreateStringTables(this.data, out lookupTable);

            // Create XML file..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"),
                FileMode.Create))
                // ..and fill it with unique strings used in the workbook
                XlsxRW.WriteStringTable(stream, stringTable);

            // Create XML file..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\worksheets\sheet1.xml"),
                FileMode.Create))
                // ..and fill it with rows and columns of the DataTable.
                XlsxRW.WriteWorksheet(stream, this.data, lookupTable);

            // ZIP temporary directory to the XLSX file.
            ret = XlsxRW.ZipDirectory(tempDir, fileName);
            if (ret == false)
                return;

            // If checkbox is checked, show XLSX file in Microsoft Excel.
            if (this.openFileCheckBox.IsChecked == true)
                System.Diagnostics.Process.Start(fileName);
        }

        /*************************************/
        /*Settings Controls*/
        /*************************************/

        private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (TabItem2.IsSelected)
            {
                LoadTabbedSettings();
                TabControlLastSelection = 2;
            }
            else
            {
                if (TabControlLastSelection == 2)
                {
                    SaveTabbedSettings();
                }
                TabControlLastSelection = 0;
            }
        }

        // This function allows for the saving of modified settings via a visual interface
        private void SaveTabbedSettings()
        {

            if (DockPanelRemoveFilter == null) {
                Settings.Default.RemoveMatchArray_str = "";
                Settings.Default.RemoveMatchArray_col = "";
                Settings.Default.RemoveMatchArray_exact = "";
            }
            else {
                string[] filter_R_string = new string[DockPanelRemoveFilter.Children.Count];
                int[] filter_R_col = new int[DockPanelRemoveFilter.Children.Count];
                bool[] filter_R_matchtype = new bool[DockPanelRemoveFilter.Children.Count];

                for (int ii = 0; ii < DockPanelRemoveFilter.Children.Count; ii++)
                {
                    //NOTE ORDER OF ADDED COMPONENTS IS IMPORTANT
                    //This sets the array locations for the checkbox and the two text_inputs
                    DockPanel childPanel = (DockPanel)DockPanelRemoveFilter.Children[ii];
                    filter_R_matchtype[ii] = (bool)((CheckBox)childPanel.Children[0]).IsChecked;
                    
                    bool result = System.Int32.TryParse(((TextBox)childPanel.Children[1]).Text, out filter_R_col[ii]);
                    if (!result)
                        filter_R_col[ii] = 0;

                    filter_R_string[ii] = ((TextBox)childPanel.Children[2]).Text;
                }

                Settings.Default.RemoveMatchArray_str = string.Join(",", filter_R_string);
                Settings.Default.RemoveMatchArray_col = string.Join(",", filter_R_col);
                Settings.Default.RemoveMatchArray_exact = string.Join(",", filter_R_matchtype);
            }


            if (DockPanelIncludeFilter == null)
            {
                Settings.Default.IncludeMatchArray_str = "";
                Settings.Default.IncludeMatchArray_col = "";
                Settings.Default.IncludeMatchArray_exact = "";
            }
            else
            {
                string[] filter_A_string = new string[DockPanelIncludeFilter.Children.Count];
                int[] filter_A_col = new int[DockPanelIncludeFilter.Children.Count];
                bool[] filter_A_matchtype = new bool[DockPanelIncludeFilter.Children.Count];

                for (int ii = 0; ii < DockPanelIncludeFilter.Children.Count; ii++)
                {
                    //NOTE ORDER OF ADDED COMPONENTS IS IMPORTANT
                    //This sets the array locations for the checkbox and the two text_inputs
                    DockPanel childPanel = (DockPanel)DockPanelIncludeFilter.Children[ii];
                    filter_A_matchtype[ii] = (bool)((CheckBox)childPanel.Children[0]).IsChecked;
                    bool result = System.Int32.TryParse(((TextBox)childPanel.Children[1]).Text, out filter_A_col[ii]);
                    if (!result)
                        filter_A_col[ii] = 0;
                    filter_A_string[ii] = ((TextBox)childPanel.Children[2]).Text;
                }

                Settings.Default.IncludeMatchArray_str = string.Join(",", filter_A_string);
                Settings.Default.IncludeMatchArray_col = string.Join(",", filter_A_col);
                Settings.Default.IncludeMatchArray_exact = string.Join(",", filter_A_matchtype);
            }
            Debug.WriteLine("SaveTabbedSettings");
        }

        // This function loads the saved settings from the configuration file into the tabbed visual interface
        private void LoadTabbedSettings()
        {
            if (!string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_col) && !string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_exact)
                && !string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_str))
            {
                string[] filter_R_string = Settings.Default.RemoveMatchArray_str.Split(',');
                int[] filter_R_col = System.Array.ConvertAll(Settings.Default.RemoveMatchArray_col.Split(','), s => int.Parse(s));
                bool[] filter_R_matchtype = System.Array.ConvertAll(Settings.Default.RemoveMatchArray_exact.Split(','), s => bool.Parse(s));


                // clear previous data
                DockPanelRemoveFilter.Children.Clear();
                DockPanelIncludeFilter.Children.Clear();

                // loop for the number of filters
                for (int ii = 0; ii < filter_R_string.Length && ii < filter_R_col.Length; ii++)
                {
                    if (string.IsNullOrEmpty(filter_R_string[ii]) == false && filter_R_col[ii] > 0)
                    {
                        // create the filter dockpanel instance
                        DockPanel filterdockpanel = new DockPanel();

                        filterdockpanel.Height = 28;
                        filterdockpanel.VerticalAlignment = VerticalAlignment.Top;

                        //NOTE ORDER OF ADDING COMPONENTS IS IMPORTANT
                        //This sets the array locations for the checkbox and the two text_inputs

                        // create tickbox
                        CheckBox exactmatch = new CheckBox();
                        exactmatch.IsChecked = filter_R_matchtype[ii];
                        exactmatch.Width = 17;
                        DockPanel.SetDock(exactmatch, Dock.Right);
                        filterdockpanel.Children.Add(exactmatch);

                        // set the column entry
                        TextBox columnEntry = new TextBox();
                        columnEntry.Text = filter_R_col[ii].ToString();
                        columnEntry.Width = 34;
                        DockPanel.SetDock(columnEntry, Dock.Right);
                        filterdockpanel.Children.Add(columnEntry);

                        // set the string entry
                        TextBox stringEntry = new TextBox();
                        stringEntry.Text = filter_R_string[ii];
                        DockPanel.SetDock(stringEntry, Dock.Left);
                        filterdockpanel.Children.Add(stringEntry);

                        // Add filter to the system
                        DockPanel.SetDock(filterdockpanel, Dock.Top);
                        DockPanelRemoveFilter.Children.Add(filterdockpanel);
                    }
                }
            }
            
            if (!string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_col) && !string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_exact)
                && !string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_str)) {
                string [] filter_A_string = Settings.Default.IncludeMatchArray_str.Split(',');
                int [] filter_A_col = System.Array.ConvertAll(Settings.Default.IncludeMatchArray_col.Split(','), s => int.Parse(s));
                bool[] filter_A_matchtype = System.Array.ConvertAll(Settings.Default.IncludeMatchArray_exact.Split(','), s => bool.Parse(s));

                // loop for the number of filters
                for (int ii = 0; ii < filter_A_string.Length && ii < filter_A_col.Length; ii++)
                {
                    if (string.IsNullOrEmpty(filter_A_string[ii]) == false && filter_A_col[ii] > 0)
                    {
                        // create the filter dockpanel instance
                        DockPanel filterdockpanel = new DockPanel();

                        filterdockpanel.Height = 28;
                        filterdockpanel.VerticalAlignment = VerticalAlignment.Top;

                        // create tickbox
                        CheckBox exactmatch = new CheckBox();
                        exactmatch.IsChecked = filter_A_matchtype[ii];
                        exactmatch.Width = 17;
                        DockPanel.SetDock(exactmatch, Dock.Right);
                        filterdockpanel.Children.Add(exactmatch);

                        // set the column entry
                        TextBox columnEntry = new TextBox();
                        columnEntry.Text = filter_A_col[ii].ToString();
                        columnEntry.Width = 34;
                        DockPanel.SetDock(columnEntry, Dock.Right);
                        filterdockpanel.Children.Add(columnEntry);

                        // set the string entry
                        TextBox stringEntry = new TextBox();
                        stringEntry.Text = filter_A_string[ii];
                        DockPanel.SetDock(stringEntry, Dock.Left);
                        filterdockpanel.Children.Add(stringEntry);

                        // Add filter to the system
                        DockPanel.SetDock(filterdockpanel, Dock.Top);
                        DockPanelIncludeFilter.Children.Add(filterdockpanel);
                    }
                }
            }
            Debug.WriteLine("LoadTabbedSettings");
        }

        
        private void AddRemoveFilter(object sender, RoutedEventArgs e)
        {
            // create the filter dockpanel instance
            DockPanel filterdockpanel = new DockPanel();

            filterdockpanel.Height = 28;
            filterdockpanel.VerticalAlignment = VerticalAlignment.Top;

            // create tickbox
            CheckBox exactmatch = new CheckBox();
            exactmatch.Width = 17;
            DockPanel.SetDock(exactmatch, Dock.Right);
            filterdockpanel.Children.Add(exactmatch);

            // set the column entry
            TextBox columnEntry = new TextBox();
            columnEntry.Width = 34;
            DockPanel.SetDock(columnEntry, Dock.Right);
            filterdockpanel.Children.Add(columnEntry);

            // set the string entry
            TextBox stringEntry = new TextBox();
            DockPanel.SetDock(stringEntry, Dock.Left);
            filterdockpanel.Children.Add(stringEntry);

            // Add filter to the system
            DockPanel.SetDock(filterdockpanel, Dock.Top);
            DockPanelRemoveFilter.Children.Add(filterdockpanel);
        }

        private void AddIncludeFilter(object sender, RoutedEventArgs e)
        {
            // create the filter dockpanel instance
            DockPanel filterdockpanel = new DockPanel();

            filterdockpanel.Height = 28;
            filterdockpanel.VerticalAlignment = VerticalAlignment.Top;

            // create tickbox
            CheckBox exactmatch = new CheckBox();
            exactmatch.Width = 17;
            DockPanel.SetDock(exactmatch, Dock.Right);
            filterdockpanel.Children.Add(exactmatch);

            // set the column entry
            TextBox columnEntry = new TextBox();
            columnEntry.Width = 34;
            DockPanel.SetDock(columnEntry, Dock.Right);
            filterdockpanel.Children.Add(columnEntry);

            // set the string entry
            TextBox stringEntry = new TextBox();
            DockPanel.SetDock(stringEntry, Dock.Left);
            filterdockpanel.Children.Add(stringEntry);

            // Add filter to the system
            DockPanel.SetDock(filterdockpanel, Dock.Top);
            DockPanelIncludeFilter.Children.Add(filterdockpanel);
        }

        // Button to save settings information
        private void SettingsSaveButton(object sender, RoutedEventArgs e)
        {
            SaveTabbedSettings();
        }

    }
}
