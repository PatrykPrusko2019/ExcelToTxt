using ExcelToTxt.GetBundles;
using ExcelToTxt.GetCylinders;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace ExcelToTxt
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string FilePath { get; set; }
        private string NameSheet { get; set; }
        private string TxtFile { get; set; }
        private string NoCompleteFieldsCylinders { get; set; }
        private string NoCompleteFieldsBundles { get; set; }


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Open_files(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files xlsx (.xlsx)|*.xlsx|Excel files xls (.xls)|*.xls";
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                image_ok_open_file.Visibility = Visibility.Visible;
                image_ok_save_file.Visibility = Visibility.Hidden;
            }
            else
            {
                image_ok_open_file.Visibility = Visibility.Hidden;
            }

        }

        private void Button_cylinder(object sender, RoutedEventArgs e)
        {
            if (name_Sheet_Cylinder.Text != "" && !string.IsNullOrEmpty(FilePath))
            {
                NameSheet = name_Sheet_Cylinder.Text;
                bool result = ConvertExcelFilesToTxt(FilePath, NameSheet, 1);
                ChangeImageOk(result, 1);
            }
            else
            {
                MessageBox.Show("empty name cylinder or empty file excel (button Open file)");
                ChangeImageOk(false, 1);
            }
            name_Sheet_Cylinder.Clear();
        }

        private void ChangeImageOk(bool result, int cylinderOrBundle)
        {
            switch (cylinderOrBundle)
            {
                case 1:
                    if (result)
                    {
                        image_ok_cylinders.Visibility = Visibility.Visible;
                        image_ok_bundles.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        image_ok_cylinders.Visibility = Visibility.Hidden;
                        image_ok_bundles.Visibility = Visibility.Hidden;
                    }
                    break;
                case 2:
                    if (result)
                    {
                        image_ok_bundles.Visibility = Visibility.Visible;
                        image_ok_cylinders.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        image_ok_cylinders.Visibility = Visibility.Hidden;
                        image_ok_bundles.Visibility = Visibility.Hidden;
                    }
                    break;
            }

        }

        private void Button_bundle(object sender, RoutedEventArgs e)
        {
            if (name_Sheet_Bundle.Text != "" && !string.IsNullOrEmpty(FilePath))
            {
                NameSheet = name_Sheet_Bundle.Text;
                bool result = ConvertExcelFilesToTxt(FilePath, NameSheet, 2);
                ChangeImageOk(result, 2);
            }
            else
            {
                MessageBox.Show("empty name bundle or empty file excel (button Open file)");
                ChangeImageOk(false, 2);
            }
            name_Sheet_Bundle.Clear();
        }

        private void Button_Save_file(object sender, RoutedEventArgs e)
        {
            string filename = "";
            if (!string.IsNullOrEmpty(TxtFile) && !string.IsNullOrEmpty(FilePath))
            {
                // Configure save file dialog box
                var dialog = new Microsoft.Win32.SaveFileDialog();
                dialog.FileName = "Cylinders / Bundles"; // Default file name

                // dialog.Filter = "Text documents (.txt)|*.txt"; // Filter files by extension

                // Show save file dialog box
                bool? result = dialog.ShowDialog();

                // Process save file dialog box results
                if (result == true)
                {
                    // Save document
                    filename = dialog.FileName;

                    File.WriteAllText(filename, TxtFile, Encoding.Unicode);

                    //using (StreamWriter sw = File.CreateText(filename))
                    //{
                    //    sw.WriteLine(TxtFile);
                    //}
                }

                int index = filename.LastIndexOf('\\');
                string path = filename.Substring(0, index + 1);
                if (!string.IsNullOrEmpty(NoCompleteFieldsCylinders)) CreateTxtFile(NoCompleteFieldsCylinders, $"{path}{"Info empty fields cylinders"}");
                if (!string.IsNullOrEmpty(NoCompleteFieldsBundles)) CreateTxtFile(NoCompleteFieldsBundles, $"{path}{"Info empty fields bundles"}");

                image_ok_save_file.Visibility = Visibility.Visible;
                image_ok_open_file.Visibility = Visibility.Hidden;
                image_ok_cylinders.Visibility = Visibility.Hidden;
                image_ok_bundles.Visibility = Visibility.Hidden;
                FilePath = "";
                TxtFile = "";
                NameSheet = "";
                NoCompleteFieldsCylinders = "";
                NoCompleteFieldsBundles = "";
            }
            else
            {
                MessageBox.Show("empty name bundle or empty name cylinder or empty file excel (button Open file)");
            }


        }

        private void Button_Exit(object sender, RoutedEventArgs e)
        {
            Environment.Exit(1);
        }


        private bool ConvertExcelFilesToTxt(string filePath, string nameSheet, int choiceUser)
        {
            bool correctData = false;
            try
            {
                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                if (workbook == null)
                {
                    MessageBox.Show("file is no excel format, please again !!!");
                    return correctData;
                }

                foreach (var sheet in workbook)
                {
                    switch (choiceUser)
                    {
                        case 1:
                            if (sheet.SheetName.ToLower().Contains(nameSheet.ToLower())) // cylinders
                            {
                                correctData = true;
                                DownloadDataCylinder dataCylinder = new DownloadDataCylinder("cylinder");
                                TxtFile = dataCylinder.CompleteCylinderData(sheet);
                                if (TxtFile == "") { MessageBox.Show("no created txt File -> empty spreadsheet cylinder !!!"); return false; }
                                else
                                {
                                    NoCompleteFieldsCylinders = dataCylinder.CheckFieldsAreComplete();
                                }
                            }
                            break;
                        case 2:
                            if (sheet.SheetName.ToLower().Contains(nameSheet.ToLower())) // bundles
                            {
                                correctData = true;
                                DownloadDataBundle dataBundle = new DownloadDataBundle("bundle");
                                TxtFile = dataBundle.CompleteBundleData(sheet);
                                if (TxtFile == "") { MessageBox.Show("no created txt File -> empty spreadsheet bundle !!!"); return false; }
                                else
                                {
                                    NoCompleteFieldsBundles = dataBundle.CheckFieldsAreComplete();
                                }
                            }
                            break;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (correctData == false) MessageBox.Show("Wrong name cylinder or name bundle, please again !!!");
            return correctData;
        }

        private void CreateTxtFile(string txtFile, string path)
        {
            using (StreamWriter sw = File.CreateText($"{path}.txt"))
            {
                sw.WriteLine(txtFile);
            }
        }
    }
}
