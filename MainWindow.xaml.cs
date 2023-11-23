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
        private string TxtFileCylinder { get; set; }
        private string TxtFileBundle { get; set; }
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
            if (Utils.CheckUserInputByValuesCylinderOrBundle(name_Sheet_Cylinder.Text, FilePath))
            {
                NameSheet = name_Sheet_Cylinder.Text;
                bool result = ConvertExcelFilesToTxt(FilePath, NameSheet, 1);
                if (result && !string.IsNullOrEmpty(TxtFileBundle)) ChangeImageOk(result, 4);
                else if (result == false && !string.IsNullOrEmpty(TxtFileBundle)) ChangeImageOk(result, 5);
                else ChangeImageOk(result, 1);
            }
            else
            {
                MessageBox.Show("empty name cylinder or empty file excel (button Open file)");
                ChangeImageOk(false, 1);
            }
            name_Sheet_Cylinder.Clear();
        }

        private void ChangeImageOk(bool result, int cylinderOrBundleOrClearButtonFileSave)
        {
            switch (cylinderOrBundleOrClearButtonFileSave)
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
                case 3:
                    if (result == true) image_ok_save_file.Visibility = Visibility.Visible;
                    else image_ok_save_file.Visibility = Visibility.Hidden;
                    image_ok_open_file.Visibility = Visibility.Hidden;
                    image_ok_cylinders.Visibility = Visibility.Hidden;
                    image_ok_bundles.Visibility = Visibility.Hidden;
                    FilePath = "";
                    TxtFileCylinder = "";
                    TxtFileBundle = "";
                    NameSheet = "";
                    NoCompleteFieldsCylinders = "";
                    NoCompleteFieldsBundles = "";
                    break;
                case 4:
                    image_ok_cylinders.Visibility = Visibility.Visible;
                    image_ok_bundles.Visibility = Visibility.Visible;
                    break;
                case 5:
                    image_ok_cylinders.Visibility = Visibility.Hidden;
                    image_ok_bundles.Visibility = Visibility.Visible;
                    break;
                case 6:
                    image_ok_cylinders.Visibility = Visibility.Visible;
                    image_ok_bundles.Visibility = Visibility.Hidden;
                    break;
            }

        }

        private void Button_bundle(object sender, RoutedEventArgs e)
        {
            if (Utils.CheckUserInputByValuesCylinderOrBundle(name_Sheet_Bundle.Text, FilePath))
            {
                NameSheet = name_Sheet_Bundle.Text;
                bool result = ConvertExcelFilesToTxt(FilePath, NameSheet, 2);
                if (result && !string.IsNullOrEmpty(TxtFileCylinder)) ChangeImageOk(result, 4);
                else if (result == false && !string.IsNullOrEmpty(TxtFileCylinder)) ChangeImageOk(result, 6);
                else ChangeImageOk(result, 2);
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
            bool correctCylinder = !string.IsNullOrEmpty(TxtFileCylinder) && !string.IsNullOrEmpty(FilePath);
            bool correctBundle = !string.IsNullOrEmpty(TxtFileBundle) && !string.IsNullOrEmpty(FilePath);
            string pathCylinder = "";
            string pathBundle = "";
            // Configure save file dialog box
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.Filter = "Text documents (.txt)|*.txt"; // Filter files by extension

            if (correctCylinder && correctBundle)
            {
                dialog.FileName = "Cylinders;Bundles";

                // Process save file dialog box results
                bool result = (bool)dialog.ShowDialog();

                string[] namesAndPaths = Utils.GetCorrectNameAndPath(dialog.FileName, 2);

                if (namesAndPaths.Length == 2)
                {
                    pathCylinder = namesAndPaths[0];
                    pathBundle = namesAndPaths[1];
                }

                    if (result == true)
                {
                    // Save document

                    using (StreamWriter sw = File.CreateText(pathCylinder + ".txt"))
                    {
                        sw.WriteLine(TxtFileCylinder);
                        File.WriteAllText(pathBundle, TxtFileBundle, Encoding.Unicode);
                    }

                }

                CheckIsNoCompleteFieldsCylinderEmpty(pathCylinder);

                CheckIsNoCompleteFieldsBundlesEmpty(pathBundle);
                
                ChangeImageOk(result, 3);
                
            }
            else if (correctCylinder || correctBundle)
            {
                dialog.FileName = "Cylinders / Bundles";

                // Process save file dialog box results
                bool result = (bool)dialog.ShowDialog();

                string[] namesAndPaths = Utils.GetCorrectNameAndPath(dialog.FileName, 1);

                if (namesAndPaths.Length == 2 && correctCylinder)
                {
                    pathCylinder = namesAndPaths[0];
                }
                else if (namesAndPaths.Length == 2 && correctBundle)
                {
                    pathBundle = namesAndPaths[0];
                }
                
                if (result == true)
                {
                    if (correctCylinder)
                    using (StreamWriter sw = File.CreateText(pathCylinder))
                    {
                        sw.WriteLine(TxtFileCylinder);
                    }
                    else
                        using (StreamWriter sw = File.CreateText(pathBundle))
                        {
                            sw.WriteLine(TxtFileBundle);
                        }
                }

                CheckIsNoCompleteFieldsCylinderEmpty(pathCylinder);

                CheckIsNoCompleteFieldsBundlesEmpty(pathBundle);

                ChangeImageOk(result, 3);
            }
            else
            {
                MessageBox.Show("empty name bundle or empty name cylinder or empty file excel (button Open file)");
            }


        }

        private void CheckIsNoCompleteFieldsBundlesEmpty(string pathBundle)
        {
            int index2 = pathBundle.LastIndexOf('\\');
            string path2 = pathBundle.Substring(0, index2 + 1);
            if (!string.IsNullOrEmpty(NoCompleteFieldsBundles)) CreateTxtFile(NoCompleteFieldsBundles, $"{path2}{"Info empty fields bundles"}");
        }

        private void CheckIsNoCompleteFieldsCylinderEmpty(string pathCylinder)
        {
            int index = pathCylinder.LastIndexOf('\\');
            string path = pathCylinder.Substring(0, index + 1);
            if (!string.IsNullOrEmpty(NoCompleteFieldsCylinders)) CreateTxtFile(NoCompleteFieldsCylinders, $"{path}{"Info empty fields cylinders"}");
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
                            if (sheet.SheetName.ToLower().Contains(nameSheet.ToLower()) && nameSheet.ToLower().Contains("cyli")) // cylinders
                            {
                                correctData = true;
                                DownloadDataCylinder dataCylinder = new DownloadDataCylinder("cylinder");
                                TxtFileCylinder = dataCylinder.CompleteCylinderData(sheet);
                                if (TxtFileCylinder == "") { MessageBox.Show("no created txt File -> empty spreadsheet cylinder !!!"); return false; }
                                else
                                {
                                    NoCompleteFieldsCylinders = dataCylinder.CheckFieldsAreComplete();
                                }
                            }
                            break;
                        case 2:
                            if (sheet.SheetName.ToLower().Contains(nameSheet.ToLower()) && nameSheet.ToLower().Contains("bund")) // bundles
                            {
                                correctData = true;
                                DownloadDataBundle dataBundle = new DownloadDataBundle("bundle");
                                TxtFileBundle = dataBundle.CompleteBundleData(sheet);
                                if (TxtFileBundle == "") { MessageBox.Show("no created txt File -> empty spreadsheet bundle !!!"); return false; }
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
                correctData=false;
            }

            if (correctData == false && choiceUser == 1) MessageBox.Show("Wrong name cylinder (enter at least 4 chars: cyli - cylinder ) or invalid excel file or you clicked button 'X', please again !!!");
            else if (correctData == false && choiceUser == 2) MessageBox.Show("Wrong name bundle (enter at least 4 chars: bund - bundle ) or invalid excel file or you clicked button 'X', please again !!!");
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
