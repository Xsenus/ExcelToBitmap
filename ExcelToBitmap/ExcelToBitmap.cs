using System;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToBitmap
{
    public partial class ExcelToBitmap : Form
    {
        private static string pathExcel;
        private static string pathOut;
        
        private string pathDVD;
        private string pathCD;
        private string pathBluRay;

        public ExcelToBitmap()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog() { Filter = "Книга Excel (*.xlsx)|*.xlsx|Книга Excel 97-2003 (*.xls)|*.xls|All Files (*.*)|*.*" } )
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcel.Text = fileDialog.FileName;
                    pathExcel = fileDialog.FileName;
                }
            }
        }

        private void btnAddPathOut_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog browserDialog = new FolderBrowserDialog())
            {
                if (browserDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPathOut.Text = browserDialog.SelectedPath;
                    pathOut = browserDialog.SelectedPath;

                    pathDVD = $"{pathOut}\\DVD";
                    pathCD = $"{pathOut}\\CD";
                    pathBluRay = $"{pathOut}\\Blu-Ray";
                }
            }
        }

        private static OleDbConnection OleDbConnectionExcel { get; set; }

        private static bool ConnectionExcel(string path, string expansion = ".xlsx")
        {
            string connectionstring = string.Empty;
            try
            {
                if (expansion == ".xls")
                {
                    connectionstring = $"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = '{path}'; Extended Properties=\"Excel 8.0;HDR=YES;\"";
                }

                if (expansion == ".xlsx")
                {
                    connectionstring = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = '{path}'; Extended Properties=\"Excel 12.0;HDR=YES;\"";
                }

                OleDbConnectionExcel = new OleDbConnection(connectionstring);
                OleDbConnectionExcel.Open();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private string GetTableName(string strPath)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;

                theWorkbook = ExcelObj.Workbooks.Open($"{strPath}", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;

                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);

                var worksheetName = worksheet.Name;

                theWorkbook.Close(true);
                ExcelObj.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelObj);

                if (string.IsNullOrWhiteSpace(worksheetName))
                {
                    return default;
                }

                return worksheetName;
            }
            catch (Exception ex)
            {
                return null;
            }
        }       

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(pathExcel))
            {
                MessageBox.Show("Укажите путь к Excel файлу");
                txtExcel.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(pathOut))
            {
                MessageBox.Show("Укажите путь к конечной директории");
                txtPathOut.Focus();
                return;
            }

            if (!Directory.Exists(pathOut))
            {
                Directory.CreateDirectory(pathOut);                
            }

            if (!Directory.Exists(pathDVD))
            {
                Directory.CreateDirectory(pathDVD);
            }

            if (!Directory.Exists(pathCD))
            {
                Directory.CreateDirectory(pathCD);
            }

            if (!Directory.Exists(pathBluRay))
            {
                Directory.CreateDirectory(pathBluRay);
            }


            if (!ConnectionExcel(pathExcel))
            {
                return;
            }

            Task.Run(() => ImportExcelKazna(GetTableName(pathExcel)));
        }

        private async void ImportExcelKazna(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                return;
            }

            var sql = $"SELECT * FROM [{tableName}$]";
            var count = $"SELECT COUNT(*) FROM [{tableName}$]";

            using (var command = new OleDbCommand { Connection = OleDbConnectionExcel, CommandText = sql })
            {
                using (var reader = await command.ExecuteReaderAsync())
                {
                    var countCD = 0;
                    var countDVD = 0;
                    var countBluRay = 0;

                    while (await reader.ReadAsync())
                    {
                        if (reader[5].ToString().Contains("CD"))
                        {
                            Drawing(reader, pathCD);
                            countCD++;
                        }
                        else if (reader[5].ToString().Contains("DVD"))
                        {
                            Drawing(reader, pathDVD);
                            countDVD++;
                        }
                        else if (reader[5].ToString().Contains("Blu-ray"))
                        {
                            Drawing(reader, pathBluRay);
                            countBluRay++;
                        }
                    }
                }
            }

            MessageBox.Show(String.Format("Импорт окончен!"), "Окончание импорта", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void Drawing(DbDataReader reader, string path)
        {
            using (Bitmap bitmap = new Bitmap(3189, 2185))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Black);

                    graphics.DrawString(
                        reader[4].ToString(),
                        new Font("Verdana", (float)10),
                        new SolidBrush(Color.White), 15, 0);
                }
                bitmap.Save($"{path}\\{reader[0]}.JPEG", ImageFormat.Jpeg);
            }
        }
    }
}
