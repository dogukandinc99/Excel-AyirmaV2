using OfficeOpenXml;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;      // Dosya okuma/yazma işlemleri için
using System;         // Genel .NET sınıfları için
using System.Windows.Forms;
using System.Diagnostics;
using System.Data; // Hata mesajlarını göstermek için (isteğe bağlı)

namespace ExcelAyirmaV2
{
    public class ExcelIslemler
    {
        _Excel.Application excel = new _Excel.Application();
        _Excel.Workbook _workbook;
        _Excel.Worksheet _worksheet;

        System.Data.DataTable dataTableList = new System.Data.DataTable("Excel-List");


        int columncontrolnumber = 8;
        int rowsCount = 0, columnsCount = 0;
        ProgressBar progress;
        Label label1;

        // Gelen adresdeki excel dosyasını açar ve 1. sayfa seçilir.


        public ExcelIslemler()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public string ConvertXlsToXlsxInPlace(string xlsFilePath)
        {
            try
            {
                if (!File.Exists(xlsFilePath))
                {
                    throw new FileNotFoundException("Dosya bulunamadı.", xlsFilePath);
                }

                _workbook = excel.Workbooks.Open(xlsFilePath);

                // Yeni dosya yolu
                string newFilePath = Path.ChangeExtension(xlsFilePath, ".xlsx");

                // Kaydet ve kapat
                _workbook.SaveAs(newFilePath, _Excel.XlFileFormat.xlOpenXMLWorkbook);
                _workbook.Close(false);
                excel.Quit();

                // Excel işlemini temizle
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                // Eski dosyayı sil ve geçici dosyayı yeniden adlandır
                File.Delete(xlsFilePath);

                // Yeni dosya yolunu döner
                return newFilePath; // Yeni dosya adresi ve adı
            }
            catch (Exception ex)
            {
                throw new Exception($"Dönüştürme sırasında hata oluştu: {ex.Message}");
            }
        }

        private ExcelPackage package;
        private ExcelWorksheet worksheet;

        public void excelOpen(string path)
        {
            try
            {
                package = new ExcelPackage(new FileInfo(path));
                worksheet = package.Workbook.Worksheets[0]; // İlk sayfayı seçer
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Dosya açılırken bir sorun oluştu.\nHata Kodu: {ex.Message}");
            }
        }

        // Sürelerde sıfır yazanları 1 e dönüştürür.
        public void zeroChangeOne()
        {
            try
            {
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cellValue = worksheet.Cells[row, 6].Value; // F sütunu (6. sütun)
                    if (cellValue != null && int.TryParse(cellValue.ToString(), out int number) && number == 0)
                    {
                        worksheet.Cells[row, 6].Value = 1; // Sıfır olan değeri bire çevirir
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Sıfırları bir ile değiştirirken hata oluştu.\nHata Kodu: {ex.Message}");
            }
        }

        public void textToColumn(int columnToSplit)
        {
            try
            {
                if (worksheet.Dimension == null)
                {
                    MessageBox.Show("Çalışma sayfası boş veya geçerli bir veri bulunamadı.");
                    return;
                }

                int lastRow = worksheet.Dimension.End.Row; // Son kullanılan satır
                int lastColumn = worksheet.Dimension.End.Column; // Son kullanılan sütun

                // Belirtilen sütundaki veriyi bölerek yeni sütunlara yaz
                for (int row = 1; row <= lastRow; row++)
                {
                    var cellValue = worksheet.Cells[row, columnToSplit].Value?.ToString(); // Bölünecek sütun

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        string[] splitValues = cellValue.Split('\\'); // `\` karakterine göre böl

                        for (int col = 0; col < splitValues.Length; col++)
                        {
                            worksheet.Cells[row, lastColumn + 1 + col].Value = splitValues[col];
                        }
                    }
                }

                // Orijinal sütunu sil
                worksheet.DeleteColumn(columnToSplit);
                worksheet.DeleteColumn(columnToSplit);
                worksheet.DeleteColumn(columnToSplit);

                // Yeni eklenen sütunlara başlık ekle (A, B, C... şeklinde)
                char columnLetter = 'A';
                for (int col = lastColumn; col <= worksheet.Dimension.End.Column; col++)
                {
                    worksheet.Cells[1, col].Value = columnLetter.ToString();
                    columnLetter++;
                }

                MessageBox.Show("Sütun bölme işlemi başarıyla tamamlandı.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Sütunlar bölünürken bir hata oluştu.\nHata Kodu: {ex.Message}");
            }
        }
        public void dataTable()
        {
            try
            {
                int rows = worksheet.Dimension.Rows;
                MessageBox.Show(rows.ToString());
                int columns = worksheet.Dimension.Columns;

                // Başlıkları ekle
                for (int col = 1; col <= columns; col++)
                {
                    dataTableList.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // Verileri ekle
                for (int row = 2; row < rows; row++)
                {
                    DataRow newRow = dataTableList.NewRow();
                    for (int col = 1; col <= columns; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTableList.Rows.Add(newRow);
                }
                Debug.Print("DataTable nesnesine aktarma başarılı...");
            }
            catch (Exception e)
            {
                MessageBox.Show("Kayıtlar aktarılırken beklenmedik bir hata oluştu. " +
                    "Lütfen teknik birim ile iletişime geçiniz.\n Hata kodu:" + e.Message.ToString());

            }
        }

        // Sayfa oluşturmak için aynı değerleri teke indirip diziye ekliyor

        private Dictionary<string, List<DataRow>> getGroupedData()
        {
            Dictionary<string, List<DataRow>> groupdata = new Dictionary<string, List<DataRow>>();

            foreach (DataRow row in dataTableList.Rows)
            {
                string key = row[columncontrolnumber].ToString();
                if (!groupdata.ContainsKey(key))
                {
                    groupdata[key] = new List<DataRow>();
                }
                groupdata[key].Add(row);
            }
            return groupdata;
        }

        public void splitAndSave()
        {
            // Her benzersiz değer için yeni sayfa oluştur
            foreach (var group in getGroupedData())
            {
                var sheetName = string.IsNullOrEmpty(group.Key) ? "Sheet_" + Guid.NewGuid().ToString() : group.Key;
                sheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                // Başlıkları yaz
                for (int col = 0; col < dataTableList.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTableList.Columns[col].ColumnName;
                }

                // Verileri yaz
                int row = 2;
                foreach (DataRow dataRow in group.Value)
                {
                    for (int col = 0; col < dataTableList.Columns.Count; col++)
                    {
                        worksheet.Cells[row, col + 1].Value = dataRow[col];
                    }
                    row++;
                }
            }

        }
        // Açık olan exceli kapatır.
        public void excelQuit()
        {
            package.Save();
        }

        /* public ExcelIslemler(ProgressBar progress, Label label)
         {
             this.progress = progress;
             this.label1 = label;
         }





         // Progresbar ilk ayarı için oluşturuldu.
         void progressBarSetting()
         {
             progress.Minimum = 0;
             progress.Maximum = rowsCount - 1;
             progress.Value = 0;
         }


         

         // Adresdeki exceli dataTableList nesnesine aktarır.
         


         // dataTableList nesnesini farklı yerlerde kullanabilmek için oluşturuldu.
         public DataTable getDataTable() { return dataTableList; }


         

         // dataTableList nesnesini yeni bir excele kaydeder.
         void newExcel()
         {
             workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
             foreach (String item in dict.Keys)
             {
                 // Yeni açılan excelde sayfalar oluşturur.
                 worksheet = workbook.Worksheets.Add();
                 worksheet.Name = sheetnamelenght(item.ToString());
                 worksheet = workbook.Worksheets[1];

                 Debug.Print("Sayfalar oluşturuluyor...");
                 for (int j = 0; j < dataTableList.Columns.Count; j++)
                 {
                     worksheet.Cells[1, j + 1] = dataTableList.Columns[j].ColumnName.ToString();
                 }

                 // Açılan sayfanın ismine göre sayfanın içine satırları koyar
                 int row = 1;
                 for (int i = 0; i < dataTableList.Rows.Count; i++)
                 {
                     if (dataTableList.Rows[i][columncontrolnumber].ToString() == item.ToString())
                     {
                         for (int j = 0; j < dataTableList.Columns.Count; j++)
                         {
                             worksheet.Cells[row + 1, j + 1] = dataTableList.Rows[i][j];
                         }
                         row++;
                     }
                 }
             }
             worksheet = workbook.Worksheets[workbook.Worksheets.Count];
             worksheet.Delete();
         }


         // Sayfa adı uzun ise ilk 15 karakteri alıyor.
         String sheetnamelenght(String value)
         {
             String control;
             if (value.Length < 32)
             {
                 control = value;
             }
             else
             {
                 control = value.Substring(0, 31).ToString();
             }
             return control;
         }


         // Kaydedilecek yeni excelde sayfa sayfa gezerek sütuna göre farklı olan satırların arasına boşluk bırakıp yeni excele aktarır.
         public void sheetRowSpace()
         {
             String sheetname = "";
             for (int i = 1; i <= workbook.Worksheets.Count; i++)
             {
                 worksheet = workbook.Worksheets[i];

                 //dataTableList nesnesini temizler
                 dataTableList.Clear();
                 Debug.Print(i.ToString() + ". sayfa DataTable nesnesine aktarımı yapılıyor...");
                 dataTable();

                 sheetname = worksheet.Name;
                 switch (sheetname)
                 {
                     case "A HABER":
                         columncontrolnumber = 10;
                         break;
                     case "A SPOR":
                         columncontrolnumber = 10;
                         break;
                     case "APARA":
                         columncontrolnumber = 10;
                         break;
                     case "ATV":
                         columncontrolnumber = 9;
                         break;
                     case "VAV":
                         columncontrolnumber = 9;
                         break;
                     case "TEKNIK BILGI ISLEM":
                         columncontrolnumber = 11;
                         break;
                     case "GENEL ARŞİV (AJANSLAR -İNGESTLE":
                         columncontrolnumber = 12;
                         break;
                     default:
                         columncontrolnumber = 8;
                         break;
                 }

                 Debug.Print("Sıralama yapılıyor...");
                 DataView dataView = dataTableList.DefaultView;
                 dataView.Sort = dataTableList.Columns[columncontrolnumber].ColumnName + " ASC";
                 dataTableList = dataView.ToTable();

                 Debug.Print("Farklı olan satırlar ayrılıyor...");
                 string prevValue = null;
                 for (int j = 0; j < dataTableList.Rows.Count; j++)
                 {
                     string currentValue = dataTableList.Rows[j][columncontrolnumber].ToString();

                     if (string.IsNullOrEmpty(currentValue))
                     {
                         continue;
                     }

                     if (prevValue == currentValue)
                     {
                         continue;
                     }

                     dataTableList.Rows.InsertAt(emptyRowSpace(), j);
                     prevValue = currentValue;
                 }

                 for (int j = 0; j < dataTableList.Rows.Count; j++)
                 {
                     for (int k = 0; k < dataTableList.Columns.Count; k++)
                     {
                         worksheet.Cells[j + 2, k + 1] = dataTableList.Rows[j][k].ToString();
                     }
                 }
             }
         }


         // Boş satır oluşturur.
         DataRow emptyRowSpace()
         {
             DataRow dr = dataTableList.NewRow();
             for (int j = 0; j < columnsCount; j++)
             {
                 dr[j] = "";
             }
             return dr;
         }


         // devam İşlem yapılacak exceldeki verileri dataTableList nesnesine aktarır. Yeni excel oluşturur ve yapılması gereken işlmelerden sonra yeni exceli kaydeder.
         public void saveExcel(String adres, String filename)
         {
             dataTableList.Clear();
             dataTable();
             sheetnamelist();
             excelquit(false);
             newExcel();
             sheetRowSpace();
             workbook.SaveAs(@adres + @"\" + filename, _Excel.XlFileFormat.xlWorkbookNormal);
             excelquit(true);
         }


         */

    }
}
