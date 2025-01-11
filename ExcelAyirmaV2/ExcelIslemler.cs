using OfficeOpenXml;
using _Excel = Microsoft.Office.Interop.Excel;
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

        public void InsertBlankRowInSpecificSheet(string sheetName, int columnToCheck)
        {
            try
            {
                if (package == null || package.Workbook == null || package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("Çalışma kitabı boş veya geçerli bir veri bulunamadı.");
                    return;
                }

                // Belirtilen sayfayı bul
                var sheet = package.Workbook.Worksheets[sheetName];
                if (sheet == null || sheet.Dimension == null)
                {
                    MessageBox.Show($"'{sheetName}' adlı sayfa bulunamadı veya boş.");
                    return;
                }

                int lastRow = sheet.Dimension.End.Row;
                string previousValue = null;

                // Son satırdan başlayarak farklı değerler arasında boşluk ekle
                for (int row = lastRow; row >= 2; row--)
                {
                    var cellValue = sheet.Cells[row, columnToCheck].Value?.ToString();

                    if (previousValue != null && cellValue != previousValue)
                    {
                        sheet.InsertRow(row + 1, 1); // Farklı değerler arasında satır ekle
                    }

                    previousValue = cellValue;
                }

                MessageBox.Show($"'{sheetName}' sayfasında boş satır ekleme işlemi başarıyla tamamlandı.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Bir hata oluştu: {ex.Message}");
            }
        }

        // Kaydedilecek yeni excelde sayfa sayfa gezerek sütuna göre farklı olan satırların arasına boşluk bırakıp yeni excele aktarır.
        public void sheetRowSpace()
        {
            String sheetname = "";
            for (int i = 1; i < package.Workbook.Worksheets.Count; i++)
            {
                worksheet = package.Workbook.Worksheets[i];

                sheetname = worksheet.Name;
                switch (sheetname)
                {
                    case "A HABER":
                        InsertBlankRowInSpecificSheet("A HABER", 11);
                        break;
                    case "A SPOR":
                        InsertBlankRowInSpecificSheet("A SPOR", 11);
                        break;
                    case "APARA":
                        InsertBlankRowInSpecificSheet("APARA", 11);
                        break;
                    case "ATV":
                        InsertBlankRowInSpecificSheet("ATV", 10);
                        break;
                    case "VAV":
                        InsertBlankRowInSpecificSheet("VAV", 10);
                        break;
                    case "TEKNIK BILGI ISLEM":
                        InsertBlankRowInSpecificSheet("TEKNIK BILGI ISLEM", 12);
                        break;
                    case "GENEL ARŞİV (AJANSLAR -İNGESTLE":
                        InsertBlankRowInSpecificSheet("GENEL ARŞİV (AJANSLAR -İNGESTLE", 13);
                        break;
                    default:
                        columncontrolnumber = 8;
                        break;
                }
            }
        }

        // Açık olan exceli kapatır.
        public void excelQuit()
        {
            package.Save();
        }        
    }
}
