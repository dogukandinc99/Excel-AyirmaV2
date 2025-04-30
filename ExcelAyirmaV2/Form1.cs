using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelAyirmaV2
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        ExcelIslemler excel;
        String folderpatch = "";

        public Form1()
        {
            InitializeComponent();
            excel = new ExcelIslemler();
        }
        private void fileselectbtn_Click(object sender, EventArgs e)
        {
            ofd.Title = "Excel Dosyasý Seçiniz.";
            ofd.Filter = "Excel Dosyasý |*.xlsx; *.xls";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            ofd.Multiselect = true;
            ofd.ShowDialog();

            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                adresstxt.Text += ofd.FileNames[i].ToString() + Environment.NewLine;
            }
        }

        private void saveselectedfolderbtn_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                saveexcelbtn.Enabled = true;
                folderpatch = fbd.SelectedPath;
                saveadressfoldertxt.Text = folderpatch;
            }
        }

        private void saveexcelbtn_Click(object sender, EventArgs e)
        {
            /*for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(ofd.FileNames[i].ToString());
                dataGridView1.DataSource = excel.getDataTable();
                Debug.Print(ofd.SafeFileNames[i].ToString() + " adlý dosya için iþlem baþlatýlýyor...");
                excel.saveExcel(saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileNames[i].ToString());
                Debug.Print((i + 1) + " kayýdýn aktarýmý tamamlandý................................................");
            }
            MessageBox.Show("Kayýt iþlemi tamamlanmýþtýr.");*/
        }

        private void rdybtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(excel.ConvertXlsToXlsxInPlace(ofd.FileNames[i].ToString(), saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileNames[i].ToString()));
                excel.zeroChangeOne();
                excel.textToColumn(7);
                excel.dataTable();
                excel.splitAndSave();
                excel.excelQuit();
                Debug.Print((i + 1) + " adlý dosya hazýrlandý................................................");
            }
            MessageBox.Show("Excel dosyalarý ayýrma iþlemi için hazýrlandý.");
        }
    }
}
