using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
namespace WpfApplication17
{
    /// <summary>
    /// Interaction logic for HomePage.xaml
    /// </summary>
    ///         
    //public class KuryeDTO
    //    {
    //        public int Id { get; set; }
    //        public string vbMüsteri { get; set; }
    //        public string txtAdiki { get; set; }
    //    }

    public partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();
        }
        #region Kaydetme islemi
        private void BtnReg_Click(object sender, RoutedEventArgs e)
        {
            //KuryeDTO Kurye= new KuryeDTO();
            try
            {
                string vbMüsteri = this.vbMüsteri.Text;
                string txtAd = this.txtAd.Text;
                string txtAdiki = this.txtAdiki.Text;
                string txtSoyad = this.txtSoyad.Text;
                string txtKartNo = this.txtKartNo.Text;
                string cbxİslem = this.cbxİslem.Text;
                string cbxKurye = this.cbxKurye.Text;
                string txtUrunKod = this.txtUrunKod.Text;
                string cbxKTipi = this.cbxKTipi.Text;
                //string kayitTarih = this.kayitTarih.Text;
                //string teslimTarih = this.teslimTarih.Text;
                //string kuryeTarih = this.kuryeTarih.Text;
                string bsMüsteri = this.bsMüsteri.Text;
                string barkod = this.barkod.Text;
                string tcNo = this.tcNo.Text;
                string txtSube = this.txtSube.Text + this.cbxSube.Text;
                string txtUrunAd = this.txtUrunAd.Text;
                string txtBayi = this.txtBayi.Text;
                string cbxSozlesme = this.cbxSozlesme.Text;
                //string basimTarih = this.basimTarih.Text;
                //string iadeTarih = this.iadeTarih.Text;
                string date = this.kuryeTarih.Text;
                string kuryeTarih = DateTime.ParseExact(date, "dd.MM.yyyy",
                                CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                string date1 = this.teslimTarih.Text;
                string teslimTarih = DateTime.ParseExact(date, "dd.MM.yyyy",
                                CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                string date2 = this.kayitTarih.Text;
                string kayitTarih = DateTime.ParseExact(date, "dd.MM.yyyy",
                                CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                string date3 = this.basimTarih.Text;
                string basimTarih = DateTime.ParseExact(date, "dd.MM.yyyy",
                                CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                string date4 = this.iadeTarih.Text;
                string iadeTarih = DateTime.ParseExact(date, "dd.MM.yyyy",
                                CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                if (vbMüsteri != "" && txtAd != "" && txtSoyad != "" && txtKartNo != ""
                    && cbxİslem != "" && cbxKurye != "" && txtUrunKod != "" && cbxKTipi != "" &&
                    kayitTarih != "" && teslimTarih != "" && kuryeTarih != "" && bsMüsteri != "" &&
                    barkod != "" && tcNo != "" && txtSube != "" &&
                    txtUrunAd != "" && txtBayi != "" && cbxSozlesme != "" && basimTarih != "" &&
                    iadeTarih != "")
                {
                    HomeBusinessLogic.SaveInfo(vbMüsteri, txtAd, txtAdiki, txtSoyad, txtKartNo, cbxİslem, cbxKurye, txtUrunKod,
                   cbxKTipi, kayitTarih, teslimTarih, kuryeTarih, bsMüsteri, barkod, tcNo, txtSube, txtUrunAd,
                   txtBayi, cbxSozlesme, basimTarih, iadeTarih);
                    MessageBox.Show("Kayıt ekleme işlemi başarılı", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    foreach (UIElement element in MyGrid.Children)
                    {
                        TextBox textbox = element as TextBox;
                        if (textbox != null)
                        {
                            textbox.Text = String.Empty;
                        }

                    }
                }
                else
                {
                    MessageBox.Show("** ile gösterilen alanlari boş birakmayiniz.", "Fail", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                } 
            }
            catch (Exception ex)
            {
                Console.Write(ex);
                MessageBox.Show("Bir şeyler yanlış gitti, lütfen tekrar deneyiniz.", "Fail", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
        #region Listeleme islemi
        private void btnLstele_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                {
                    con.Open();
                    SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP", con);
                    System.Data.DataSet dsFald = new System.Data.DataSet();
                    adapvare.Fill(dsFald, "KARTTAKIP");
                    lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
        #region Kritere gore arama
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string tur = aracbx.Text.ToString();
            if (tur == "Kart No")
            {
                string sorgu = txtKartNo.Text.ToString();
                try
                {
                    using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        con.Open();
                        SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP WHERE KARTNO= "+"'"+sorgu+"'", con);
                        System.Data.DataSet dsFald = new System.Data.DataSet();
                        adapvare.Fill(dsFald, "KARTTAKIP");
                        lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            if (tur == "Ad")
            {
                string sorgu1 = txtAd.Text.ToString();
                try
                {
                    using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        con.Open();
                        SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP WHERE AD= " + "'" + sorgu1 + "'", con);
                        System.Data.DataSet dsFald = new System.Data.DataSet();
                        adapvare.Fill(dsFald, "KARTTAKIP");
                        lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            if (tur == "Soyad")
            {
                string sorgu1 = txtSoyad.Text.ToString();
                try
                {
                    using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        con.Open();
                        SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP WHERE SOYAD= " + "'" + sorgu1 + "'", con);
                        System.Data.DataSet dsFald = new System.Data.DataSet();
                        adapvare.Fill(dsFald, "KARTTAKIP");
                        lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            if (tur == "Vb Müşteri No")
            {
                string sorgu1 = vbMüsteri.Text.ToString();
                try
                {
                    using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        con.Open();
                        SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP WHERE VBMUSTERINO= " + "'" + sorgu1 + "'", con);
                        System.Data.DataSet dsFald = new System.Data.DataSet();
                        adapvare.Fill(dsFald, "KARTTAKIP");
                        lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            if (tur == "TC No")
            {
                string sorgu1 = tcNo.Text.ToString();
                try
                {
                    using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                    {
                        con.Open();
                        SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KARTTAKIP WHERE TCNO= " + "'" + sorgu1 + "'", con);
                        System.Data.DataSet dsFald = new System.Data.DataSet();
                        adapvare.Fill(dsFald, "KARTTAKIP");
                        lstDeneme.DataContext = dsFald.Tables["KARTTAKIP"].DefaultView;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        #endregion
        #region Detayekleme
        private void btnDetayEkle_Click(object sender, RoutedEventArgs e)
        {
            Page2 p2 = new Page2();
            this.NavigationService.Navigate(p2);
        }
        #endregion
        #region Databasedeki verileri sqle kaydetme
        private void btnSqltoexcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SqlConnection cnn;
                string connectionString = null;
                string sql = null;
                string data = null;
                int i = 0;
                int j = 0;

                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                connectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30";
                cnn = new SqlConnection(connectionString);
                cnn.Open();
                sql = "SELECT * FROM KARTTAKIP";
                SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn);
                DataSet ds = new DataSet();
                dscmd.Fill(ds);

                foreach (DataTable dt in ds.Tables)
                {
                    for (int i1 = 0; i1 < dt.Columns.Count; i1++)
                    {
                        xlWorkSheet.Cells[1, i1 + 1] = dt.Columns[i1].ColumnName;
                    }
                }

                for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    int s = i + 1;
                    for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                    {
                        data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                        xlWorkSheet.Cells[s + 1, j + 1] = data;
                    }
                }

                xlWorkBook.SaveAs(@"D:\Users\st900398\Documents\Visual Studio 2013\Projects\WpfApplication17\WpfApplication17\KARTTAKIP.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Excel dosyası olusturuldu , dosya yolu D:\\Users\\st900398\\Documents\\Visual Studio 2013\\Projects\\WpfApplication17\\WpfApplication17\\KARTTAKIP.xls");
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Please Move Your SaleOrder.xls and Retry Again...........");

            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }
        #endregion
    }
    }

