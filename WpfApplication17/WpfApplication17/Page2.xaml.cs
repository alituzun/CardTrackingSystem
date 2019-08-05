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
    /// Interaction logic for Page2.xaml
    /// </summary>
    public partial class Page2 : Page
    {
        public Page2()
        {
            InitializeComponent();
            SqlConnection oSqlConn = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30");
            SqlDataAdapter da = new SqlDataAdapter("SELECT KARTID FROM KARTTAKIP", oSqlConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cbxId.Items.Add(dt.Rows[i]["KARTID"].ToString());
            }
        }
        #region kaydetme islemi
        private void btnKaydet_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (cbxId.Text.ToString() != "" && kayitDate.Text.ToString() != "")
                {
                    string kartId = cbxId.Text.ToString();
                    string Islemkodu = this.txtIslemKodu.Text.ToString();
                    string islemaciklama = this.txtİslemAciklama.Text.ToString();
                    string kayit = this.kayitDate.Text.ToString();
                    string kayittarihi = DateTime.ParseExact(kayit, "dd.MM.yyyy",
                                        CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                    string kayitusercode = this.txtKayitUser.Text.ToString();
                    string kayitkanalkodu = this.txtKayitKanal.Text.ToString();
                    HomeBusinessLogic.DetaySave(kartId, Islemkodu, kayittarihi, kayitusercode, kayitkanalkodu, islemaciklama);
                    MessageBox.Show("Kayıt ekleme işlemi başarılı", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("KardId kısmını ve Kayıt Tarihi alanlarını boş bırakmayınız.", "Fail", MessageBoxButton.OK, MessageBoxImage.Error);
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
        #region listeleme islemi
        private void btnListele_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30"))
                {
                    con.Open();
                    SqlDataAdapter adapvare = new SqlDataAdapter("SELECT * FROM KURYETAKIP", con);
                    System.Data.DataSet dsFald = new System.Data.DataSet();
                    adapvare.Fill(dsFald, "KURYETAKIP");
                    detayList.DataContext = dsFald.Tables["KURYETAKIP"].DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
    }
}
