using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace Lisans_Tezi11
{
    public partial class Form3 : Form
    {
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls; Extended Properties='Excel 12.0 xml;HDR=YES;'");
        // ana verilerin tutulduğu excel dosyasının bağlantısı gerçekleştirilir.
        DataTable tablo = new DataTable();
        ErrorProvider provider = new ErrorProvider();
        int Textbox_kontrol_sayacı = 0; //kat sayıların ve kontenjanların girildiği textboxların kontrol bayrağıdır.


        string Milliolmayan_BE_ERKEK = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls";
        string Milliolmayan_BE_KADIN = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls";
        string Milliolmayan_ANT_ERKEK = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls";
        string Milliolmayan_ANT_KADIN = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls";
        string Milliolmayan_REK_ERKEK = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls";
        string Milliolmayan_REK_KADIN = @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls";
        string[] dizi = { @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls", @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls", @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls", @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls", @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls", @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls" };
        // üst kısım excel dosyalarının dosya yolunun stringe aktarılmış halidir. uzun kod yazımlarından kurtulmak için bu kullanım benimsenmiştir.

        int s_m_olmayan_be_e = 0;
        int s_m_olmayan_ant_e = 0;
        int s_m_olmayan_rek_e = 0;
        int s_m_olmayan_be_k = 0;
        int s_m_olmayan_ant_k = 0;
        int s_m_olmayan_rek_k = 0;
        int m_olmayan_be_e = 0;
        int m_olmayan_ant_e = 0;
        int m_olmayan_rek_e = 0;
        int m_olmayan_be_k = 0;
        int m_olmayan_ant_k = 0;
        int m_olmayan_rek_k = 0;
        int baslangıc_aday_kontenjan_kontrol_sayac = 0;
        // kontenjanların tutulduğu sabit ve değişken kontenjanlardır. kontenjan aktarımı ve aday yerleştirmede kullanılır.

        public Form3()
        {
            InitializeComponent();
        }
        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            // başlangıçta ana verilerin tutulduğu excelden veriler çekilir. ve aday yerleştirme excelleri yoksa oluşturulur.
            progressBar1.Visible = false;
            Excel_Dataset_Veri_Cekme();
            Aday_Tabloları_Kontrol();
        }
        void textbox_kontrol()
        {  // kontenjanların ve kat sayıların alındığı textboxların kontrol methodudur. hatalı olan işlemlerde textboxın yanında kırmızı ünlem ile hatayı belirtir.
            Textbox_kontrol_sayacı = 0;
            provider.BlinkRate = 0;
            provider.Clear();
            if (gk_ant_e_txt.Text == ""
                || gk_ant_k_txt.Text == ""
                || gk_be_e_txt.Text == ""
                || gk_be_k_txt.Text == ""
                || gk_rek_e_txt.Text == ""
                || m_o_ant_e_txt.Text == ""
                || m_o_ant_k_txt.Text == ""
                || m_o_be_e_txt.Text == ""
                || m_o_be_k_txt.Text == ""
                || m_o_rek_e_txt.Text == ""
                || m_o_rek_k_txt.Text == ""
                || m_o_olmayan_ant_e_txt.Text == ""
                || m_o_olmayan_ant_k_txt.Text == ""
                || m_o_olmayan_be_e_txt.Text == ""
                || m_o_olmayan_be_k_txt.Text == ""
                || m_o_olmayan_rek_e_txt.Text == ""
                || m_o_olmayan_rek_k_txt.Text == ""
                || gk_rek_k_txt.Text == "") { MessageBox.Show("!!!Kontenjan Değerleri Boş Geçilemez!!!\nKontenjan'ı Açılmayan Bölümler İçin 0 Değerini Giriniz."); Textbox_kontrol_sayacı++; }
            if (maskedTextBox1.Text == " ,") { provider.SetError(maskedTextBox1, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (maskedTextBox2.Text == " ,") { provider.SetError(maskedTextBox2, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (maskedTextBox3.Text == " ,") { provider.SetError(maskedTextBox3, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (maskedTextBox4.Text == " ,") { provider.SetError(maskedTextBox4, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (maskedTextBox5.Text == " ,") { provider.SetError(maskedTextBox5, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (maskedTextBox6.Text == " ,") { provider.SetError(maskedTextBox6, "Bu alan boş geçilemez"); Textbox_kontrol_sayacı++; }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox1.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox1, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox2.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox2, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox3.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox3, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox4.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox4, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox5.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox5, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
            if (Textbox_kontrol_sayacı == 0)
            {
                string a = maskedTextBox6.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 4) { provider.SetError(maskedTextBox6, "Eksik veya Hatalı Giriş Yaptınız."); Textbox_kontrol_sayacı++; }
            }
        }
        void Aday_Tabloları_Kontrol()
        {
            // aday tablolarlının kontolleri gerçekleştirilir. olmayanlar oluşturulur.
            if (Directory.Exists("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları") == false)
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları");
                Directory.CreateDirectory("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milliler");
                Directory.CreateDirectory("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milli Olmayanlar");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls");
            }
            if (Directory.Exists("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milli Olmayanlar") == false) 
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milli Olmayanlar");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls");
            }
            if (Directory.Exists("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milliler") == false) 
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Düzenlenmiş Aday Tabloları\\Milliler");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_ERKEK.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_KADIN.xls");
                Aday_Tablo_Oluşturma_Tek_tek(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_KADIN.xls");
            }

            if (Directory.Exists("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları") == false)
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları");
                Directory.CreateDirectory("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milliler");
                Directory.CreateDirectory("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_KADIN.xls");
            }
            if (Directory.Exists("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milliler") == false)
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milliler");
            }
            if (Directory.Exists("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar") == false)
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_KADIN.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_ERKEK.xls");
                Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek("C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_KADIN.xls");
            }

        }
        void Yerlesmis_Aday_Tablo_Oluşturma_Tek_tek(string dosyayolu)
        {
            // gömülü excel dosyalarını ilgili bölüm ve isime göre oluşturma methodudur. (yerleştirilmiş adayların tutulduğu excel şablonu)
            byte[] excel = Properties.Resources.aday_tabloları;
            System.IO.FileStream fs = new System.IO.FileStream(dosyayolu, FileMode.CreateNew, FileAccess.ReadWrite);
            foreach (byte b in excel)
            {
                fs.WriteByte(b);
            }
            fs.Close();
        }
        void Aday_Tablo_Oluşturma_Tek_tek(string dosyayolu)
        {    // gömülü excel dosyalarını ilgili bölüm ve isime göre oluşturma methodudur. (yerleştirme öncesi adayların ayrıştırılıp puanlarının hesaplandığı excel şablonu)
            byte[] excel = Properties.Resources.YERLEŞTİRME;
            System.IO.FileStream fs = new System.IO.FileStream(dosyayolu, FileMode.CreateNew, FileAccess.ReadWrite);
            foreach (byte b in excel)
            {
                fs.WriteByte(b);
            }
            fs.Close();
        }
        void Excel_Silme(string Dosyayolu)
        {
            // dosya konumuna göre excel dosyası silme methodu.

            if (System.IO.File.Exists(Dosyayolu))
            {
                System.IO.File.Delete(Dosyayolu);
            }
        }
        void Excel_Dataset_Veri_Cekme()
        {   // ana verilerin tutulduğu excel dosyasından verileri çekme komutu
            baglanti.Open();
            tablo.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", baglanti);
            da.Fill(tablo);
            baglanti.Close();
        }
        void Excel_Dosya_Konum_ve_Veri_Ekleme(string Dosyayolu, DataTable dt, int satır)
        { 
            // başvuruduğu bölüme göre adayların tektek ayrıştırıldığı methodur.
           
            OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Dosyayolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            OleDbCommand komut = new OleDbCommand();
            nbaglanti.Open();
            komut.Connection = nbaglanti;
            string sql = "Insert into [KAYIT-GİRİŞ$] ([Kayıt No],[TC Kimlik No],[Adı ],Soyadı,Cinsiyeti,[Doğum Tarihi],[Alan/Kol],[TYT P],AOÖBP,[SÖ Durumu],[SÖ Belge Türü],[Millilik ],Olimpiklik,[Spor Özgeçmiş Katsayısı BE],[Spor Özgeçmiş Katsayısı ANT],[Spor Özgeçmiş Katsayısı REK],[Millilikten Yerleşemediği Durumda SÖK BE],[Millilikten Yerleşemediği Durumda SÖK ANT],[Millilikten Yerleşemediği Durumda SÖK REK],[Daha Önce Yerleşme Durumu],[Başvurduğu Programlar BE],[Başvurduğu Programlar ANT],[Başvurduğu Programlar REK],[Millilikten Yerleştiği Tercihi],[ANTRENÖRLÜK BRANŞ],[MEKİK KOŞUSU]) values('" + dt.Rows[satır][0].ToString() + "','" + dt.Rows[satır][1].ToString() + "','" + dt.Rows[satır][2].ToString() + "','" + dt.Rows[satır][3].ToString() + "','" + dt.Rows[satır][4].ToString() + "','" + dt.Rows[satır][5].ToString() + "','" + dt.Rows[satır][6].ToString() + "','" + dt.Rows[satır][7].ToString() + "','" + dt.Rows[satır][8].ToString() + "','" + dt.Rows[satır][9].ToString() + "','" + dt.Rows[satır][10].ToString() + "','" + dt.Rows[satır][11].ToString() + "','" + dt.Rows[satır][12].ToString() + "','" + dt.Rows[satır][13].ToString() + "','" + dt.Rows[satır][14].ToString() + "','" + dt.Rows[satır][15].ToString() + "','" + dt.Rows[satır][16].ToString() + "','" + dt.Rows[satır][17].ToString() + "','" + dt.Rows[satır][18].ToString() + "','" + dt.Rows[satır][19].ToString() + "','" + dt.Rows[satır][20].ToString() + "','" + dt.Rows[satır][21].ToString() + "','" + dt.Rows[satır][22].ToString() + "','" + dt.Rows[satır][23].ToString() + "','" + dt.Rows[satır][24].ToString() + "','" + dt.Rows[satır][25].ToString() + "')";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            nbaglanti.Close();

        }
        void Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(string Dosyayolu, DataTable dt, int satır)
        {   
            // milliliğe başvurup yerleşemeyen adaylar için kullanılan methoddur. (normal sök değerleri millilikten yerleşemediği durumdaki sök değerleri ile güncellenir ve milli olamayan adayların listesine eklenir.)

            OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Dosyayolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            OleDbCommand komut = new OleDbCommand();
            nbaglanti.Open();
            komut.Connection = nbaglanti;
            string sql = "Insert into [KAYIT-GİRİŞ$] ([Kayıt No],[TC Kimlik No],[Adı ],Soyadı,Cinsiyeti,[Doğum Tarihi],[Alan/Kol],[TYT P],AOÖBP,[SÖ Durumu],[SÖ Belge Türü],[Millilik ],Olimpiklik,[Spor Özgeçmiş Katsayısı BE],[Spor Özgeçmiş Katsayısı ANT],[Spor Özgeçmiş Katsayısı REK],[Millilikten Yerleşemediği Durumda SÖK BE],[Millilikten Yerleşemediği Durumda SÖK ANT],[Millilikten Yerleşemediği Durumda SÖK REK],[Daha Önce Yerleşme Durumu],[Başvurduğu Programlar BE],[Başvurduğu Programlar ANT],[Başvurduğu Programlar REK],[Millilikten Yerleştiği Tercihi],[ANTRENÖRLÜK BRANŞ],[MEKİK KOŞUSU]) values('" + dt.Rows[satır][0].ToString() + "','" + dt.Rows[satır][1].ToString() + "','" + dt.Rows[satır][2].ToString() + "','" + dt.Rows[satır][3].ToString() + "','" + dt.Rows[satır][4].ToString() + "','" + dt.Rows[satır][5].ToString() + "','" + dt.Rows[satır][6].ToString() + "','" + dt.Rows[satır][7].ToString() + "','" + dt.Rows[satır][8].ToString() + "','" + dt.Rows[satır][9].ToString() + "','" + dt.Rows[satır][10].ToString() + "','" + dt.Rows[satır][11].ToString() + "','" + dt.Rows[satır][12].ToString() + "','" + dt.Rows[satır][16].ToString() + "','" + dt.Rows[satır][17].ToString() + "','" + dt.Rows[satır][17].ToString() + "','" + dt.Rows[satır][16].ToString() + "','" + dt.Rows[satır][17].ToString() + "','" + dt.Rows[satır][18].ToString() + "','" + dt.Rows[satır][19].ToString() + "','" + dt.Rows[satır][20].ToString() + "','" + dt.Rows[satır][21].ToString() + "','" + dt.Rows[satır][22].ToString() + "','" + dt.Rows[satır][23].ToString() + "','" + dt.Rows[satır][24].ToString() + "','" + dt.Rows[satır][25].ToString() + "')";
            komut.CommandText = sql;
            komut.ExecuteNonQuery();
            nbaglanti.Close();

        }

        void milli_adayları_ayır() 
        {  
            // milli adayların tercihi, cinsiyeti, ve olimpiklik kontenjanına göre ayrıştırıldığı methoddur.

            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Maximum = tablo.Rows.Count;
            for (int i = 0; i < tablo.Rows.Count; i++)
            {
                progressBar1.Value++;
                string millilik = tablo.Rows[i][11].ToString();
                string cinsiyet = tablo.Rows[i][4].ToString();
                string BE = tablo.Rows[i][20].ToString();
                string ANT = tablo.Rows[i][21].ToString();
                string REK = tablo.Rows[i][22].ToString();
                string olimpiklik = tablo.Rows[i][12].ToString();
                string mekikkosusu = tablo.Rows[i][25].ToString();
                string millik_yerleşme_durumu = tablo.Rows[i][23].ToString();
                if (millilik == "VAR")
                {
                    if (cinsiyet == "E" || cinsiyet == "ERKEK")
                    {
                        if (((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "VAR") || ((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "YOK" && (m_o_olmayan_be_e_txt.Text != "" && m_o_olmayan_be_e_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_ERKEK.xls", tablo, i);
                        }
                        if (((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "VAR") || ((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "YOK" && (m_o_olmayan_ant_e_txt.Text != "" && m_o_olmayan_ant_e_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_ERKEK.xls", tablo, i);
                        }
                        if (((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "VAR") || ((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "YOK" && (m_o_olmayan_rek_e_txt.Text != "" && m_o_olmayan_rek_e_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_ERKEK.xls", tablo, i);
                        }
                    }
                    else if (cinsiyet == "K" || cinsiyet == "KADIN")
                    {
                        if (((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "VAR") || ((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "YOK" && (m_o_olmayan_be_k_txt.Text != "" && m_o_olmayan_be_k_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_KADIN.xls", tablo, i);
                        }
                        if (((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "VAR") || ((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "YOK" && (m_o_olmayan_ant_k_txt.Text != "" && m_o_olmayan_ant_k_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_KADIN.xls", tablo, i);
                        }
                        if (((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "VAR") || ((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "YOK" && (m_o_olmayan_rek_k_txt.Text != "" && m_o_olmayan_rek_k_txt.Text != "0")))
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_KADIN.xls", tablo, i);
                        }
                    }

                }
            }
        }

        void Adayları_Excellere_Ayır()
        {
            // milli olmayan adayların ve millilikten yerleşemeyen adayların tercih, cinsiyet, olimpiklik kontenjanına göre ayrıştırıldığı methoddur.

            progressBar1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Maximum = tablo.Rows.Count;
            for (int i = 0; i < tablo.Rows.Count; i++)
            {
                progressBar1.Value++;
                string millilik = tablo.Rows[i][11].ToString();
                string cinsiyet = tablo.Rows[i][4].ToString();
                string BE = tablo.Rows[i][20].ToString();
                string ANT = tablo.Rows[i][21].ToString();
                string REK = tablo.Rows[i][22].ToString();
                string olimpiklik = tablo.Rows[i][12].ToString();
                string mekikkosusu = tablo.Rows[i][25].ToString();
                string millik_yerleşme_durumu = tablo.Rows[i][23].ToString();
                if (millilik == "VAR" && mekikkosusu == "BAŞARILI")
                {
                    if (cinsiyet == "E" || cinsiyet == "ERKEK")
                    {
                        if (((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "VAR") || ((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "YOK" && (m_o_olmayan_be_e_txt.Text != "" && m_o_olmayan_be_e_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(BE) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls", tablo, i);
                        }


                        if (((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "VAR") || ((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "YOK" && (m_o_olmayan_ant_e_txt.Text != "" && m_o_olmayan_ant_e_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(ANT) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls", tablo, i);
                        }
                        if (((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "VAR") || ((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "YOK" && (m_o_olmayan_rek_e_txt.Text != "" && m_o_olmayan_rek_e_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(REK) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls", tablo, i);
                        }
                    }
                    else if (cinsiyet == "K" || cinsiyet == "KADIN")
                    {
                        if (((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "VAR") || ((BE == "1" || BE == "2" || BE == "3") && olimpiklik == "YOK" && (m_o_olmayan_be_k_txt.Text != "" && m_o_olmayan_be_k_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(BE) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls", tablo, i);
                        }
                        if (((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "VAR") || ((ANT == "1" || ANT == "2" || ANT == "3") && olimpiklik == "YOK" && (m_o_olmayan_ant_k_txt.Text != "" && m_o_olmayan_ant_k_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(ANT) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls", tablo, i);
                        }
                        if (((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "VAR") || ((REK == "1" || REK == "2" || REK == "3") && olimpiklik == "YOK" && (m_o_olmayan_rek_k_txt.Text != "" && m_o_olmayan_rek_k_txt.Text != "0")))
                        {
                            if (millik_yerleşme_durumu != "")
                            {
                                if (Convert.ToInt32(REK) < Convert.ToInt32(millik_yerleşme_durumu))
                                {
                                    Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls", tablo, i);
                                }
                            }
                            else if (millik_yerleşme_durumu == "") Excel_Dosya_Konum_ve_Veri_Ekleme_Millilikten_Yerlesemezse(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls", tablo, i);
                        }
                    }
                }
                else if (millilik == "YOK" && mekikkosusu == "BAŞARILI")
                {
                    if (cinsiyet == "E" || cinsiyet == "ERKEK")
                    {
                        if (BE == "1" || BE == "2" || BE == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls", tablo, i);
                        }
                        if (ANT == "1" || ANT == "2" || ANT == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls", tablo, i);
                        }
                        if (REK == "1" || REK == "2" || REK == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls", tablo, i);
                        }
                    }
                    else if (cinsiyet == "K" || cinsiyet == "KADIN")
                    {
                        if (BE == "1" || BE == "2" || BE == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls", tablo, i);
                        }
                        if (ANT == "1" || ANT == "2" || ANT == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls", tablo, i);
                        }
                        if (REK == "1" || REK == "2" || REK == "3")
                        {
                            Excel_Dosya_Konum_ve_Veri_Ekleme(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls", tablo, i);
                        }
                    }
                }

            }
            progressBar1.Visible = false;
        }
        void PUAN_hesapla(string Dosyayolu)
        {  
            // spor bilimleri tarafndan yayınlanan puan hesaplama formülleri kullanılarak adayların puanlarının hesaplandığı ve puana göre azalan sıralamaya sokulduğu methoddur.
            
            progressBar1.Visible = true;
            OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Dosyayolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            DataTable dt = new DataTable();
            nbaglanti.Open();
            dt.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", nbaglanti);
            da.Fill(dt);
            nbaglanti.Close();
            progressBar1.Value = 0;
            progressBar1.Maximum = dt.Rows.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                progressBar1.Value++;
                double sok_kat_sayı = 0.0;
                double be_katsayı = 0.0;
                double ant_katsayı = 0.0;
                double rek_katsayı = 0.0;

                if (dt.Rows[i][13].ToString() != "" && dt.Rows[i][13].ToString() != " ,") { be_katsayı = Convert.ToDouble(dt.Rows[i][13]); }

                if (dt.Rows[i][14].ToString() != "" && dt.Rows[i][14].ToString() != " ,") { ant_katsayı = Convert.ToDouble(dt.Rows[i][14]); }

                if (dt.Rows[i][15].ToString() != "" && dt.Rows[i][15].ToString() != " ,") { rek_katsayı = Convert.ToDouble(dt.Rows[i][15]); }

                double tyt_p = Convert.ToDouble(dt.Rows[i][7].ToString());

                string kayıt_no = dt.Rows[i][0].ToString();


                if (Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_KADIN.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls") { sok_kat_sayı = be_katsayı; }

                if (Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_KADIN.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls") { sok_kat_sayı = ant_katsayı; }

                if (Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_KADIN.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls" ||
                    Dosyayolu == @"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls") { sok_kat_sayı = rek_katsayı; }

                double sop = tyt_p * sok_kat_sayı;

                OleDbCommand komut = new OleDbCommand();

                nbaglanti.Open();

                komut.Connection = nbaglanti;

                string sql = "Update [KAYIT-GİRİŞ$] set SÖP='" + sop.ToString() + "' WHERE [Kayıt No]='" + kayıt_no + "'";

                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                nbaglanti.Close();
            }

            nbaglanti.Open();
            dt.Clear();
            da.Fill(dt);
            nbaglanti.Close();

            double sopdo = 0.0;                                     //spor özgeçmiş puanları dağılımı ortalaması
            double sopt = 0.0;                                      //spor özgeçmiş puanları toplamı
            double adaysayısı = Convert.ToDouble(dt.Rows.Count);
            double sopkarelerinintoplamı = 0.0;                     //spor özgeçmiş puanlarının kareleri toplamı
            double soptoplamlarınınkaresi = 0.0;                    //spor özgeçmiş puanları toplamlarının karesi
            double sopdss = 0.0;                                    //spor özgeçmiş puanları dağılımının standart sapması 

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                double pt = Convert.ToDouble(dt.Rows[i][26]); //spor özgeçmiş puanları
                sopt += pt;
                sopkarelerinintoplamı += pt * pt;
            }

            sopdo = sopt / adaysayısı;

            soptoplamlarınınkaresi = sopt * sopt;

            double soptk_adaysayısı = soptoplamlarınınkaresi / adaysayısı;

            double sopds_ustkısım = sopkarelerinintoplamı - soptk_adaysayısı;

            sopdss = Math.Sqrt(sopds_ustkısım / (adaysayısı - 1));

            progressBar1.Value = 0;
            progressBar1.Maximum = dt.Rows.Count;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                progressBar1.Value++;

                string kayıt_no = dt.Rows[i][0].ToString();

                double asop = Convert.ToDouble(dt.Rows[i][26]);

                double sop_sp_ustkisim = asop - sopdo;

                double sop_orta_sonuc = sop_sp_ustkisim / sopdss;

                double sop_sp = (10 * sop_orta_sonuc) + 50;
                sop_sp = Math.Round(sop_sp, 6);

                OleDbCommand komut = new OleDbCommand();

                nbaglanti.Open();
                komut.Connection = nbaglanti;
                string sql = "Update [KAYIT-GİRİŞ$] set [SÖP-SP]='" + sop_sp.ToString() + "',[ÖYSP-SP]='" + sop_sp.ToString() + "' WHERE [Kayıt No]='" + kayıt_no + "'";
                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                nbaglanti.Close();
            }

            nbaglanti.Open();
            dt.Clear();
            da.Fill(dt);
            nbaglanti.Close();
            progressBar1.Value = 0;
            progressBar1.Maximum = dt.Rows.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                progressBar1.Value++;
                string kayıt_no = dt.Rows[i][0].ToString();

                double OYSP_SP_Katsayı = 0.0;

                double OBP_Katsayı = 0.0;

                double TYT_P_Katsayı = 0.0;

                double OYSP = Convert.ToDouble(dt.Rows[i][28]);
                double OBP = Convert.ToDouble(dt.Rows[i][8]);
                double TYT_P = Convert.ToDouble(dt.Rows[i][7]);
                double YP = 0.0;
                string alan_kol = dt.Rows[i][6].ToString();
                string yerlesme_durumu = dt.Rows[i][19].ToString();
                if (alan_kol == "SPOR")
                {
                    OYSP_SP_Katsayı = Convert.ToDouble(maskedTextBox1.Text);
                    OBP_Katsayı = Convert.ToDouble(maskedTextBox2.Text);
                    TYT_P_Katsayı = Convert.ToDouble(maskedTextBox3.Text);
                }
                else if (alan_kol == "DİĞER")
                {
                    OYSP_SP_Katsayı = Convert.ToDouble(maskedTextBox4.Text);
                    OBP_Katsayı = Convert.ToDouble(maskedTextBox5.Text);
                    TYT_P_Katsayı = Convert.ToDouble(maskedTextBox6.Text);
                }
                if (yerlesme_durumu == "EVET") { OBP_Katsayı = OBP_Katsayı / 2; }
                YP = (OYSP_SP_Katsayı * OYSP) + (OBP_Katsayı * OBP) + (TYT_P_Katsayı * TYT_P);
                YP = Math.Round(YP, 6);
                OleDbCommand komut = new OleDbCommand();
                nbaglanti.Open();
                komut.Connection = nbaglanti;
                string sql = "Update [KAYIT-GİRİŞ$] set YP='" + YP + "' WHERE [Kayıt No]='" + kayıt_no + "'";
                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                nbaglanti.Close();
            }

            nbaglanti.Open();
            dt.Clear();
            da.Fill(dt);
            DataView dtw = dt.DefaultView;
            dtw.Sort = "YP desc";
            DataTable dts = dtw.ToTable();
            nbaglanti.Close();
            Excel_Silme(Dosyayolu);
            Aday_Tablo_Oluşturma_Tek_tek(Dosyayolu);
            progressBar1.Value = 0;
            progressBar1.Maximum = dt.Rows.Count;
            for (int i = 0; i < dts.Rows.Count; i++)
            {
                progressBar1.Value++;
                OleDbCommand komut = new OleDbCommand();
                nbaglanti.Open();
                komut.Connection = nbaglanti;
                string sql = "Insert into [KAYIT-GİRİŞ$] ([Kayıt No],[TC Kimlik No],[Adı ],Soyadı,Cinsiyeti,[Doğum Tarihi],[Alan/Kol],[TYT P],AOÖBP,[SÖ Durumu],[SÖ Belge Türü],[Millilik ],Olimpiklik,[Spor Özgeçmiş Katsayısı BE],[Spor Özgeçmiş Katsayısı ANT],[Spor Özgeçmiş Katsayısı REK],[Millilikten Yerleşemediği Durumda SÖK BE],[Millilikten Yerleşemediği Durumda SÖK ANT],[Millilikten Yerleşemediği Durumda SÖK REK],[Daha Önce Yerleşme Durumu],[Başvurduğu Programlar BE],[Başvurduğu Programlar ANT],[Başvurduğu Programlar REK],[Millilikten Yerleştiği Tercihi],[ANTRENÖRLÜK BRANŞ],[MEKİK KOŞUSU],SÖP,[SÖP-SP],[ÖYSP-SP],YP,[DİĞER LİSTE],SONUÇ,[YERLEŞME SIRALAMASI]) values('" + dts.Rows[i][0].ToString() + "','" + dts.Rows[i][1].ToString() + "','" + dts.Rows[i][2].ToString() + "','" + dts.Rows[i][3].ToString() + "','" + dts.Rows[i][4].ToString() + "','" + dts.Rows[i][5].ToString() + "','" + dts.Rows[i][6].ToString() + "','" + dts.Rows[i][7].ToString() + "','" + dts.Rows[i][8].ToString() + "','" + dts.Rows[i][9].ToString() + "','" + dts.Rows[i][10].ToString() + "','" + dts.Rows[i][11].ToString() + "','" + dts.Rows[i][12].ToString() + "','" + dts.Rows[i][13].ToString() + "','" + dts.Rows[i][14].ToString() + "','" + dts.Rows[i][15].ToString() + "','" + dts.Rows[i][16].ToString() + "','" + dts.Rows[i][17].ToString() + "','" + dts.Rows[i][18].ToString() + "','" + dts.Rows[i][19].ToString() + "','" + dts.Rows[i][20].ToString() + "','" + dts.Rows[i][21].ToString() + "','" + dts.Rows[i][22].ToString() + "','" + dts.Rows[i][23].ToString() + "','" + dts.Rows[i][24].ToString() + "','" + dts.Rows[i][25].ToString() + "','" + dts.Rows[i][26].ToString() + "','" + dts.Rows[i][27].ToString() + "','" + dts.Rows[i][28].ToString() + "','" + dts.Rows[i][29].ToString() + "','" + dts.Rows[i][30].ToString() + "','" + dts.Rows[i][31].ToString() + "','" + dts.Rows[i][32].ToString() + "')";
                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                nbaglanti.Close();
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {  
            // milli olmayan adaylar butonuna tıklandığında yapılan işlemler sıralaması

            textbox_kontrol();
            if (Textbox_kontrol_sayacı == 0)
            {
                s_m_olmayan_be_e = Convert.ToInt32(gk_be_e_txt.Text);
                s_m_olmayan_ant_e = Convert.ToInt32(gk_ant_e_txt.Text);
                s_m_olmayan_rek_e = Convert.ToInt32(gk_rek_e_txt.Text);
                s_m_olmayan_be_k = Convert.ToInt32(gk_be_k_txt.Text);
                s_m_olmayan_ant_k = Convert.ToInt32(gk_ant_k_txt.Text);
                s_m_olmayan_rek_k = Convert.ToInt32(gk_rek_k_txt.Text);
                m_olmayan_be_e = Convert.ToInt32(gk_be_e_txt.Text);
                m_olmayan_ant_e = Convert.ToInt32(gk_ant_e_txt.Text);
                m_olmayan_rek_e = Convert.ToInt32(gk_ant_e_txt.Text);
                m_olmayan_be_k = Convert.ToInt32(gk_be_k_txt.Text);
                m_olmayan_ant_k = Convert.ToInt32(gk_ant_k_txt.Text);
                m_olmayan_rek_k = Convert.ToInt32(gk_rek_k_txt.Text);
                Adayları_Excellere_Ayır();
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_BE_KADIN.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_ANT_KADIN.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar\Milliolmayan_REK_KADIN.xls");
                yerlerstir();
            }
        }
        private void gk_ant_k_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar >= 48 && (int)e.KeyChar <= 57)
            {
                e.Handled = false;//eğer rakamsa  yazdır.
            }

            else if ((int)e.KeyChar == 8)
            {
                e.Handled = false;//eğer basılan tuş backspace ise yazdır.
            }
            else
            {
                e.Handled = true;//bunların dışındaysa hiçbirisini yazdırma
            }
        }
        void baslangıc_aday_kontenjan_kontrol()
        {  
            /* yerleştirme işlemine başlamadan önce ayrıştırılmış adayların tutulduğu excel dosyasından başvuran aday sayılarına bakılır ve kontenjanlarla kıyaslanır
             yeterli başvuru yoksa veya bir bölüme kimse başvurmamışsa cinsiyetler arası kontenjan aktarımı yapılır.*/

            int be_erkek_liste_kisi_sayisi = 0;
            int ant_erkek_liste_kisi_sayisi = 0;
            int rek_erkek_liste_kisi_sayisi = 0;
            int be_kadın_liste_kisi_sayisi = 0;
            int ant_kadın_liste_kisi_sayisi = 0;
            int rek_kadın_liste_kisi_sayisi = 0;

            if (baslangıc_aday_kontenjan_kontrol_sayac == 0)
            {
                DataTable dt = new DataTable();
                progressBar1.Value = 0;
                progressBar1.Maximum = dizi.Length;
                for (int i = 0; i < dizi.Length; i++)
                {
                    progressBar1.Value++;
                    OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    baglanti.Open();
                    dt.Clear();
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", baglanti);
                    da.Fill(dt);
                    baglanti.Close();
                    if (dizi[i] == Milliolmayan_BE_ERKEK) { be_erkek_liste_kisi_sayisi = dt.Rows.Count; }
                    else if (dizi[i] == Milliolmayan_BE_KADIN) { be_kadın_liste_kisi_sayisi = dt.Rows.Count; }
                    else if (dizi[i] == Milliolmayan_ANT_ERKEK) { ant_erkek_liste_kisi_sayisi = dt.Rows.Count; }
                    else if (dizi[i] == Milliolmayan_ANT_KADIN) { ant_kadın_liste_kisi_sayisi = dt.Rows.Count; }
                    else if (dizi[i] == Milliolmayan_REK_ERKEK) { rek_erkek_liste_kisi_sayisi = dt.Rows.Count; }
                    else if (dizi[i] == Milliolmayan_REK_KADIN) { rek_kadın_liste_kisi_sayisi = dt.Rows.Count; }
                }

                if (be_erkek_liste_kisi_sayisi == 0 && be_kadın_liste_kisi_sayisi == 0)
                {
                    m_olmayan_be_e = 0;
                    s_m_olmayan_be_e = 0;      //HEM KADIN HEM ERKEK BEDEN EĞİTİMİNE BAŞVURAN HİÇ ADAY YOKSA KONTEJANLARI SIFIRLA
                    m_olmayan_be_k = 0;
                    s_m_olmayan_be_k = 0;
                }
                else if (be_kadın_liste_kisi_sayisi < s_m_olmayan_be_k && be_erkek_liste_kisi_sayisi < s_m_olmayan_be_e)
                {
                    m_olmayan_be_e = be_erkek_liste_kisi_sayisi;
                    s_m_olmayan_be_e = be_erkek_liste_kisi_sayisi;
                    m_olmayan_be_k = be_kadın_liste_kisi_sayisi;  // HEM KADIN HEM ERKEK BEDEN EĞİTİMİNE BAŞVURAN ADAY SAYISI KONTENJANLARDAN DÜŞÜKSE KONTENJANLARI BAŞVURAN SAYISI KADAR GÜNCELLE
                    s_m_olmayan_be_k = be_kadın_liste_kisi_sayisi;
                }
                else if (be_kadın_liste_kisi_sayisi < s_m_olmayan_be_k)
                {
                    m_olmayan_be_e = m_olmayan_be_e + (s_m_olmayan_be_k - be_kadın_liste_kisi_sayisi);
                    s_m_olmayan_be_e = m_olmayan_be_e;
                    m_olmayan_be_k = be_kadın_liste_kisi_sayisi;  // BEDEN EĞİTİMİ KADIN BAŞVURAN SAYISI KONTENJANDAN KÜÇÜKSE AÇIKTA KALAN KONTENJANI ERKEKLERE EKLE
                    s_m_olmayan_be_k = m_olmayan_be_k;
                }
                else if (be_erkek_liste_kisi_sayisi < s_m_olmayan_be_e)
                {
                    m_olmayan_be_k = m_olmayan_be_k + (s_m_olmayan_be_e - be_erkek_liste_kisi_sayisi);
                    s_m_olmayan_be_k = m_olmayan_be_k;
                    m_olmayan_be_e = be_erkek_liste_kisi_sayisi;    // BEDEN EĞİTİMİ ERKEK BAŞVURAN SAYISI KONTENJANDAN KÜÇÜKSE AÇIKTA KALAN KONTENJANI KADINLARA EKLE
                    s_m_olmayan_be_e = be_erkek_liste_kisi_sayisi;
                }

                if (ant_erkek_liste_kisi_sayisi == 0 && ant_kadın_liste_kisi_sayisi == 0)
                {
                    m_olmayan_ant_e = 0;
                    s_m_olmayan_ant_e = 0;
                    m_olmayan_ant_k = 0;
                    s_m_olmayan_ant_k = 0;
                }
                else if (ant_kadın_liste_kisi_sayisi < s_m_olmayan_ant_k && ant_erkek_liste_kisi_sayisi < s_m_olmayan_ant_e)
                {
                    m_olmayan_ant_e = ant_erkek_liste_kisi_sayisi;
                    s_m_olmayan_ant_e = ant_erkek_liste_kisi_sayisi;
                    m_olmayan_ant_k = ant_kadın_liste_kisi_sayisi;
                    s_m_olmayan_ant_k = ant_kadın_liste_kisi_sayisi;
                }
                else if (ant_kadın_liste_kisi_sayisi < s_m_olmayan_ant_k)
                {
                    m_olmayan_ant_e = m_olmayan_ant_e + (s_m_olmayan_ant_k - ant_kadın_liste_kisi_sayisi);
                    s_m_olmayan_ant_e = m_olmayan_ant_e;
                    m_olmayan_ant_k = ant_kadın_liste_kisi_sayisi;
                    s_m_olmayan_ant_k = m_olmayan_ant_k;
                }
                else if (ant_erkek_liste_kisi_sayisi < s_m_olmayan_ant_e)
                {
                    m_olmayan_ant_k = m_olmayan_ant_k + (s_m_olmayan_ant_e - ant_erkek_liste_kisi_sayisi);
                    s_m_olmayan_ant_k = m_olmayan_ant_k;
                    m_olmayan_ant_e = ant_erkek_liste_kisi_sayisi;
                    s_m_olmayan_ant_e = ant_erkek_liste_kisi_sayisi;
                }

                if (rek_erkek_liste_kisi_sayisi == 0 && rek_kadın_liste_kisi_sayisi == 0)
                {
                    m_olmayan_rek_e = 0;
                    s_m_olmayan_rek_e = 0;
                    m_olmayan_rek_k = 0;
                    s_m_olmayan_rek_k = 0;
                }
                else if (rek_kadın_liste_kisi_sayisi < s_m_olmayan_rek_k && rek_erkek_liste_kisi_sayisi < s_m_olmayan_rek_e)
                {
                    m_olmayan_rek_e = rek_erkek_liste_kisi_sayisi;
                    s_m_olmayan_rek_e = rek_erkek_liste_kisi_sayisi;
                    m_olmayan_rek_k = rek_kadın_liste_kisi_sayisi;
                    s_m_olmayan_rek_k = rek_kadın_liste_kisi_sayisi;
                }
                else if (rek_kadın_liste_kisi_sayisi < s_m_olmayan_rek_k)
                {
                    m_olmayan_rek_e = m_olmayan_rek_e + (s_m_olmayan_rek_k - rek_kadın_liste_kisi_sayisi);
                    s_m_olmayan_rek_e = m_olmayan_rek_e;
                    m_olmayan_rek_k = rek_kadın_liste_kisi_sayisi;
                    s_m_olmayan_rek_k = m_olmayan_rek_k;
                }
                else if (rek_erkek_liste_kisi_sayisi < s_m_olmayan_rek_e)
                {
                    m_olmayan_rek_k = m_olmayan_rek_k + (s_m_olmayan_rek_e - rek_erkek_liste_kisi_sayisi);
                    s_m_olmayan_rek_k = m_olmayan_rek_k;
                    m_olmayan_rek_e = rek_erkek_liste_kisi_sayisi;
                    s_m_olmayan_rek_e = rek_erkek_liste_kisi_sayisi;
                }

            }
        }
        void yerlerstir()
        {
            /* hangi kontenjanın dikkate alınacağının bilinmesi için arka planda açık olan excel dosyasın dosya yolu kontorol edilir.
             dosya yolu belirlendikten sonra kontenjan sayısına bakılır. yerleşme durumu " " (boş) durumda olan adaylar kontenjanlar kadar yerleşme durumu "Asil" olarak güncellenir.*/

            baslangıc_aday_kontenjan_kontrol();
            baslangıc_aday_kontenjan_kontrol_sayac = 1;
            progressBar1.Value = 0;
            progressBar1.Maximum = dizi.Length;
            for (int i = 0; i < dizi.Length; i++)
            {
                progressBar1.Value++;
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    string kayıt_no = dt.Rows[j][0].ToString();

                    if (dizi[i] == Milliolmayan_BE_ERKEK && m_olmayan_be_e > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'"; 
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_be_e--;

                    }
                    if (dizi[i] == Milliolmayan_BE_KADIN && m_olmayan_be_k > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_be_k--;

                    }
                    if (dizi[i] == Milliolmayan_ANT_KADIN && m_olmayan_ant_k > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_ant_k--;

                    }
                    if (dizi[i] == Milliolmayan_ANT_ERKEK && m_olmayan_ant_e > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_ant_e--;

                    }
                    if (dizi[i] == Milliolmayan_REK_KADIN && m_olmayan_rek_k > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_rek_k--;

                    }
                    if (dizi[i] == Milliolmayan_REK_ERKEK && m_olmayan_rek_e > 0 && dt.Rows[j][31].ToString() == "")
                    {
                        OleDbCommand komut = new OleDbCommand();
                        nbaglanti.Open();
                        komut.Connection = nbaglanti;
                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='ASİL' WHERE [Kayıt No]='" + kayıt_no + "'";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        nbaglanti.Close();
                        m_olmayan_rek_e--;

                    }
                }
            }
            adım2();
        }
        void adım2()
        {  // yerleşme durumu asil olarak güncellenen adaylar diğer tablolarda yerleşip yerleşmediği kontrol edilir. yerlerşmişse  büyük olan tercihindeki yerleşme durumu "Diğer Liste " olarak güncellenir. 
            progressBar1.Value = 0;
            progressBar1.Maximum = dizi.Length;
            for (int i = 0; i < dizi.Length; i++)
            {
                progressBar1.Value++;
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    string kayıt_no = dt.Rows[j][0].ToString();
                    string tercih_be1 = dt.Rows[j][20].ToString();
                    string tercih_ant1 = dt.Rows[j][21].ToString();
                    string tercih_rek1 = dt.Rows[j][22].ToString();
                    string secili_tercih1 = "";
                    string sonuc1 = dt.Rows[j][31].ToString();
                    if (dizi[i] == Milliolmayan_BE_ERKEK || dizi[i] == Milliolmayan_BE_KADIN) { secili_tercih1 = tercih_be1; }
                    if (dizi[i] == Milliolmayan_ANT_ERKEK || dizi[i] == Milliolmayan_ANT_KADIN) { secili_tercih1 = tercih_ant1; }
                    if (dizi[i] == Milliolmayan_REK_ERKEK || dizi[i] == Milliolmayan_REK_KADIN) { secili_tercih1 = tercih_rek1; }
                    if (sonuc1 == "ASİL")
                    {

                        for (int k = 0; k < dizi.Length; k++)
                        {
                            if (i != k)
                            {
                                OleDbConnection nnbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[k] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                                DataTable dtt = new DataTable();
                                int ii = 0;
                                nnbaglanti.Open();
                                dtt.Clear();
                                OleDbCommand komutt = new OleDbCommand("SELECT count([Kayıt No]) FROM [KAYIT-GİRİŞ$] WHERE [Kayıt No]='" + kayıt_no + "'", nnbaglanti);
                                komutt.Connection = nnbaglanti;
                                ii = Convert.ToInt32(komutt.ExecuteScalar());
                                nnbaglanti.Close();
                                if (ii > 0)
                                {
                                    nnbaglanti.Open();
                                    OleDbDataAdapter daa = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] WHERE [Kayıt No]='" + kayıt_no + "'", nnbaglanti);
                                    dtt.Clear();
                                    daa.Fill(dtt);
                                    nnbaglanti.Close();
                                    string tercih_be = dtt.Rows[0][20].ToString();
                                    string tercih_ant = dtt.Rows[0][21].ToString();
                                    string tercih_rek = dtt.Rows[0][22].ToString();
                                    string secili_tercih2 = "";
                                    if (dizi[k] == Milliolmayan_BE_ERKEK || dizi[k] == Milliolmayan_BE_KADIN) { secili_tercih2 = tercih_be; }
                                    if (dizi[k] == Milliolmayan_ANT_ERKEK || dizi[k] == Milliolmayan_ANT_KADIN) { secili_tercih2 = tercih_ant; }
                                    if (dizi[k] == Milliolmayan_REK_ERKEK || dizi[k] == Milliolmayan_REK_KADIN) { secili_tercih2 = tercih_rek; }
                                    string sonuc2 = dtt.Rows[0][31].ToString();

                                    if (Convert.ToInt32(secili_tercih1) == 1)
                                    {
                                        OleDbCommand komut = new OleDbCommand();
                                        nnbaglanti.Open();
                                        komut.Connection = nnbaglanti;
                                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='DİĞER LİSTEDE YERLEŞEN' WHERE [Kayıt No]='" + kayıt_no + "'";
                                        komut.CommandText = sql;
                                        komut.ExecuteNonQuery();  // 
                                        nnbaglanti.Close();
                                    }
                                    if (Convert.ToInt32(secili_tercih1) == 2 && Convert.ToInt32(secili_tercih2) == 3)
                                    {
                                        OleDbCommand komut = new OleDbCommand();
                                        nnbaglanti.Open();
                                        komut.Connection = nnbaglanti;
                                        string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='DİĞER LİSTEDE YERLEŞEN' WHERE [Kayıt No]='" + kayıt_no + "'";
                                        komut.CommandText = sql;
                                        komut.ExecuteNonQuery();
                                        nnbaglanti.Close();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            MilliOlmayanKontrol();
        }
        void MilliOlmayanKontrol()
        { /* 2. adımdan sonra  yerleşme durumu asil olan adaylar sayılır. ve kontenjanlardan çıkarılı ve kontenjanlar güncellenmiş olur.
           * daha sonra asil adaylar sayılır, diğerlistede yerleşenler sayılır toplanır ve listedeki kişi sayısından çıkarılır sonuç "0" sıkarsa listede yerleşebilecek aday kalmadığından cinsiyetler arası kontenjan aktarımı yapılır. */
            int be_erkek_yerleşen = 0;
            int be_erkek_diger_liste_yerlesen = 0;
            int be_erkek_listedeki_kisi_sayisi = 0;
            int be_erkek_yerleşemeyen = 0;

            int ant_erkek_yerleşen = 0;
            int ant_erkek_diger_liste_yerlesen = 0;
            int ant_erkek_yerleşemeyen = 0;
            int ant_erkek_listedeki_kisi_sayisi = 0;

            int rek_erkek_yerleşen = 0;
            int rek_erkek_diger_liste_yerlesen = 0;
            int rek_erkek_yerleşemeyen = 0;
            int rek_erkek_listedeki_kisi_sayisi = 0;

            int be_kadın_yerleşen = 0;
            int be_kadın_diger_liste_yerlesen = 0;
            int be_kadın_yerleşemeyen = 0;
            int be_kadın_listedeki_kisi_sayisi = 0;

            int ant_kadın_yerleşen = 0;
            int ant_kadın_diger_liste_yerlesen = 0;
            int ant_kadın_yerleşemeyen = 0;
            int ant_kadın_listedeki_kisi_sayisi = 0;

            int rek_kadın_yerleşen = 0;
            int rek_kadın_diger_liste_yerlesen = 0;
            int rek_kadın_yerleşemeyen = 0;
            int rek_kadın_listedeki_kisi_sayisi = 0;

            progressBar1.Value = 0;
            progressBar1.Maximum = dizi.Length;
            for (int i = 0; i < dizi.Length; i++)
            {
                progressBar1.Value++;
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT count(SONUÇ) FROM [KAYIT-GİRİŞ$] WHERE SONUÇ='ASİL'", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();
                int yerlesen_sayisi = Convert.ToInt32(dt.Rows[0][0].ToString());

                DataTable dts = new DataTable();
                nbaglanti.Open();
                dts.Clear();
                OleDbDataAdapter da1 = new OleDbDataAdapter("SELECT count(SONUÇ) FROM [KAYIT-GİRİŞ$] WHERE SONUÇ='DİĞER LİSTEDE YERLEŞEN'", nbaglanti);
                da1.Fill(dts);
                nbaglanti.Close();
                int diger_liste_yerlesen = Convert.ToInt32(dts.Rows[0][0].ToString());

                DataTable dtss = new DataTable();
                nbaglanti.Open();
                dtss.Clear();
                OleDbDataAdapter da2 = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", nbaglanti);
                da2.Fill(dtss);
                nbaglanti.Close();
                int listedeki_kisi_sayisi = Convert.ToInt32(dtss.Rows.Count.ToString());

                if (dizi[i] == Milliolmayan_BE_ERKEK)
                {
                    be_erkek_yerleşen = yerlesen_sayisi;
                    be_erkek_diger_liste_yerlesen = diger_liste_yerlesen;
                    be_erkek_listedeki_kisi_sayisi = listedeki_kisi_sayisi;
                    be_erkek_yerleşemeyen = be_erkek_listedeki_kisi_sayisi - (be_erkek_yerleşen + be_erkek_diger_liste_yerlesen);
                    m_olmayan_be_e = s_m_olmayan_be_e - yerlesen_sayisi;
                }
                else if (dizi[i] == Milliolmayan_ANT_ERKEK)
                {
                    ant_erkek_yerleşen = yerlesen_sayisi;
                    ant_erkek_diger_liste_yerlesen = diger_liste_yerlesen;
                    ant_erkek_listedeki_kisi_sayisi = listedeki_kisi_sayisi;
                    ant_erkek_yerleşemeyen = ant_erkek_listedeki_kisi_sayisi - (ant_erkek_yerleşen + ant_erkek_diger_liste_yerlesen);
                    m_olmayan_ant_e = s_m_olmayan_ant_e - yerlesen_sayisi;
                }
                else if (dizi[i] == Milliolmayan_REK_ERKEK)
                {
                    rek_erkek_yerleşen = yerlesen_sayisi;
                    rek_erkek_diger_liste_yerlesen = diger_liste_yerlesen;
                    rek_erkek_listedeki_kisi_sayisi = listedeki_kisi_sayisi;     // BU KISIMDA ASİL OLARAK KAÇ ADAY YERLEŞMİŞ KONTROL EDİLİR VE KONTENJANLAR GÜNCELLENİR BÜTÜN KONTENJANLAR DOLDUĞUNDA İŞLEME SON VERİLİR
                    rek_erkek_yerleşemeyen = rek_erkek_listedeki_kisi_sayisi - (rek_erkek_yerleşen + rek_erkek_diger_liste_yerlesen);
                    m_olmayan_rek_e = s_m_olmayan_rek_e - yerlesen_sayisi;
                }
                else if (dizi[i] == Milliolmayan_BE_KADIN)
                {
                    be_kadın_yerleşen = yerlesen_sayisi;
                    be_kadın_diger_liste_yerlesen = diger_liste_yerlesen;
                    be_kadın_listedeki_kisi_sayisi = listedeki_kisi_sayisi;
                    be_kadın_yerleşemeyen = be_kadın_listedeki_kisi_sayisi - (be_kadın_yerleşen + be_kadın_diger_liste_yerlesen);
                    m_olmayan_be_k = s_m_olmayan_be_k - yerlesen_sayisi;
                }
                else if (dizi[i] == Milliolmayan_ANT_KADIN)
                {
                    ant_kadın_yerleşen = yerlesen_sayisi;
                    ant_kadın_diger_liste_yerlesen = diger_liste_yerlesen;
                    ant_kadın_listedeki_kisi_sayisi = listedeki_kisi_sayisi;
                    ant_kadın_yerleşemeyen = ant_kadın_listedeki_kisi_sayisi - (ant_kadın_yerleşen + ant_kadın_diger_liste_yerlesen);
                    m_olmayan_ant_k = s_m_olmayan_ant_k - yerlesen_sayisi;
                }
                else if (dizi[i] == Milliolmayan_REK_KADIN)
                {
                    rek_kadın_yerleşen = yerlesen_sayisi;
                    rek_kadın_diger_liste_yerlesen = diger_liste_yerlesen;
                    rek_kadın_listedeki_kisi_sayisi = listedeki_kisi_sayisi;
                    rek_kadın_yerleşemeyen = rek_kadın_listedeki_kisi_sayisi - (rek_kadın_yerleşen + rek_kadın_diger_liste_yerlesen);
                    m_olmayan_rek_k = s_m_olmayan_rek_k - yerlesen_sayisi;
                }
            }
            if (be_erkek_yerleşemeyen == 0 && be_kadın_yerleşemeyen == 0)
            {
                m_olmayan_be_e = 0;
                s_m_olmayan_be_e = 0;
                m_olmayan_be_k = 0;
                s_m_olmayan_be_k = 0;
            }
            else if (be_erkek_yerleşemeyen != 0 && be_kadın_yerleşemeyen == 0)
            {
                m_olmayan_be_e = m_olmayan_be_e + m_olmayan_be_k;
                s_m_olmayan_be_e = s_m_olmayan_be_e + m_olmayan_be_k;
                m_olmayan_be_k = 0;
                s_m_olmayan_be_k = 0;
            }
            else if (be_erkek_yerleşemeyen == 0 && be_kadın_yerleşemeyen != 0)
            {
                m_olmayan_be_k = m_olmayan_be_e + m_olmayan_be_k;
                s_m_olmayan_be_k = s_m_olmayan_be_k + m_olmayan_be_e;
                m_olmayan_be_e = 0;
                s_m_olmayan_be_e = 0;
            }


            if (ant_erkek_yerleşemeyen == 0 && ant_kadın_yerleşemeyen == 0)
            {
                m_olmayan_ant_e = 0;
                s_m_olmayan_ant_e = 0;
                m_olmayan_ant_k = 0;
                s_m_olmayan_ant_k = 0;
            }
            else if (ant_erkek_yerleşemeyen != 0 && ant_kadın_yerleşemeyen == 0)
            {
                m_olmayan_ant_e = m_olmayan_ant_e + m_olmayan_ant_k;
                s_m_olmayan_ant_e = s_m_olmayan_ant_e + m_olmayan_ant_k;
                m_olmayan_ant_k = 0;
                s_m_olmayan_ant_k = 0;
            }
            else if (ant_erkek_yerleşemeyen == 0 && ant_kadın_yerleşemeyen != 0)
            {
                m_olmayan_ant_k = m_olmayan_ant_e + m_olmayan_ant_k;
                s_m_olmayan_ant_k = s_m_olmayan_ant_k + m_olmayan_ant_e;
                m_olmayan_ant_e = 0;
                s_m_olmayan_ant_e = 0;
            }

            if (rek_erkek_yerleşemeyen == 0 && rek_kadın_yerleşemeyen == 0)
            {
                m_olmayan_rek_e = 0;
                s_m_olmayan_rek_e = 0;
                m_olmayan_rek_k = 0;
                s_m_olmayan_rek_k = 0;
            }
            else if (rek_erkek_yerleşemeyen != 0 && rek_kadın_yerleşemeyen == 0)
            {
                m_olmayan_rek_e = m_olmayan_rek_e + m_olmayan_rek_k;
                s_m_olmayan_rek_e = s_m_olmayan_rek_e + m_olmayan_rek_k;
                m_olmayan_rek_k = 0;
                s_m_olmayan_rek_k = 0;
            }
            else if (rek_erkek_yerleşemeyen == 0 && rek_kadın_yerleşemeyen != 0)
            {
                m_olmayan_rek_k = m_olmayan_rek_e + m_olmayan_rek_k;
                s_m_olmayan_rek_k = s_m_olmayan_rek_k + m_olmayan_rek_e;
                m_olmayan_rek_e = 0;
                s_m_olmayan_rek_e = 0;
            }


            if (m_olmayan_be_e <= 0 && m_olmayan_ant_e <= 0 && m_olmayan_rek_e <= 0 && m_olmayan_be_k <= 0 && m_olmayan_ant_k <= 0 && m_olmayan_rek_k <= 0) { yedekleri_duzenle(); }
            else yerlerstir();
        }
        void yedekleri_duzenle()
        { 
            /* asil adayların yerleştirme işlemi bittikten sonra kontenjanların 3 katı kadar aday yedek olakak güncellenir (burada kriter yerleşme durumunun "" (boş) olmasıdır) 
             eğer 3 katı kadar aday yoksa geriye kaç kişi kalmışsa o kadar aday yedek olarak gösterilir.*/

            progressBar1.Value = 0;
            progressBar1.Maximum = dizi.Length;
            for (int i = 0; i < dizi.Length; i++)
            {
                progressBar1.Value++;
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();
                int sayac = 0;
                if (dizi[i] == Milliolmayan_BE_ERKEK) { sayac = Convert.ToInt32(gk_be_e_txt.Text) * 3; }
                else if (dizi[i] == Milliolmayan_BE_KADIN) { sayac = Convert.ToInt32(gk_be_k_txt.Text) * 3; }
                else if (dizi[i] == Milliolmayan_ANT_ERKEK) { sayac = Convert.ToInt32(gk_ant_e_txt.Text) * 3; }
                else if (dizi[i] == Milliolmayan_ANT_KADIN) { sayac = Convert.ToInt32(gk_ant_k_txt.Text) * 3; }
                else if (dizi[i] == Milliolmayan_REK_ERKEK) { sayac = Convert.ToInt32(gk_rek_e_txt.Text) * 3; }
                else if (dizi[i] == Milliolmayan_REK_KADIN) { sayac = Convert.ToInt32(gk_rek_k_txt.Text) * 3; }

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    string kayıt_no = dt.Rows[j][0].ToString();
                    string sonuç = dt.Rows[j][31].ToString();
                    if (sayac > 0)
                    {
                        if (sonuç == "")
                        {
                            OleDbCommand komut = new OleDbCommand();
                            nbaglanti.Open();
                            komut.Connection = nbaglanti;
                            string sql = "Update [KAYIT-GİRİŞ$] set SONUÇ='YEDEK' WHERE [Kayıt No]='" + kayıt_no + "'"; // BU KISIMDA ASİL KONTENJANIN 3 KATI KADAR ADAY YEDEK OLARAK LİSTEYE ALININ EĞER YETERLİ KİŞİ SAYISINA ULAŞILAMADIYSA BOŞTA KALAN ADAY SAYISI KADAR YEDEK LİSTEYE ALINIR.
                            komut.CommandText = sql;
                            komut.ExecuteNonQuery();
                            nbaglanti.Close();
                            sayac--;
                        }
                    }
                }
            }
            yerlesmis_aday_tablolarına_yedek_asil_ayrıstır();
        }
        void yerlesmis_aday_tablolarına_yedek_asil_ayrıstır()
        { 
            // bu kısımda asil ve yedek olarak yerlerşen adaylar yayınlanmak için belirlenen şablona göre excel listelerine ayrılır.
            
            progressBar1.Value = 0;
            progressBar1.Maximum = dizi.Length;
            for (int i = 0; i < dizi.Length; i++)
            {
                progressBar1.Value++;
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] WHERE SONUÇ='ASİL'", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();

                if (dizi[i] == Milliolmayan_BE_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][13].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                                                          
                }
                else if (dizi[i] == Milliolmayan_BE_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_ASİL_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][13].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_ANT_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][14].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_ANT_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_ASİL_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][14].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_REK_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][15].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_REK_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_ASİL_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][15].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
            }

            for (int i = 0; i < dizi.Length; i++)
            {
                OleDbConnection nbaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dizi[i] + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                DataTable dt = new DataTable();
                nbaglanti.Open();
                dt.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] WHERE SONUÇ='YEDEK'", nbaglanti);
                da.Fill(dt);
                nbaglanti.Close();

                if (dizi[i] == Milliolmayan_BE_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][13].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }

                }
                else if (dizi[i] == Milliolmayan_BE_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_BE_YEDEK_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][13].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_ANT_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][14].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_ANT_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_ANT_YEDEK_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][14].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_REK_ERKEK)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_ERKEK.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][15].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
                else if (dizi[i] == Milliolmayan_REK_KADIN)
                {
                    string Yol = "C:\\Özel Yetenek\\Yerleştirilmiş Aday Tabloları\\Milli Olmayanlar\\Milliolmayan_REK_YEDEK_KADIN.xls";
                    OleDbConnection ybaglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Yol + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int sıra_no = j + 1;
                        OleDbCommand komut = new OleDbCommand();
                        ybaglanti.Open();
                        komut.Connection = ybaglanti;
                        string sql = "Insert into [Aday Tabloları$] ([SIRA NO],[KAYIT NO],[ADI],[SOYADI],[ALAN/KOL],[TYT-P],[OBP],[SPOR ÖZGEÇMİŞ BELGE TÜRÜ],[SPOR ÖZGEÇMİŞ KATSAYISI],[BİR ÖNCEKİ SENE YERLEŞME DURUMU],[SÖP],[ÖYSP-SP],[YP],[SONUÇ]) values('" + sıra_no.ToString() + "','" + dt.Rows[j][0].ToString() + "','" + dt.Rows[j][2].ToString() + "','" + dt.Rows[j][3].ToString() + "','" + dt.Rows[j][6].ToString() + "','" + dt.Rows[j][7].ToString() + "','" + dt.Rows[j][8].ToString() + "','" + dt.Rows[j][10].ToString() + "','" + dt.Rows[j][15].ToString() + "','" + dt.Rows[j][19].ToString() + "','" + dt.Rows[j][26].ToString() + "','" + dt.Rows[j][28].ToString() + "','" + dt.Rows[j][29].ToString() + "','" + dt.Rows[j][31].ToString() + "')";
                        komut.CommandText = sql;
                        komut.ExecuteNonQuery();
                        ybaglanti.Close();
                    }
                }
            }
            progressBar1.Visible = false;
            MessageBox.Show("!!!Yerleştirme İşlemleri Başarıyla Tamamlandı!!!");
        }

        private void button2_Click(object sender, EventArgs e)
        { 
            // milli adayları ayrıştırma butonu
           
            textbox_kontrol();
            if (Textbox_kontrol_sayacı == 0)
            {
                s_m_olmayan_be_e = Convert.ToInt32(gk_be_e_txt.Text);
                s_m_olmayan_ant_e = Convert.ToInt32(gk_ant_e_txt.Text);
                s_m_olmayan_rek_e = Convert.ToInt32(gk_rek_e_txt.Text);
                s_m_olmayan_be_k = Convert.ToInt32(gk_be_k_txt.Text);
                s_m_olmayan_ant_k = Convert.ToInt32(gk_ant_k_txt.Text);
                s_m_olmayan_rek_k = Convert.ToInt32(gk_rek_k_txt.Text);
                m_olmayan_be_e = Convert.ToInt32(gk_be_e_txt.Text);
                m_olmayan_ant_e = Convert.ToInt32(gk_ant_e_txt.Text);
                m_olmayan_rek_e = Convert.ToInt32(gk_ant_e_txt.Text);
                m_olmayan_be_k = Convert.ToInt32(gk_be_k_txt.Text);
                m_olmayan_ant_k = Convert.ToInt32(gk_ant_k_txt.Text);
                m_olmayan_rek_k = Convert.ToInt32(gk_rek_k_txt.Text);
                milli_adayları_ayır();
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_ERKEK.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_BE_KADIN.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_ANT_KADIN.xls");
                PUAN_hesapla(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler\Milli_REK_KADIN.xls");
                MessageBox.Show("Milli Adayları Ayrıştırma ve Puanlarını Hesaplama Tamamlandı.");
            }
        }
    }
}
