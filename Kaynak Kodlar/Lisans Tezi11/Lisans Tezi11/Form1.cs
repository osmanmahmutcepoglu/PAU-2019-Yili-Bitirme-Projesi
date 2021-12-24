using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Lisans_Tezi11
{
    public partial class Form1 : Form
    {
        ErrorProvider provider = new ErrorProvider();
        int sayac = 0;  // kayıt işlemi gerçekleştirilirken textbox ve comboboxların kontrol bayrağı.
        int millisayac = 0;
        string deger = "";
        int kayıtkontrolsayacı = 0; // metin belgesinde tuttuğumuz kayıt numarasının aktarıldığı değişken.
        int sayac_kacar_artacak;
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls; Extended Properties='Excel 12.0 xml;HDR=YES;'");
        DataTable tablo = new DataTable();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Exceldosyasıvarmı();  // başlangıçta ana excel dosyasından okuma işlemi. metin belgelerini okuma işlemleri burada gerçekleştirilir.
            txtkontrol();         
            textboxsayackkontrol();
            textBox13.Visible = false;
            datagrid_veri_cekme();
            button3.Enabled = false;
            button4.Enabled = false;
            button7.Enabled = false;
            maskedTextBox3.Enabled = false;
            maskedTextBox4.Enabled = false;
            maskedTextBox5.Enabled = false;
            maskedTextBox6.Enabled = false;
            maskedTextBox7.Enabled = false;
            maskedTextBox8.Enabled = false;
            comboBox9.Enabled = false;
            button5.Visible = false;
            comboBox4_MouseClick(null, null);
            comboBox7_MouseClick(null, null);
            label26.Visible = false;
            textBox14.Visible = false;
        }
        void ogrencibilgikontrol() // Aday bilgileri girilirken istenilen tipte olması için gerekli kontoroller burada yapılır.
        {
            sayac = 0;
            provider.BlinkRate = 0;
            provider.Clear();
            if (radioButton1.Checked == false && radioButton2.Checked == false) { MessageBox.Show("Lütfen Kayıt Sayacının Kaçar Kaçar Artacağını Seçiniz.."); sayac++; }
            if (comboBox1.Text == "") { provider.SetError(comboBox1, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox2.Text == "") { provider.SetError(comboBox2, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox3.Text == "") { provider.SetError(comboBox3, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox4.Text == "") { provider.SetError(comboBox4, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox5.Text == "") { provider.SetError(comboBox5, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox5.SelectedIndex == 0 && comboBox9.SelectedIndex == -1) { provider.SetError(comboBox9, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox6.Text == "") { provider.SetError(comboBox6, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox7.Text == "") { provider.SetError(comboBox7, "Bu alan boş geçilemez"); sayac++; }
            //if (comboBox8.Text == "") { provider.SetError(comboBox8, "Bu alan boş geçilemez"); sayac++; }
            if (textBox1.Text == "") { provider.SetError(textBox1, "Bu alan boş geçilemez"); sayac++; }
            if (textBox2.Text == "") { provider.SetError(textBox2, "Bu alan boş geçilemez"); sayac++; }
            if (textBox3.Text == "") { provider.SetError(textBox3, "Bu alan boş geçilemez"); sayac++; }
            if (textBox4.Text == "") { provider.SetError(textBox4, "Bu alan boş geçilemez"); sayac++; }
            if (textBox5.Text == "") { provider.SetError(textBox5, "Kayıt Başlangıç Numarasını Giriniz."); sayac++; }
            if (maskedTextBox1.Text == "   ,") { provider.SetError(maskedTextBox1, "Bu alan boş geçilemez"); sayac++; }
            if (maskedTextBox2.Text == "   ,") { provider.SetError(maskedTextBox2, "Bu alan boş geçilemez"); sayac++; }
            if (comboBox8.SelectedIndex == 1 && textBox13.Text == "") { provider.SetError(textBox13, "Bu alan boş geçilemez"); sayac++; }
            else if (comboBox8.SelectedIndex == 1 && Convert.ToInt32(textBox13.Text) > 100) { provider.SetError(textBox13, "100'den büyük bir değer giremezsiniz."); sayac++; }
            if (sayac == 0) { Int64 TC = Convert.ToInt64(textBox1.Text); if (TC < 10000000000) { sayac++; provider.SetError(textBox1, "Geçerli Bir TC Kimlik Numarası Giriniz.."); } }
            if (sayac == 0) { int tarih = Convert.ToInt32(textBox4.Text); DateTime dt = DateTime.Today; int yil = dt.Year; if (tarih < 1950 || tarih > (yil - 16)) { sayac++; provider.SetError(textBox4, "Geçerli Bir Tarih Giriniz.."); } }
            if (sayac == 0)
            {
                string a = maskedTextBox1.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 10) { provider.SetError(maskedTextBox1, "Eksik veya Hatalı Giriş Yaptınız."); sayac++; }
            }
            if (sayac == 0)
            {
                string a = maskedTextBox2.Text;
                int k = 0;
                for (int i = 0; i < a.Length; i++)
                {
                    if (a.Substring(i, 1) == " ") k++;
                }
                if (k >= 1 || a.Length < 10) { provider.SetError(maskedTextBox2, "Eksik veya Hatalı Giriş Yaptınız."); sayac++; }
            }
            if (sayac == 0) { if (textBox10.Text == "" && textBox11.Text == "" && textBox12.Text == "") { MessageBox.Show("Öğrencinin Başvurduğu Programlar'ın Tercih Değerlerini Giriniz!!"); sayac++; } }
            if (sayac == 0)
            {
                if (textBox10.Text == "1" && textBox11.Text == "1") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox10.Text == "1" && textBox12.Text == "1") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox12.Text == "1" && textBox11.Text == "1") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox10.Text == "2" && textBox11.Text == "2") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox10.Text == "2" && textBox12.Text == "2") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox12.Text == "2" && textBox11.Text == "2") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox10.Text == "3" && textBox11.Text == "3") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox10.Text == "3" && textBox12.Text == "3") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
                else if (textBox12.Text == "3" && textBox11.Text == "3") { MessageBox.Show("!!Öğrencinin Başvurduğu Programlar'ın Tercih Değerleri Aynı Olamaz!!"); sayac++; }
            }
            if (maskedTextBox3.Enabled == true) { if (maskedTextBox3.Text == " ," || maskedTextBox3.Text.Length != 3) { provider.SetError(maskedTextBox3, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }
            if (maskedTextBox4.Enabled == true) { if (maskedTextBox4.Text == " ," || maskedTextBox4.Text.Length != 3) { provider.SetError(maskedTextBox4, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }
            if (maskedTextBox5.Enabled == true) { if (maskedTextBox5.Text == " ," || maskedTextBox5.Text.Length != 3) { provider.SetError(maskedTextBox5, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }

            if (maskedTextBox6.Enabled == true) { if (maskedTextBox6.Text == " ," || maskedTextBox6.Text.Length != 3) { provider.SetError(maskedTextBox6, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }
            if (maskedTextBox7.Enabled == true) { if (maskedTextBox7.Text == " ," || maskedTextBox7.Text.Length != 3) { provider.SetError(maskedTextBox7, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }
            if (maskedTextBox8.Enabled == true) { if (maskedTextBox8.Text == " ," || maskedTextBox8.Text.Length != 3) { provider.SetError(maskedTextBox8, "Bu Alan Eksik veya Boş Geçilemez.."); sayac++; } }
            int tercih_sayac = 0;
            if (textBox12.Text != "") { tercih_sayac++; }
            if (textBox11.Text != "") { tercih_sayac++; }
            if (textBox10.Text != "") { tercih_sayac++; }
            if ((textBox12.Text != "" && (Convert.ToInt32(textBox12.Text) > tercih_sayac)) || (textBox11.Text != "" && (Convert.ToInt32(textBox11.Text) > tercih_sayac)) || (textBox10.Text != "" && (Convert.ToInt32(textBox10.Text) > tercih_sayac))) { MessageBox.Show("Tercih Değerlerini Kontrol Ediniz.."); sayac++; tercih_sayac = 0; }
            //if (comboBox5.SelectedIndex == 0 && textBox14.Text == "") { provider.SetError(textBox14, "Yerleştiği Tercihi Giriniz.."); sayac++; }
            if (textBox14.Text != "" && Convert.ToInt32(textBox14.Text) > tercih_sayac) { MessageBox.Show("Yerleştiği Tercih Değerini Kontrol Ediniz.."); sayac++; tercih_sayac = 0; }

        }
        void textsıfırlama()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1; // aday ekleme, silme, güncelleme işlemlerinden sonra veri giriş ekranını sıfırlar.
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            maskedTextBox1.Text = "   ,";
            maskedTextBox2.Text = "   ,";
            maskedTextBox3.Text = " ,";
            maskedTextBox4.Text = " ,";
            maskedTextBox5.Text = " ,";
            maskedTextBox6.Text = " ,";
            maskedTextBox7.Text = " ,";
            maskedTextBox8.Text = " ,";

        }
        void Exceldosyasıvarmı()
        {  // Verilerin tutulduğu excel dosyası vermı dite kontrol eder yoksa ExcelDosyasıOluştur methosu ile yeni Excel dosyası Oluşturur.
            string dosyayolu = "C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls";
            if (File.Exists(dosyayolu) == true)
            {
                excelkontrol.Text = "C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls";
                excelkontrol.ForeColor = Color.Green;
            }
            else
            {
                ExcelDosyasıOlustur();
                excelkontrol.Text = "C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls";
                excelkontrol.ForeColor = Color.Green;
            }
        }
        void ExcelDosyasıOlustur()
        { // Adayların bilgilerini tutulduğu excel dosyasını oluşturur.
            Directory.CreateDirectory("C:\\Özel Yetenek");
            string dosyayollu = "C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls";
            byte[] excel = Properties.Resources.OZEL_YETENEK_VERİ_GİRİŞİ;
            System.IO.FileStream fs = new System.IO.FileStream(dosyayollu, FileMode.CreateNew, FileAccess.ReadWrite);
            foreach (byte b in excel)
            {
                fs.WriteByte(b);
            }
            fs.Close();
        }
        void txtkontrol()
        { // branşarların, spor öz geçmiş belge türlerinin, kayıt sayac numarasının, kayıt sayacının kaçar kaçar artacağının tutulduğu metin belgelerini kontol eder. yoksa oluşturur.
            if (Directory.Exists("C:\\Özel Yetenek\\BRANŞLAR") == false)
            {
                Directory.CreateDirectory("C:\\Özel Yetenek\\BRANŞLAR");
            }
            if (Directory.Exists("C:\\Özel Yetenek\\BRANŞLAR") == true)
            {
                if (File.Exists("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt") == false)
                {
                    FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt", FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine("ARTİSTİK CİMNASTİK"); sw.WriteLine("ATLETİZM"); sw.WriteLine("BADMİNTON"); sw.WriteLine("BASKETBOL"); sw.WriteLine("CİMNASTİK"); sw.WriteLine("FUTBOL"); sw.WriteLine("TENİS"); sw.WriteLine("VOLEYBOL");
                    sw.WriteLine("YÜZME"); sw.WriteLine("YOK");
                    sw.Flush();
                    sw.Close();
                    fs.Close();
                }
                if (File.Exists("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt") == false)
                {
                    FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt", FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine("ATLETİZM"); sw.WriteLine("ATLETİZM/HAKEM"); sw.WriteLine("ATLETİZM/YRD.ANT."); sw.WriteLine("BASKETBOL"); sw.WriteLine("BASKETBOL/ANT.");
                    sw.WriteLine("BİLEK GÜREŞİ"); sw.WriteLine("BİSİKLET"); sw.WriteLine("FUTBOL"); sw.WriteLine("FUTSAL"); sw.WriteLine("GÜREŞ"); sw.WriteLine("HALK OYUNLARI");
                    sw.WriteLine("HENTBOL"); sw.WriteLine("HİS/WELLNESS"); sw.WriteLine("KARATE"); sw.WriteLine("KİCK BOKS"); sw.WriteLine("MASA TENİSİ"); sw.WriteLine("MOTORSİKLET");
                    sw.WriteLine("TAEKWONDO"); sw.WriteLine("TENİS"); sw.WriteLine("VOLEYBOL"); sw.WriteLine("YÜZME"); sw.WriteLine("YOK");
                    sw.Flush();
                    sw.Close();
                    fs.Close();
                }
                if (File.Exists("C:\\Özel Yetenek\\BRANŞLAR\\Kayit_Sayac.txt") == false)
                {
                    FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\Kayit_Sayac.txt", FileMode.OpenOrCreate, FileAccess.Write);
                    fs.Close();
                }
                if (File.Exists("C:\\Özel Yetenek\\BRANŞLAR\\Kayit_Sayac_Artis.txt") == false)
                {
                    FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\Kayit_Sayac_Artis.txt", FileMode.OpenOrCreate, FileAccess.Write);
                    fs.Close();
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void comboBox4_MouseClick(object sender, MouseEventArgs e)
        {     // comboboxa tıklandığı anda metin belgesinden okuma işlemini gerçekleştirir ve bütün verileri comboboxa ekler.
            comboBox4.Items.Clear();
            var bölümler = File.ReadLines(@"C:\Özel Yetenek\BRANŞLAR\SÖ BELGE TÜRÜ.txt");
            foreach (var bölüm in bölümler)
            {
                comboBox4.Items.Add(bölüm);
            }
        }

        private void comboBox7_MouseClick(object sender, MouseEventArgs e)
        {     // comboboxa tıklandığı anda metin belgesinden okuma işlemini gerçekleştirir ve bütün verileri comboboxa ekler.
            comboBox7.Items.Clear();
            var bölümler = File.ReadLines(@"C:\Özel Yetenek\BRANŞLAR\ANTRENÖRLÜK BRANŞ.txt");
            foreach (var bölüm in bölümler)
            {
                comboBox7.Items.Add(bölüm);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {  // OleDb kütüphanesini kullanarak excele headers(Başlık)'lara göre veri ekleme işlemini yapar.
            ogrencibilgikontrol();
            if (sayac == 0)
            {
                string a;
                if (comboBox8.SelectedIndex == 1) { a = comboBox8.Text + "(" + textBox13.Text + ")"; }
                else a = comboBox8.Text;

                OleDbCommand komut = new OleDbCommand();
                baglanti.Open();
                komut.Connection = baglanti;
                string sql = "Insert into [KAYIT-GİRİŞ$] ([Kayıt No],[TC Kimlik No],[Adı ],Soyadı,Cinsiyeti,[Doğum Tarihi],[Alan/Kol],[TYT P],AOÖBP,[SÖ Durumu],[SÖ Belge Türü],[Millilik ],Olimpiklik,[Spor Özgeçmiş Katsayısı BE],[Spor Özgeçmiş Katsayısı ANT],[Spor Özgeçmiş Katsayısı REK],[Millilikten Yerleşemediği Durumda SÖK BE],[Millilikten Yerleşemediği Durumda SÖK ANT],[Millilikten Yerleşemediği Durumda SÖK REK],[Daha Önce Yerleşme Durumu],[Başvurduğu Programlar BE],[Başvurduğu Programlar ANT],[Başvurduğu Programlar REK],[Millilikten Yerleştiği Tercihi],[ANTRENÖRLÜK BRANŞ],[MEKİK KOŞUSU]) values('" + textBox5.Text + "','" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + comboBox1.Text + "','" + textBox4.Text + "','" + comboBox2.Text + "','" + maskedTextBox1.Text + "','" + maskedTextBox2.Text + "','" + comboBox3.Text + "','" + comboBox4.Text + "','" + comboBox5.Text + "','" + comboBox9.Text + "','" + maskedTextBox3.Text + "','" + maskedTextBox4.Text + "','" + maskedTextBox5.Text + "','" + maskedTextBox8.Text + "','" + maskedTextBox7.Text + "','" + maskedTextBox6.Text + "','" + comboBox6.Text + "','" + textBox12.Text + "','" + textBox11.Text + "','" + textBox10.Text + "','"+textBox14.Text+"','" + comboBox7.Text + "','" + a + "')";
                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                baglanti.Close();
                kayıt_sayac();
                textsıfırlama();
                datagrid_veri_cekme();
                MessageBox.Show("Aday Başarıyla Eklendi");
                millisayac = 0;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {  // sadece harfleri yazdırır.
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
            && !char.IsSeparator(e.KeyChar);
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar >= 49 && (int)e.KeyChar <= 51)
            {
                e.Handled = false;//eğer 1 ve 3 arasında ise  yazdır.
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

        void kayıt_sayac()
        {   // kayıt butonuna tıklandığı anda ilk bu method çağırılı. kayıt sayac metin belgesinde veri varmı diye kontrol eder yoksa textboxtaki değeri alır ve üzerine kaçar artacaksa onu ekler öyle kaydeder. varsa o değeri alır kaçar artacaksa üzerine ekler ve günceller.
            {
                if (kayıtkontrolsayacı == 0)
                {
                    StringBuilder newFile = new StringBuilder();
                    string temp = "";
                    string[] file = File.ReadAllLines(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt");
                    int a = Convert.ToInt32(textBox5.Text);
                    foreach (string line in file)
                    {
                        if (line.Contains(textBox5.Text))
                        {
                            a = a + sayac_kacar_artacak;
                            temp = line.Replace(
                    textBox5.Text, a.ToString());
                            newFile.Append(temp +
                    "\r\n");
                            continue;
                        }
                        newFile.Append(line +
                    "\r\n");
                    }
                    File.WriteAllText(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt", newFile.ToString());
                    textBox5.Text = a.ToString();
                }
                if (kayıtkontrolsayacı > 0)
                {
                    FileStream fss = new FileStream(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac_Artis.txt", FileMode.Open, FileAccess.Write);
                    StreamWriter yazz = new StreamWriter(fss);
                    int b = 0;
                    if (radioButton1.Checked == true) { b = 1; sayac_kacar_artacak = 1; }
                    if (radioButton2.Checked == true) { b = 2; sayac_kacar_artacak = 2; }
                    yazz.WriteLine(b.ToString());
                    yazz.Flush();
                    yazz.Close();
                    fss.Close();
                    kayıtkontrolsayacı = 0;

                    FileStream fs = new FileStream(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt", FileMode.Open, FileAccess.Write);
                    StreamWriter yaz = new StreamWriter(fs);
                    int a = Convert.ToInt32(textBox5.Text);
                    a=a+sayac_kacar_artacak;
                    yaz.WriteLine(a.ToString());
                    yaz.Flush();
                    yaz.Close();
                    fs.Close();
                    kayıtkontrolsayacı = 0;

                }
                textboxsayackkontrol();
            }
        }

        void textboxsayackkontrol()
        { // başlangıçta  kayıt sayacı ve kayıt sayac artışı metin belgelerini  kontorol eder boşsalar radiobuton ve textboxları etkinleştirirler. doluysalar o kısıma müdahale edilmeyi kapatırlar. 
            StreamReader okuu = File.OpenText(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac_Artis.txt");
            string yazii;
            yazii = okuu.ReadLine();
            if (yazii == null) { radioButton1.Enabled = true; radioButton2.Enabled = true; kayıtkontrolsayacı++; }
            else if (yazii != null) { radioButton1.Enabled = false; radioButton2.Enabled = false; if (yazii == "1") { radioButton1.Checked = true; sayac_kacar_artacak = 1; } if (yazii == "2") { radioButton2.Checked = true; sayac_kacar_artacak = 2; } }
            okuu.Close();

            StreamReader oku = File.OpenText(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt");
            string yazi;
            yazi = oku.ReadLine();
            if (yazi == null) { textBox5.Enabled = true; kayıtkontrolsayacı++; }
            else if (yazi != null) { textBox5.Enabled = false; textBox5.Text = yazi.ToString(); }
            oku.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // bu kısımda data gridden seçilen kişinin kayıt numarası hafızaya alınır. kayıt numarası excel dosyasında 1. kolonda aranır bulunduğu satır komple silinir. ve silme işlemi gerçekleştirilir.
            DialogResult secenek = MessageBox.Show("Adayı Silmek İstediğinize Eminmisiniz?", "Onay Penceresi", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (secenek == DialogResult.Yes)
            {
                const string xlsPath = "C:\\Özel Yetenek\\OZEL_YETENEK_VERİ_GİRİŞİ.xls";
                Excel._Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(xlsPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Worksheet sheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("KAYIT-GİRİŞ");
                Microsoft.Office.Interop.Excel.Range Rng = sheet.get_Range("A1", Type.Missing);
                Microsoft.Office.Interop.Excel.Range findRange = Rng.Find(deger);
                int a = Convert.ToInt32(findRange.Row.ToString());
                Excel.Range ran = (Excel.Range)sheet.Rows[a, Type.Missing];
                ran.Select();
                ran.Delete(Excel.XlDirection.xlUp);
                string tmpName = System.IO.Path.GetTempFileName();
                System.IO.File.Delete(tmpName);
                excelWorkbook.SaveAs(tmpName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelWorkbook.Close(false, Type.Missing, Type.Missing);
                excelApp.Quit();
                System.IO.File.Delete(xlsPath);
                System.IO.File.Move(tmpName, xlsPath);
                datagrid_veri_cekme();
                button3.Enabled = false;
                button2.Enabled = true;
                button4.Enabled = false;
                button7.Enabled = false;
                textsıfırlama();
                textboxsayackkontrol();
                millisayac = 0;
                MessageBox.Show("Kayıt Başarı İle Silindi");
            }
            else if (secenek == DialogResult.No) { deger = ""; button3.Enabled = false; button4.Enabled = false; button2.Enabled = true; textsıfırlama(); }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        { // mekik koşusu başarısızsa % kaç başarısız olduğunu girebilmek için textboxın görünmezliği kaldırılıp veri girişie açılır.
            if (comboBox8.SelectedIndex == 1) { textBox13.Visible = true; }
            else textBox13.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();
            this.Hide();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        { // girilen harfleri büyük harfe çevirir.
            textBox2.Text = textBox2.Text.ToUpper();
            textBox2.SelectionStart = textBox2.Text.Length;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = textBox3.Text.ToUpper();
            textBox3.SelectionStart = textBox3.Text.Length;
        }

        void datagrid_veri_cekme()
        { // belirli excel konumdan data table a verileri aktarır ordanda data gride doldurur.
            baglanti.Open();
            tablo.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$]", baglanti);
            da.Fill(tablo);
            dataGridView1.DataSource = tablo;
            baglanti.Close();
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void textBox6_Enter(object sender, EventArgs e)
        {
            textBox6.Text = "";
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            if (textBox6.Text == "")
            {
                textBox6.Text = "T.C. Kimlik Numarasını Giriniz..";
            }
        }

        private void textBox7_Enter(object sender, EventArgs e)
        {
            textBox7.Text = "";
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            if (textBox7.Text == "")
            {
                textBox7.Text = "Adını Giriniz..";
            }
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            textBox8.Text = "";
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            if (textBox8.Text == "")
            {
                textBox8.Text = "Soyadını Giriniz..";
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {  // Sql sorguları kullanılarak excel dosyasında T.C. kimliğe göre arama yapılmasını sağlar.
            if (textBox6.Text != "T.C. Kimlik Numarasını Giriniz..")
            {
                tablo.Clear();
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] where [TC Kimlik No] LIKE'%" + textBox6.Text + "%'", baglanti);
                da.Fill(tablo);
                dataGridView1.DataSource = tablo;
                baglanti.Close();
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {   // Sql sorguları kullanılarak excel dosyasında ada göre arama yapılmasını sağlar.
            if (textBox7.Text != "Adını Giriniz..")
            {
                tablo.Clear();
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] where [Adı ] LIKE'%" + textBox7.Text + "%'", baglanti);
                da.Fill(tablo);
                dataGridView1.DataSource = tablo;
                baglanti.Close();
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        { // Sql sorguları kullanılarak excel dosyasında soyada göre arama yapılmasını sağlar.
            if (textBox8.Text != "Soyadını Giriniz..")
            {
                tablo.Clear();
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] where [Soyadı] LIKE'%" + textBox8.Text + "%'", baglanti);
                da.Fill(tablo);
                dataGridView1.DataSource = tablo;
                baglanti.Close();
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {  // datagride 2 kere tıklandığında seçilen adayın tük verileri ilgili kısımlara doldurulur.
            millisayac++;
            textBox5.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                string b = comboBox1.Items[i].ToString();
                if (a == b) { comboBox1.SelectedIndex = i; }
            }
            for (int i = 0; i < comboBox2.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                string b = comboBox2.Items[i].ToString();
                if (a == b) { comboBox2.SelectedIndex = i; }
            }
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            for (int i = 0; i < comboBox3.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                string b = comboBox3.Items[i].ToString();
                if (a == b) { comboBox3.SelectedIndex = i; }
            }
            for (int i = 0; i < comboBox4.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                string b = comboBox4.Items[i].ToString();
                if (a == b) { comboBox4.SelectedIndex = i; }
            }
            for (int i = 0; i < comboBox5.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                string b = comboBox5.Items[i].ToString();
                if (a == b) { comboBox5.SelectedIndex = i; }
            }
            for (int i = 0; i < comboBox9.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                string b = comboBox9.Items[i].ToString();
                if (a == b) { comboBox9.SelectedIndex = i; break; }
                else comboBox9.SelectedIndex = -1; 
            }
            maskedTextBox3.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            maskedTextBox4.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            maskedTextBox5.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            maskedTextBox8.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            maskedTextBox7.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            maskedTextBox6.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();



            for (int i = 0; i < comboBox6.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                string b = comboBox6.Items[i].ToString();
                if (a == b) { comboBox6.SelectedIndex = i; }
            }
            textBox12.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            textBox11.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            textBox10.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            textBox14.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            for (int i = 0; i < comboBox7.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[24].Value.ToString();
                string b = comboBox7.Items[i].ToString();
                if (a == b) { comboBox7.SelectedIndex = i; }
            }
            int x = 0;
            for (int i = 0; i < comboBox8.Items.Count; i++)
            {
                string a = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                string b = comboBox8.Items[i].ToString();
                if (a == b) { comboBox8.SelectedIndex = i; }
                else { x++; }
            }
            if (x == 3)
            {
                string a = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                string b = comboBox8.Items[1].ToString();
                comboBox8.SelectedIndex = 1;
                string[] parcalar;
                string k = "()";
                char[] ayrac = k.ToCharArray();
                parcalar = a.Split(ayrac);
                textBox13.Text = parcalar[1];
            }

            deger = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            button3.Enabled = true;
            button7.Enabled = true;
            button2.Enabled = false;
            button4.Enabled = true;
            button5.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {  
            // data gridden seçilen adayın verileri güncellendikten sonra update sorgusu ile aday bilgileri güncellenir.
            
            ogrencibilgikontrol();
            if (sayac == 0)
            {
                string a;
                if (comboBox8.SelectedIndex == 1) { a = comboBox8.Text + "(" + textBox13.Text + ")"; }
                else a = comboBox8.Text;

                OleDbCommand komut = new OleDbCommand();
                baglanti.Open();
                komut.Connection = baglanti;
                string sql = "Update [KAYIT-GİRİŞ$] set [TC Kimlik No]='" + textBox1.Text + "',[Adı ]='" + textBox2.Text + "',Soyadı='" + textBox3.Text + "',Cinsiyeti='" + comboBox1.Text + "',[Doğum Tarihi]='" + textBox4.Text + "',[Alan/Kol]='" + comboBox2.Text + "',[TYT P]='" + maskedTextBox1.Text + "',AOÖBP='" + maskedTextBox2.Text + "',[SÖ Durumu]='" + comboBox3.Text + "',[SÖ Belge Türü]='" + comboBox4.Text + "',[Millilik ]='" + comboBox5.Text + "',Olimpiklik='" + comboBox9.Text + "',[Spor Özgeçmiş Katsayısı BE]='" + maskedTextBox3.Text + "',[Spor Özgeçmiş Katsayısı ANT]='" + maskedTextBox4.Text + "',[Spor Özgeçmiş Katsayısı REK]='" + maskedTextBox5.Text + "',[Millilikten Yerleşemediği Durumda SÖK BE]='" + maskedTextBox8.Text + "',[Millilikten Yerleşemediği Durumda SÖK ANT]='" + maskedTextBox7.Text + "',[Millilikten Yerleşemediği Durumda SÖK REK]='" + maskedTextBox6.Text + "',[Daha Önce Yerleşme Durumu]='" + comboBox6.Text + "',[Başvurduğu Programlar BE]='" + textBox12.Text + "',[Başvurduğu Programlar ANT]='" + textBox11.Text + "',[Başvurduğu Programlar REK]='" + textBox10.Text + "',[Millilikten Yerleştiği Tercihi]='"+textBox14.Text+"',[ANTRENÖRLÜK BRANŞ]='" + comboBox7.Text + "',[MEKİK KOŞUSU]='" + a + "' WHERE [Kayıt No]='" + textBox5.Text + "'";
                komut.CommandText = sql;
                komut.ExecuteNonQuery();
                baglanti.Close();
                textboxsayackkontrol();
                textsıfırlama();
                datagrid_veri_cekme();
                button4.Enabled = false;
                button3.Enabled = false;
                button2.Enabled = true;
                button7.Enabled = false;
                millisayac = 0;
                MessageBox.Show("Aday Başarıyla Güncellendi");
            }
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            textBox9.Text = "";
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (textBox9.Text == "")
            {
                textBox9.Text = "Kayıt No Giriniz..";
            }

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {  // Sql sorguları kullanılarak excel dosyasında kayıt numarasına göre arama yapılmasını sağlar.
            if (textBox9.Text != "Kayıt No Giriniz..")
            {
                tablo.Clear();
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [KAYIT-GİRİŞ$] where [Kayıt No] LIKE'%" + textBox9.Text + "%'", baglanti);
                da.Fill(tablo);
                dataGridView1.DataSource = tablo;
                baglanti.Close();
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        { // Tercih kısımında BE textbox'ına veri girildiği anda aday milli ise millilik kat sayı ve yerleşti tercihi kısımları aktifleştirilir.
            if (textBox12.Text != "") { maskedTextBox3.Enabled = true; maskedTextBox3.Focus();  }

            else { maskedTextBox3.Enabled = false; maskedTextBox3.Clear();  }

            if (textBox12.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox8.Enabled = true; label26.Visible = true; textBox14.Visible = true; }

            else { maskedTextBox8.Text = " ,"; maskedTextBox8.Enabled = false; label26.Visible = false; textBox14.Visible = false; }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        { // Tercih kısımında ANT textbox'ına veri girildiği anda aday milli ise millilik kat sayı ve yerleşti tercihi kısımları aktifleştirilir.
            if (textBox11.Text != "") { maskedTextBox4.Enabled = true; maskedTextBox4.Focus(); }

            else { maskedTextBox4.Enabled = false; maskedTextBox4.Clear(); }

            if (textBox11.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox7.Enabled = true; label26.Visible = true; textBox14.Visible = true; }

            else { maskedTextBox7.Text = " ,"; maskedTextBox7.Enabled = false; label26.Visible = false; textBox14.Visible = false; }


        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {  // Tercih kısımında REK textbox'ına veri girildiği anda aday milli ise millilik kat sayı ve yerleşti tercihi kısımları aktifleştirilir.
            if (textBox10.Text != "") { maskedTextBox5.Enabled = true; maskedTextBox5.Focus(); }

            else { maskedTextBox5.Enabled = false; maskedTextBox5.Clear(); }

            if (textBox10.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox6.Enabled = true; label26.Visible = true; textBox14.Visible = true; }

            else { maskedTextBox6.Text = " ,"; maskedTextBox6.Enabled = false; label26.Visible = false; textBox14.Visible = false; }
        }

        private void button5_Click(object sender, EventArgs e)
        { // İptal butonuna tıklandığında yapılan işlemler.
            millisayac = 0;
            textboxsayackkontrol();
            deger = "";
            button3.Enabled = false;
            button4.Enabled = false;
            button7.Enabled = false;
            button2.Enabled = true;
            textsıfırlama();
            button5.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            this.Hide();
            f3.Show();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        { // aday milli olarak seçildikten sonra diyelimki verileri girdik ama aday milli değilmiş bu kısımda aday milli olmayan olarak seçildiği anda millilikle ilgili alanlar kapatılıp sıfırlanır.

            if (textBox12.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox8.Enabled = true; maskedTextBox8.Text = maskedTextBox3.Text;  }

            else { maskedTextBox8.Text = " ,"; maskedTextBox8.Enabled = false; }

            if (textBox11.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox7.Enabled = true; maskedTextBox7.Text = maskedTextBox4.Text;}

            else { maskedTextBox7.Text = " ,"; maskedTextBox7.Enabled = false; }

            if (textBox10.Text != "" && comboBox5.SelectedIndex == 0) { maskedTextBox6.Enabled = true; maskedTextBox6.Text = maskedTextBox5.Text; }

            else { maskedTextBox6.Text = " ,"; maskedTextBox6.Enabled = false; }

            if (comboBox5.SelectedIndex == 0 && (textBox12.Text != "" || textBox11.Text != "" || textBox10.Text != "")) { label26.Visible = true; textBox14.Visible = true; }
            else { label26.Visible = false; textBox14.Text = ""; textBox14.Visible = false; }

            if (comboBox5.SelectedIndex == 0) { comboBox9.Enabled = true; }
            else if (comboBox5.SelectedIndex == 1) { comboBox9.Enabled = false; comboBox9.SelectedIndex = -1; }
        }

        private void maskedTextBox3_TextChanged(object sender, EventArgs e)
        {
            if (millisayac == 0)
            {
                if (maskedTextBox8.Enabled == true) { maskedTextBox8.Text = maskedTextBox3.Text; }
            }


        }

        private void maskedTextBox4_TextChanged(object sender, EventArgs e)
        {

            if (millisayac == 0)
            {
                if (maskedTextBox7.Enabled == true) { maskedTextBox7.Text = maskedTextBox4.Text; }
            }

        }

        private void maskedTextBox5_TextChanged(object sender, EventArgs e)
        {

            if (millisayac == 0)
            {
                if (maskedTextBox6.Enabled == true) { maskedTextBox6.Text = maskedTextBox5.Text; }
            }

        }

        Font Baslik = new Font("Times New Roman",14,FontStyle.Bold);  // yazdırma işlemi yazı tipi fontu ve kalınlığı.
        Font AltBaslik = new Font("Times New Roman",12,FontStyle.Bold);
        Font yazı = new Font("Times New Roman",12);
        SolidBrush sb = new SolidBrush(Color.Black); // yazdırma işlemi fırça rengi.
        DateTime Tarih = DateTime.Today;         // yazdırma işlemi için bu günün tarihi.
        int yil = Convert.ToInt32(DateTime.Now.Year);   // yazdırma işlemi başlık tarihi için şimdiki yıl.
        int sonrakiyil = Convert.ToInt32(DateTime.Now.Year) + 1;   // yazdırma işlemi başlık tarihi için bir sonraki yıl.

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {  
            // burada a4 kağıdına belirlediğimiz şablon ver kordinatta veriler otomatik girilir ve yazdırma işlemi gerçekleştirilir.
           
            StringFormat st = new StringFormat();
            st.Alignment = StringAlignment.Near;
            e.Graphics.DrawString("PAMUKKALE ÜNİVERSİTESİ SPOR BİLİMLERİ FAKÜLTESİ \n         "+yil.ToString()+"-"+sonrakiyil.ToString()+" ÖZEL YETENEK SINAVI KAYIT FORMU",Baslik,sb,75,20,st);

            e.Graphics.DrawString("-----------------------------------------------------------------------------------------------------", AltBaslik, sb, 75, 100, st);

            e.Graphics.DrawString("Kayıt No:", AltBaslik, sb, 125, 120, st);
            e.Graphics.DrawString(textBox5.Text, yazı, sb, 203, 120, st);

            e.Graphics.DrawString("T.C Kimlik No:", AltBaslik, sb, 125, 155, st);
            e.Graphics.DrawString(textBox1.Text, yazı, sb, 243, 155, st);

            e.Graphics.DrawString("Adı:", AltBaslik, sb, 125, 190, st);
            e.Graphics.DrawString(textBox2.Text, yazı, sb, 160, 190, st);

            e.Graphics.DrawString("Soyadı:", AltBaslik, sb, 125, 225, st);
            e.Graphics.DrawString(textBox3.Text, yazı, sb, 185, 225, st);

            e.Graphics.DrawString("Cinsiyeti:", AltBaslik, sb, 125, 260, st);
            e.Graphics.DrawString(comboBox1.Text, yazı, sb, 201, 260, st);

            e.Graphics.DrawString("Doğum Tarihi:", AltBaslik, sb, 125, 295, st);
            e.Graphics.DrawString(textBox4.Text, yazı, sb, 234, 295, st);

            e.Graphics.DrawString("Alan/Kol:", AltBaslik, sb, 125, 330, st);
            e.Graphics.DrawString(comboBox2.Text, yazı, sb, 203, 330, st);

            e.Graphics.DrawString("TYT P:", AltBaslik, sb, 125, 365, st);
            e.Graphics.DrawString(maskedTextBox1.Text, yazı, sb, 183, 365, st);

            e.Graphics.DrawString("AOÖBP:", AltBaslik, sb, 125, 400, st);
            e.Graphics.DrawString(maskedTextBox2.Text, yazı, sb, 194, 400, st);

            e.Graphics.DrawString("SÖ Durumu:", AltBaslik, sb, 125, 435, st);
            e.Graphics.DrawString(comboBox3.Text, yazı, sb, 225, 435, st);

            e.Graphics.DrawString("SÖ Belge Türü:", AltBaslik, sb, 125, 470, st);
            e.Graphics.DrawString(comboBox4.Text, yazı, sb, 243, 470, st);

            e.Graphics.DrawString("Millilik:", AltBaslik, sb, 125, 505, st);
            e.Graphics.DrawString(comboBox5.Text, yazı, sb, 194, 505, st);

            e.Graphics.DrawString("Spor Özgeçmiş Kat Sayısı BE:", AltBaslik, sb, 125, 540, st);
            e.Graphics.DrawString(maskedTextBox3.Text, yazı, sb, 346, 540, st);

            e.Graphics.DrawString("Spor Özgeçmiş Kat Sayısı ANT:", AltBaslik, sb, 125, 575, st);
            e.Graphics.DrawString(maskedTextBox4.Text, yazı, sb, 356, 575, st);

            e.Graphics.DrawString("Spor Özgeçmiş Kat Sayısı REK:", AltBaslik, sb, 125, 610, st);
            e.Graphics.DrawString(maskedTextBox5.Text, yazı, sb, 359, 610, st);

            e.Graphics.DrawString("Bir Önceki Sene Yerleşme Durumu:", AltBaslik, sb, 125, 645, st);
            e.Graphics.DrawString(comboBox6.Text, yazı, sb, 384, 645, st);

            e.Graphics.DrawString("Başvurduğu Programalar BE:", AltBaslik, sb, 125, 680, st);
            e.Graphics.DrawString(textBox12.Text, yazı, sb, 346, 680, st);

            e.Graphics.DrawString("Başvurduğu Programalar ANT:", AltBaslik, sb, 125, 715, st);
            e.Graphics.DrawString(textBox11.Text, yazı, sb, 356, 715, st);

            e.Graphics.DrawString("Başvurduğu Programalar REK:", AltBaslik, sb, 125, 750, st);
            e.Graphics.DrawString(textBox10.Text, yazı, sb, 359, 750, st);

            e.Graphics.DrawString("Antrenörlük Branş:", AltBaslik, sb, 125, 785, st);
            e.Graphics.DrawString(comboBox7.Text, yazı, sb, 272, 785, st);


            e.Graphics.DrawString("-----------------------------------------------------------------------------------------------------", AltBaslik, sb, 75, 805, st);

            e.Graphics.DrawString("Yukarıda belirtilen bilgilerin doğruluğunu kabul ediyorum.", yazı, sb, 75, 845, st);
            e.Graphics.DrawString("Not: Bu formda belirtilen puanlar ön değerlendirme puanları olup, komisyon \nincelemesi sonucunda kesinleşecektir.", AltBaslik, sb, 75, 870, st);
            
            e.Graphics.DrawString("ADAYIN İMZASI", yazı, sb, 75, 930, st);

            e.Graphics.DrawString("TARİH", yazı, sb, 315, 930, st);
            e.Graphics.DrawString(Tarih.ToShortDateString(), yazı, sb, 300, 955, st);

            e.Graphics.DrawString("KONTROL EDEN", yazı, sb, 505, 935, st);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.ShowDialog();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Hakkımızda h = new Hakkımızda();
            h.Show();
        }
    }
}
