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

namespace Lisans_Tezi11
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        void listview_veri_cekme()
        {   // SÖ BELGE TÜRÜ ve ANTRENÖRLÜK BRANŞ isimli metin belgelerinden  veriler listboxlara eklenir.
            listBox1.Items.Clear();
            var bölümler = File.ReadLines("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt");
            foreach (var bölüm in bölümler)
            {
                if (bölüm != "")
                {
                    listBox1.Items.Add(bölüm);
                }
            }

            listBox2.Items.Clear();
            var bölümlerr = File.ReadLines("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt");
            foreach (var bölüm in bölümlerr)
            {
                if (bölüm != "")
                {
                    listBox2.Items.Add(bölüm);
                }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {  // başlangıç kısmında listview_veri_cekme methodu tetiklenir.
            listview_veri_cekme();
            button4.Visible = false;
            button5.Visible = false;
        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {  //textbox girilen harler büyük harfe dönüştürülür.
            textBox2.Text = textBox2.Text.ToUpper();
            textBox2.SelectionStart = textBox2.Text.Length;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {       //textbox girilen harler büyük harfe dönüştürülür.
            textBox1.Text = textBox1.Text.ToUpper();
            textBox1.SelectionStart = textBox1.Text.Length;
        }

        private void button1_Click(object sender, EventArgs e)
        {     
            // bu kısımda textboxtaki veri filestream ve streamwriter aracılığı ile metin belgesine eklenir.
           
            if (textBox1.Text != "")
            {
                listBox1.Items.Add(textBox1.Text);
                textBox1.Text = "";
                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt", FileMode.Open, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    sw.WriteLine(listBox1.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                MessageBox.Show("SÖ Belge Türleri Kaydedildi");
            }
            else MessageBox.Show("Bir SÖ Belge Türü Giriniz");
        }

        private void button8_Click(object sender, EventArgs e)
        {     // bu kısımda textboxtaki veri filestream ve streamwriter aracılığı ile metin belgesine eklenir.
            if (textBox2.Text != "")
            {
                listBox2.Items.Add(textBox2.Text);
                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt", FileMode.Open, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox2.Items.Count; i++)
                {
                    sw.WriteLine(listBox2.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                textBox2.Text = "";
                MessageBox.Show("Antrenörlük Branş Türü Eklendi..");

            }
            else MessageBox.Show("Bir Antrenörlük Branş Türü Giriniz");
        }

        private void button2_Click(object sender, EventArgs e)
        { 
            // bu kısımda listboxtan seçilen  veri filestream ve streamwriter aracılığı ile metin belgesinden silinir.
            
            if (listBox1.SelectedIndex != -1)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                File.Delete("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt");
                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    sw.WriteLine(listBox1.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                textBox1.Text = "";
                button1.Enabled = true;
                button4.Visible = false;
                MessageBox.Show("SÖ Belge Türü Silindi..");
            }
            else MessageBox.Show("Silme İşlemi İçin Seçiminizi Yapınız..");
        }

        private void button7_Click(object sender, EventArgs e)
        {    
            // bu kısımda listboxtan seçilen  veri filestream ve streamwriter aracılığı ile metin belgesinden silinir.
            
            if (listBox2.SelectedIndex != -1)
            {
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
                File.Delete("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt");
                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox2.Items.Count; i++)
                {
                    sw.WriteLine(listBox2.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                textBox2.Text = "";
                button8.Enabled = true;
                button5.Visible = false;
                MessageBox.Show("Antrenörlük Branş Türü Silindi..");
            }
            else MessageBox.Show("Silme İşlemi İçin Seçiminizi Yapınız..");
        }

        private void button3_Click(object sender, EventArgs e)
        {  
            // bu kısımda listboxtaki seçilen  veri filestream ve streamwriter aracılığı ile metin belgesinde güncellenir.
            
            if (listBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Güncelleme İstediğiniz Satırı Seçiniz.");
            }
            else if (textBox1.Text == "") { MessageBox.Show("Güncellemek İstediğiniz SÖ Belge Türünü Giriniz"); }
            else
            {
                int index = listBox1.SelectedIndex;
                listBox1.Items.Remove(listBox1.SelectedItem);
                listBox1.Items.Insert(index, textBox1.Text);
                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\SÖ BELGE TÜRÜ.txt", FileMode.Open, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    sw.WriteLine(listBox1.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                textBox1.Text = "";
                button1.Enabled = true;
                button4.Visible = false;
                MessageBox.Show("SÖ Belge Türü Güncellendi..");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            // bu kısımda listboxtaki seçilen  veri filestream ve streamwriter aracılığı ile metin belgesinde güncellenir.

            if (listBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Güncelleme İstediğiniz Satırı Seçiniz.");
            }
            else if (textBox2.Text == "") { MessageBox.Show("Güncellemek İstediğiniz Antrenörlük Branş Türünü Giriniz"); }
            else
            {
                int index = listBox2.SelectedIndex;
                listBox2.Items.Remove(listBox2.SelectedItem);
                listBox2.Items.Insert(index, textBox2.Text);

                FileStream fs = new FileStream("C:\\Özel Yetenek\\BRANŞLAR\\ANTRENÖRLÜK BRANŞ.txt", FileMode.Open, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < listBox2.Items.Count; i++)
                {
                    sw.WriteLine(listBox2.Items[i].ToString());
                }
                sw.Close();
                fs.Close();
                textBox2.Text = "";
                button8.Enabled = true;
                button5.Visible = false;
                MessageBox.Show("Antrenörlük Branş Türü Güncellendi..");
            }
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 fr = new Form1();
            fr.Show();
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            // listboxtan seçilen veri textboxa aktarılır.
            try
            {
                textBox1.Text = listBox1.SelectedItem.ToString();
                button1.Enabled = false;
                button4.Visible = true;
            }
            catch (Exception)
            {
            }

        }

        private void listBox2_MouseClick(object sender, MouseEventArgs e)
        {
            // listboxtan seçilen veri textboxa aktarılır.

            try
            {
                textBox2.Text = listBox2.SelectedItem.ToString();
                button8.Enabled = false;
                button5.Visible = true;
            }
            catch (Exception)
            {
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            // iptal butonu işlemleri

            textBox1.Text = "";
            button1.Enabled = true;
            button4.Visible = false;
            listBox1.SelectedIndex = -1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // iptal butonu işlemleri

            textBox2.Text = "";
            button8.Enabled = true;
            button5.Visible = false;
            listBox2.SelectedIndex = -1;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // programdaki tüm verileri silme işlemi burada geçekleştirilir. önce dosyalar varmı diye kontrol edilir varsa silinir.
            DialogResult secenek = MessageBox.Show("Verileri Sıfırlamak İstediğinize Eminmisiniz?", "Onay Penceresi", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (secenek == DialogResult.Yes)
            {
                if (File.Exists(@"C:\Özel Yetenek\OZEL_YETENEK_VERİ_GİRİŞİ.xls"))
                {
                    File.Delete(@"C:\Özel Yetenek\OZEL_YETENEK_VERİ_GİRİŞİ.xls");
                }
                if (File.Exists(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac_Artis.txt"))
                {
                    File.Delete(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac_Artis.txt");
                }
                if (File.Exists(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt"))
                {
                    File.Delete(@"C:\Özel Yetenek\BRANŞLAR\Kayit_Sayac.txt");
                }
                if (Directory.Exists(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları", true);
                }
                if (Directory.Exists(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları", true);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            // programdaki millilerin verileri silme işlemi burada geçekleştirilir. önce dosyalar varmı diye kontrol edilir varsa silinir.
            
            DialogResult secenek = MessageBox.Show("Millilerin Verilerin Sıfırlamak İstediğinize Eminmisiniz?", "Onay Penceresi", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (secenek == DialogResult.Yes)
            {
                if (Directory.Exists(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milliler", true);
                }
                if (Directory.Exists(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları\Milliler"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları\Milliler", true);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // programdaki milli olmayanların verileri silme işlemi burada geçekleştirilir. önce dosyalar varmı diye kontrol edilir varsa silinir.

            DialogResult secenek = MessageBox.Show("Milli Olmayanların Verilerini Sıfırlamak İstediğinize Eminmisiniz?", "Onay Penceresi", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (secenek == DialogResult.Yes)
            {
                if (Directory.Exists(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Düzenlenmiş Aday Tabloları\Milli Olmayanlar", true);
                }
                if (Directory.Exists(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları\Milli Olmayanlar"))
                {
                    Directory.Delete(@"C:\Özel Yetenek\Yerleştirilmiş Aday Tabloları\Milli Olmayanlar", true);
                }
            }
        }



    }
}
