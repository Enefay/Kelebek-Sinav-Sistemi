using KelebekSistemi.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;

namespace KelebekSistemi
{
    public partial class Default : System.Web.UI.Page
    {
        List<Ogrenci> ogrenci = new List<Ogrenci>();
        List<Sinif> sinif = new List<Sinif>();
        static List<Sinif> secilenSinif = new List<Sinif>();
        int kontenjan;
        static int ogrencikontenjani;

        static string dersadi = "";
        static string hocaadi = "";
        static string sinavtipi = "";
        static string sinavtarihi = "";
        static string sinavsaati = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            string csvyolu = @"C:\Users\batuh\source\repos\KelebekSistemi\KelebekSistemi\siniflar.csv";

            VeriCek(csvyolu);

        }

        private void VeriCek(string csvyolu)
        {
            string[] satirlar = System.IO.File.ReadAllLines(csvyolu);
            if (satirlar.Length > 0)
            {
                //Veriler için kodlarımız
                for (int i = 0; i < satirlar.Length; i++)
                {
                    string[] veriler = satirlar[i].Split(',');
                    Sinif snf = new Sinif();
                    snf.sinifAdi = veriler[0];
                    snf.sinifKontenjan = Convert.ToInt32(veriler[1]);
                    sinif.Add(snf);
                }
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (fluExcell.HasFile)
            {
                string fileExtension = System.IO.Path.GetExtension(fluExcell.FileName).ToLower();
                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string guid = System.Guid.NewGuid().ToString();
                    string newFileName = guid + fileExtension;
                    string filePath = System.IO.Path.GetFullPath(Server.MapPath("~/App_Data/"));
                    fluExcell.PostedFile.SaveAs(filePath + newFileName);
                    LoadExcell(newFileName);
                }
                else
                {
                    Lblmessage.Text = "Excell dosyası seçiniz !";
                }
            }
        }

        private void LoadExcell(string newFileName)
        {
            OleDbConnection oleDbConn = new OleDbConnection();
            string path = System.IO.Path.GetFullPath(Server.MapPath("~/App_Data/")) + newFileName;
            if (Path.GetExtension(path) == ".xls")
            {
                oleDbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";");
            }
            else if (Path.GetExtension(path) == ".xlsx")
            {
                oleDbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0 Xml; HDR = NO\"; ");
            }

            oleDbConn.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = oleDbConn;
            cmd.CommandText = "SELECT * FROM [Sayfa1$]";
            cmd.CommandType = CommandType.Text;
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

            DataTable dt = new DataTable();
            adapter.Fill(dt);
            oleDbConn.Close();
            string ogrencino, adsoyad;




            for (int i = 1; i < dt.Rows.Count; i++)
            {
                ogrencino = dt.Rows[i][1].ToString();
                adsoyad = dt.Rows[i][2].ToString();

                SqlConnection baglanti;
                SqlCommand komut, kmt;

                string baglanStr = ConfigurationManager.ConnectionStrings["kelebekbaglan"].ConnectionString;
                baglanti = new SqlConnection(baglanStr);
                kmt = new SqlCommand("SELECT * FROM OgrenciTablosu where OgrenciNumara = @no", baglanti);
                kmt.Parameters.AddWithValue("@no", ogrencino);

                baglanti.Open();
                SqlDataReader dr = kmt.ExecuteReader();
                if (dr.Read())
                {
                    dr.Close();
                }
                else
                {
                    komut = new SqlCommand("INSERT INTO OgrenciTablosu (OgrenciAdSoyad, OgrenciNumara) VALUES( @adi, @numarasi)", baglanti);
                    komut.Parameters.AddWithValue("@adi", adsoyad);
                    komut.Parameters.AddWithValue("@numarasi", ogrencino);
                    dr.Close();
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                }
            }



            ogrencikontenjani = dt.Rows.Count - 1;
            


            if (dt.Rows[0]["F1"].ToString() != "SN" || dt.Rows[0]["F2"].ToString() != "Öğrenci No" || dt.Rows[0]["F3"].ToString() != "Adı Soyadı" || dt.Rows[0]["F6"].ToString() != "Ders Adı" || dt.Rows[1]["F6"].ToString() != "Öğretim Elemanı Adı" || dt.Rows[2]["F6"].ToString() != "Sınav Tipi" || dt.Rows[3]["F6"].ToString() != "Sınav Tarihi" || dt.Rows[4]["F6"].ToString() != "Sınavın Saati" || dt.Rows[0]["F7"].ToString() == "" || dt.Rows[1]["F7"].ToString() == "" || dt.Rows[2]["F7"].ToString() == "" || dt.Rows[3]["F7"].ToString() == "" || dt.Rows[4]["F7"].ToString() == "")
            {
                Lblhata.Text = "Yüklenen Dosya Hatalıdır. Lütfen Dosyayı Tekrar Kontrol Edin !";

            }
            else
            {
                rptExcell.DataSource = dt;
                rptExcell.DataBind();
                foreach (var item in sinif)
                {
                    CheckBoxList.Items.Add(item.sinifAdi + " Kontenjan: " + item.sinifKontenjan);
                }
                ToplamOgrenciSayisi.Text = "Toplam Öğrenci Sayısı: " + (dt.Rows.Count - 1).ToString();
            }

            //SINAV BILGILERI
            dersadi = dt.Rows[0]["F7"].ToString();
            hocaadi = dt.Rows[1]["F7"].ToString();
            sinavtipi = dt.Rows[2]["F7"].ToString();
            sinavtarihi = dt.Rows[3]["F7"].ToString();
            sinavsaati = dt.Rows[4]["F7"].ToString();


            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        protected void CheckBoxList_SelectedIndexChanged(object sender, EventArgs e)
        {
            kontenjan = 0;
            secilenSinif.Clear();
            foreach (System.Web.UI.WebControls.ListItem item in CheckBoxList.Items)
            {
                if (item.Selected)
                {
                    kontenjan += Convert.ToInt32(item.Value.Substring(15, 2));
                    //secilensinif doldur...
                    Sinif snf = new Sinif();
                    snf.sinifAdi = item.Value.Substring(0, 3);
                    snf.sinifKontenjan = Convert.ToInt32(item.Value.Substring(15, 2));
                    secilenSinif.Add(snf); //secilen siniflar

                }
            }
            SecilenSinif.Text = "Seçilen Kontenjan Sayısı: " + kontenjan.ToString();



            if (ogrencikontenjani <= kontenjan)
            {
                KaydetDevamEt.Visible = true;
            }
            else
            {
                KaydetDevamEt.Visible = false;
            }
        }

        protected void KaydetDevamEt_Click(object sender, EventArgs e)
        {

            SqlCommand command;
            SqlConnection cnn;
            string cnntr = ConfigurationManager.ConnectionStrings["kelebekbaglan"].ConnectionString;
            cnn = new SqlConnection(cnntr);
            cnn.Open();
            SqlDataReader dataReader;
            String sql;
            sql = "Select OgrenciAdSoyad, OgrenciNumara from OgrenciTablosu";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                Ogrenci ogr = new Ogrenci();//ogrenci bilgilerinin tutuldugu model
                ogr.AdSoyad = dataReader.GetValue(0).ToString();
                ogr.Numara = dataReader.GetValue(1).ToString();
                ogrenci.Add(ogr);
            }

            dataReader.Close();
            command.Dispose();
            cnn.Close();




            //random farklı sayilar üretme
            ArrayList sayilar = new ArrayList();
            Random r = new Random();
            int i = 0;
            while (i < ogrenci.Count)
            {
                int sayi = r.Next(0, ogrenci.Count);
                if (sayilar.Contains(sayi))
                    continue;
                sayilar.Add(sayi);
                i++;
            }
            //üretilen random sayilarla ogrenci listesinden rastgele isim cagirarak siralama
            List<Ogrenci> RandomOgrenciler = new List<Ogrenci>();
            int rndm, sayac = 0;
            while (sayac < ogrenci.Count)
            {
                rndm = Convert.ToInt32(sayilar[sayac]);
                RandomOgrenciler.Add(ogrenci[rndm]); //ogrenci ad ve numara tutulan model
                sayac++;
            }
            //int syc = 0;
            int syckontenjan = 0;
            foreach (var item in secilenSinif)
            {
                for (int x = 0; x < item.sinifKontenjan; x++)
                {
                    if (syckontenjan >= RandomOgrenciler.Count)
                    {
                        break;
                    }
                    SqlDataReader dr;
                    SqlCommand kmt = new SqlCommand("SELECT * FROM YerlestirmeTablosu where OgrenciNumara = @no", cnn);
                    kmt.Parameters.AddWithValue("@no", RandomOgrenciler[syckontenjan].Numara);
                    //rastgele ogrencilerin siniflara aktarildiktan sonra vt'ye kaydedilmesi
                    cnn.Open();
                    dr = kmt.ExecuteReader();
                    SqlCommand cmdx;
                    if (dr.Read())
                    {
                        cmdx = new SqlCommand("update YerlestirmeTablosu set OgrenciNumara= '" + RandomOgrenciler[syckontenjan].Numara + "',OgrenciAdSoyad='" + RandomOgrenciler[syckontenjan].AdSoyad + "',OgrenciSinif= '" + item.sinifAdi + "' where OgrenciNumara='" + RandomOgrenciler[syckontenjan].Numara + "'", cnn);
                    }
                    else
                    {
                        cmdx = new SqlCommand("Insert into YerlestirmeTablosu (OgrenciNumara,OgrenciAdSoyad,OgrenciSinif) values('" + RandomOgrenciler[syckontenjan].Numara + "','" + RandomOgrenciler[syckontenjan].AdSoyad + "','" + item.sinifAdi + "')", cnn);
                    }
                    cnn.Close();
                    dr.Close();
                    cnn.Open();
                    cmdx.ExecuteNonQuery();
                    cnn.Close();
                    syckontenjan++;

                }
            }





            //sınıf listeleri
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;


            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));

            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            oSheet.Columns[1].ColumnWidth = 7;
            oSheet.Columns[2].ColumnWidth = 14;
            oSheet.Columns[3].ColumnWidth = 40;
            oSheet.Columns[4].ColumnWidth = 50;

            oSheet.Rows[1].RowHeight = 50;


            oSheet.Cells[2, 1] = "Sıra No";
            oSheet.Cells[2, 2] = "Öğrenci No";
            oSheet.Cells[2, 3] = "Ad Soyad";
            oSheet.Cells[2, 4] = "İmza";

            //BASLİK
            oSheet.get_Range("A1", "D1").Merge();
            oSheet.get_Range("A1", "D1").Font.Bold = true;
            oSheet.get_Range("A1", "D1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oSheet.get_Range("A1", "D1").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

            oSheet.get_Range("A2", "D2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oSheet.get_Range("A2", "D2").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

            //ALTBASLIK
            oSheet.get_Range("A2", "D2").Font.Bold = true;
            oSheet.get_Range("A2", "D2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;




            //excele yazdırma
            int syc = 0;
            foreach (var item in secilenSinif)
            {
                for (int x = 0; x < item.sinifKontenjan; x++)
                {
                    if (syc >= RandomOgrenciler.Count)
                    {
                        break;
                    }
                    oSheet.Name = item.sinifAdi;
                    oSheet.Cells[1, 1] = " Kırklareli Üniversitesi" + dersadi + " " + sinavtipi + " " + sinavtarihi + " " + sinavsaati + " " + hocaadi + " " + item.sinifAdi + " " + "Numaralı Sınıf";
                    oSheet.Cells[x + 3, 1] = x + 1;
                    oSheet.Cells[x + 3, 2] = RandomOgrenciler[syc].Numara;
                    oSheet.Cells[x + 3, 3] = RandomOgrenciler[syc].AdSoyad;
                    syc++;
                }
                oWB.SaveAs(Server.MapPath("~/SınıfListeleri/" + "SINIFADI" + item.sinifAdi + ".xlsx"));
            }
            oWB.Close();


            if (KaydetDevamEt.Visible == true)
            {
                Ara.Visible = true;
                GirisKagidi.Visible = true;
            }
            else
            {
                Ara.Visible = false;
                GirisKagidi.Visible = false;
            }

        }


        protected void Ara_Click(object sender, EventArgs e)
        {

            string adsoyad = "";
            string sinif = "";
            string numara = GirisKagidi.Text;

            SqlConnection cnn;
            string cnntr = ConfigurationManager.ConnectionStrings["kelebekbaglan"].ConnectionString;
            cnn = new SqlConnection(cnntr);


            SqlDataReader dr;
            SqlCommand kmt = new SqlCommand("Select OgrenciAdSoyad, OgrenciSinif, OgrenciNumara from YerlestirmeTablosu where OgrenciNumara = @no", cnn);
            kmt.Parameters.AddWithValue("@no", numara);
            cnn.Open();
            dr = kmt.ExecuteReader();
            if (dr.Read())
            {
                adsoyad = dr.GetValue(0).ToString();
                sinif = dr.GetValue(1).ToString();
                numara = dr.GetValue(2).ToString();
            }
            cnn.Close();

            if (adsoyad == "")
            {
                MessageBox.Show("Lutfen gecerli bir ogrenci no girin", "Hata");
            }
            else
            {
                Bitmap resim = new Bitmap(Server.MapPath("~/resimler/klu.jpg"));
                Graphics yaziyaz = Graphics.FromImage(resim);

                string yazi3 = "Dersin Adı:" + dersadi;
                string yazi4 = "Sınav Tipi:" + sinavtipi;
                string yazi5 = "Sınav Tarihi:" + sinavtarihi;
                string yazi6 = "Sınav Saati:" + sinavsaati;
                string yazi = "Öğrenci Ad Soyad:" + adsoyad;
                string yazi2 = "Öğrenci Numarası:" + numara;
                string yazi7 = "Sınıf:" + sinif;

                yaziyaz.DrawString(yazi3, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 440));
                yaziyaz.DrawString(yazi4, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 490));
                yaziyaz.DrawString(yazi5, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 340));
                yaziyaz.DrawString(yazi6, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 390));
                yaziyaz.DrawString(yazi, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 540));
                yaziyaz.DrawString(yazi2, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 590));
                yaziyaz.DrawString(yazi7, new Font("Arial", 10, FontStyle.Bold), SystemBrushes.WindowText, new Point(10, 640));

                resim.Save(Server.MapPath("~/resimler/"+numara+".jpg"));
                resim.Dispose();
                yaziyaz.Dispose();
            }
        }
    }
}