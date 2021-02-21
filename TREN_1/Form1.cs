using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Globalization;


namespace TREN_1
{
    public partial class Form1 : Form
    {
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        String TaskFromFsspApi = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            //label1.Text = textBox1.Text;
            

            /*
             
            string[] dirs = Directory.GetFiles(@"C:\Users\Zhukea\Downloads\Telegram Desktop\", "*", SearchOption.AllDirectories);
            foreach (string FL in dirs)
            {
                listBox1.Items.Add(FL);
            }

            
            listBox1.Items.AddRange(Directory.GetFiles(@"C:\Users\Zhukea\Downloads\Telegram Desktop\", "*"));

            listBox1.Items.Add(Path.GetFileName(@"C:\Users\Zhukea\Downloads\Telegram Desktop\3 случай.txt")); // 3 случай.txt
            
            listBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(@"C:\Users\Zhukea\Downloads\Telegram Desktop\3 случай.txt")); // 3 случай
            
            */
            
            DirectoryInfo dir = new DirectoryInfo(textBox1.Text);
            foreach (FileInfo files in dir.GetFiles("*", SearchOption.AllDirectories))
            {
                listBox1.Items.Add(files.Name); //показывает имена файлов с расширением
                listBox1.Items.Add(Path.GetFileNameWithoutExtension(files.FullName)); //показывает имена файлов без расширений
            }

            /*
            

            // выводит полностью весь путь вложенной папки
            string[] dirs = Directory.GetDirectories(@"C:\Users\Zhukea\Downloads\Telegram Desktop\", "*");
            foreach (string FL in dirs)
            {
                listBox1.Items.Add(FL);
            }

            

            string rootFolder = @"C:\Users\Zhukea\Downloads\Telegram Desktop\";
            foreach (var file in Directory.EnumerateFiles(rootFolder, "*", SearchOption.AllDirectories))
            {
                //File.AppendAllText(file, "текст для добавления");
                listBox1.Items.Add(file);
            }

            

            // выводит все директории в пути
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\Zhukea\Downloads\Telegram Desktop\");
            
            DirectoryInfo[] directories =  di.GetDirectories("*", SearchOption.AllDirectories);
            foreach (var DR in directories)
            {
                listBox1.Items.Add(DR);
            }

            

            // выводит все файлы, в том числе и вложенные
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\Zhukea\Downloads\Telegram Desktop\");
            FileInfo[] files = di.GetFiles("*", SearchOption.AllDirectories);
            foreach (var FL in files)
            {
                listBox1.Items.Add(FL);
            }
            */


        }

        private void button2_Click(object sender, EventArgs e)
        {            
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.Description = "Выберите директорию";

            if (dlg.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = dlg.SelectedPath;                    
                }





        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.Description = "Выберите директорию";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dlg.SelectedPath;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            label1.Text = textBox1.Text;
            DirectoryInfo dir = new DirectoryInfo(textBox1.Text);
            foreach (FileInfo files in dir.GetFiles("*", SearchOption.AllDirectories))
            {
                listBox1.Items.Add(files.Name); //показывает имена файлов с расширением
                listBox1.Items.Add(Path.GetFileNameWithoutExtension(files.FullName)); //показывает имена файлов без расширений
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
                      
            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                for (int j = 0; j < listBox3.Items.Count; j++)
                {
                    if (listBox2.Items[i] == listBox3.Items[j])
                    listBox4.Items.Add(listBox3.Items[j]);
                    //MessageBox.Show("listBox4.Items.Count = "+ listBox4.Items.Count.ToString()+ "   listBox3.Items.Count = "+ listBox3.Items.Count.ToString());
                }
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
            bool flag = false;

            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                for (int j = 0; j < listBox3.Items.Count; j++)
                {
                    if (listBox2.Items[i] == listBox3.Items[j])
                    { flag = true; break; }
                    else flag = false;                    
                }
                
                if (flag == false) listBox4.Items.Add(listBox2.Items[i]);
            }

            for (int i = 0; i < listBox3.Items.Count; i++)
            {
                for (int j = 0; j < listBox2.Items.Count; j++)
                {
                    if (listBox3.Items[i] == listBox2.Items[j])
                    { flag = true; break; }
                    else flag = false;
                }

                if (flag == false) listBox4.Items.Add(listBox3.Items[i]);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                listBox4.Items.Add(listBox2.Items[i]);
            }

            bool flag = false;
            for (int i = 0; i < listBox3.Items.Count; i++)
            {
                for (int j = 0; j < listBox2.Items.Count; j++)
                {
                    if (listBox3.Items[i] == listBox2.Items[j])
                    { flag = true; break; }
                    else flag = false;
                }

                if (flag == false) listBox4.Items.Add(listBox3.Items[i]);
            }



        }

        private void button6_Click(object sender, EventArgs e)
        {
            /*
            using var con = new NpgsqlConnection("Host=172.17.75.4;Username=postgres;Password=postgres;Database=ums");
            con.Open();

            //var sql = "SELECT version()";
            var sql = "select last_name from visitors where id = 965449"; //СТАРОСТИН

            using var cmd = new NpgsqlCommand(sql, con);

            var version = cmd.ExecuteScalar().ToString();
            label3.Text=($"Значение: {version}");
            */


            NpgsqlConnection con = new NpgsqlConnection("Host=172.17.75.4;Username=postgres;Password=postgres;Database=ums");
            con.Open();
            //string sql = ("select * from visitors where id = 965449");
            string sql = ("select visit_date as Дата_визита, last_name as Фамилия, first_name as Имя, patronymic as Отчество, number as Номер_документа , portrait_image_format as Формат " +
                          $"from visitors inner join identity_documents on visitors.id = identity_documents.visitor_id where court_object_id = 176 and visit_date >= '{dateTimePicker1.Text}'");

            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, con);
            ds.Reset();
            da.Fill(ds);
            dt = ds.Tables[0];
            dataGridView1.DataSource = dt;
            con.Close();

            label3.Text = dateTimePicker1.Text;


        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();

            //Отобразить Excel
            ex.Visible = true;

            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 2;

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            //Название листа (вкладки снизу)
            sheet.Name = "Отчет за 13.12.2017";

            //Пример заполнения ячеек

            label3.Text = "СТОЛБЦОВ: "+dataGridView1.ColumnCount.ToString()+"         СТРОК: "+dataGridView1.RowCount.ToString();
            


             
            
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    sheet.Cells[i+1, j+1] = String.Format(dataGridView1.Rows[i].Cells[j].Value.ToString());
            }


            /*   
             
            //Захватываем диапазон ячеек
            Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[8, 8]);

            //Шрифт для диапазона
            range1.Cells.Font.Name = "Tahoma";

            //Размер шрифта для диапазона
            range1.Cells.Font.Size = 10;

            //Захватываем другой диапазон ячеек
            Excel.Range range2 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 2]);
            range2.Cells.Font.Name = "Times New Roman";

            //Задаем цвет этого диапазона. Необходимо подключить System.Drawing
            range2.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);

            //Фоновый цвет
            range2.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xFF, 0xFF, 0xCC));
            */
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            string STR = "{\n" +
                "\t\"token\": \"aMYRXmjGwPFm\",\n" +
                "\t\"request\": [\n" +
                "\t {\n" +
                "\t  \"type\": 1,\n" +
                "\t  \"params\": {\n" +
                "\t     \"firstname\": \""+textBox3.Text+"\",\n" +
                "\t     \"lastname\": \""+textBox4.Text+"\",\n" +
                "\t     \"secondname\": \""+textBox5.Text+"\",\n" +
                "\t     \"region\": \"77\",\n" +
                "\t     \"birthdate\": \""+textBox8.Text+"\"\n" +
                "\t   }\n" +
                "\t }\n" +
                "\t]\n" +
                "}";

            richTextBox2.Text = STR;


            
            WebRequest request = WebRequest.Create("https://api-ip.fssprus.ru/api/v1.0/search/group");
            
            request.Method = "POST";

            byte[] DATA = Encoding.UTF8.GetBytes(STR);
            //byte[] DATA = Encoding.Unicode.GetBytes(STR);

            //byte[] unicodeBytes = Encoding.unicode.GetBytes(unicodeString);
            //byte[] asciiBytes = Encoding.Convert(Encoding.Unicode, Encoding.ASCII, DATA);

            request.ContentType = "application/json; charset=utf-8";
            
            request.ContentLength = DATA.Length;
            
            Stream dataStream = request.GetRequestStream();
            
            dataStream.Write(DATA, 0, DATA.Length);
            
            dataStream.Close();

            
            WebResponse response = request.GetResponse();
            
            textBox7.Text = ((HttpWebResponse)response).StatusDescription;

            using (dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                richTextBox4.Text = responseFromServer;

                int A = responseFromServer.IndexOf("task");
                label12.Text = responseFromServer.Substring(A + 7, 36);
                TaskFromFsspApi = responseFromServer.Substring(A + 7, 36);
            }
            
            response.Close();            

        }
                

        private void button10_Click(object sender, EventArgs e)
        {
            /*
            WebRequest request = WebRequest.Create("https://api-ip.fssprus.ru/api/v1.0/result?token=aMYRXmjGwPFm&task=" + TaskFromFsspApi );
            WebResponse response = await request.GetResponseAsync();
            using (Stream stream = response.GetResponseStream())
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    richTextBox3.Text = reader.ReadToEnd();
                }
            }
            response.Close();
            */
            //Вариант №2 - тоже рабочий 

            var request = WebRequest.Create("https://api-ip.fssprus.ru/api/v1.0/result?token=aMYRXmjGwPFm&task=" + TaskFromFsspApi);
            
            var response = request.GetResponse();

            textBox9.Text = (((HttpWebResponse)response).StatusDescription);

            //var encodedBytes = Encoding.UTF8.GetBytes(response);
            //Encoding.Convert(Encoding.UTF8, Encoding.Unicode, response);

            /*
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            richTextBox3.Text = responseString;

            var result = Regex.Replace(responseString, @"\\[Uu]([0-9A-Fa-f]{4})", m => char.ToString((char)ushort.Parse(m.Groups[1].Value, NumberStyles.AllowHexSpecifier)));
            richTextBox3.Text = result;

            */

            
            using (Stream dataStream = response.GetResponseStream())
            {             
                StreamReader reader = new StreamReader(dataStream);                
                string responseFromServer = reader.ReadToEnd();
                var result = Regex.Replace(responseFromServer, @"\\[Uu]([0-9A-Fa-f]{4})", m => char.ToString((char)ushort.Parse(m.Groups[1].Value, NumberStyles.AllowHexSpecifier)));
                richTextBox3.Text = result;
            }                        
            response.Close();
            



        }

        private void button11_Click(object sender, EventArgs e)
        {
            string AllText = richTextBox5.Text;

            int A = AllText.IndexOf('[');
            int B = AllText.LastIndexOf(']');
            AllText = AllText.Remove(B);
            AllText = AllText.Remove(0,A+1);            
            richTextBox5.Text = AllText.Trim();

            int A1 = AllText.IndexOf('[');
            int B1 = AllText.LastIndexOf(']');
            AllText = AllText.Remove(B1+1);
            AllText = AllText.Remove(0, A1-1);


            for (int i = 0; i <= 3; i++)
            { 
                int A2 = AllText.IndexOf('[');
                int B2 = AllText.IndexOf(']');
                string NEGODYAY = AllText.Substring(A2, B2+1);
                richTextBox6.Text = NEGODYAY;

                               
                int tochkaOtscheta = 0;
                decimal SUMMA;
                decimal ZADOLZHENOST = 0;                

                while (NEGODYAY.IndexOf(" руб", tochkaOtscheta) != -1)
                {
                    int RUB = NEGODYAY.IndexOf(" руб", tochkaOtscheta);
                    int DVOETOCH = NEGODYAY.LastIndexOf(':', RUB);
                    string DOLG = NEGODYAY.Substring(DVOETOCH + 2, RUB - DVOETOCH - 2);
                    SUMMA = Decimal.Parse(DOLG.Replace('.', ','));
                    ZADOLZHENOST += SUMMA;                 
                    tochkaOtscheta = RUB + 1;
                    
                }
                listBox5.Items.Add(ZADOLZHENOST.ToString());


                AllText = AllText.Remove(0, B2+1);
                if (AllText != "")
                {
                    AllText = AllText.Remove(0, AllText.IndexOf('['));
                }
            }

         
           
        }
    }
}
