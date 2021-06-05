using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace aida2csv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //функция поиска в тексте
        public string textfromtag(string source, string begino, string finalo)
        {//функция получения данных из тегов
            string[] res = { };
            int i = 0;
            finalo.Replace("/", @"\/");
            /////////
            //Regex Reg = new Regex(@"""" + begino + ""(.*?)"" + finalo+"""");
            string quer = @begino + @"(.*?|(\s?|\S?))" + finalo;
            Regex Reg = new Regex(quer);
            textBox1.Text += quer;
            MatchCollection reHref = Reg.Matches(source);
            textBox1.Text += source;
            foreach (Match match in reHref)
            {
                Array.Resize<string>(ref res, i + 1);
                res[i] = match.ToString();
               
                res[i] = res[i].Remove(0, begino.Length);
                res[i] = res[i].Remove(res[i].Length - finalo.Length, finalo.Length);
                i++; textBox1.Text += i.ToString();
           //     textBox7.Text += i.ToString() + "=";
            }
            /////////        
            return res[0];
        }
        //
        private void button1_Click(object sender, EventArgs e)
        {//получаем список полей для вібора из хтмл
            int flag = 0;
            string path = @"c:\!invRep\";
            string str="", path_words=path+"words.txt";
            string[] str_ar_file = { "Кабинет","Имя комп","Имя польз.инв.","Перифирия","Лицензия","Корпус" };
            string[] str_ar = {"Операционная система", "Пакет обновления ОС", "Имя компьютера","Имя пользователя", "Вход в домен","Тип ЦП", "Системная плата","Системная память", "Имя монитора", "ID монитора", "Дисковый накопитель", "Аппаратный адрес", "IP / маска подсети", "Шлюз", "Имя принтера" };
            //string[] tab0 = File.ReadAllLines(path+"words.txt", Encoding.UTF8);
            //using (
            //FileStream fs = File.Open(path + "out.csv",FileMode.Create);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            //если файла words.txt нет, то создаем свой по умолчанию
            if (!File.Exists(path_words))
            {   FileStream fs_txt = File.Open(path_words, FileMode.Append);
                using (StreamWriter sw_txt = new StreamWriter(fs_txt, Encoding.GetEncoding("Windows-1251")))
                {for(int i=0;i<str_ar.Length;i++)
                        sw_txt.WriteLine(str_ar[i]);}
            }
            //считываем поля из words.txt
            string[] tab0 = File.ReadAllLines(path_words, Encoding.Default);
            if (!File.Exists(path + "out.csv")) flag = 1;
                FileStream fs = File.Open(path + "out.csv", FileMode.Append);
                        //  )
            //{
            StreamWriter sw = new StreamWriter(fs,Encoding.GetEncoding("Windows-1251"));            
            if (flag == 1)//если файла вообще небыло - формируем заголовок
            {
                for (int i = 0; i < str_ar_file.Length; i++) str += str_ar_file[i] + ";";                
                for (int i = 0; i < str_ar.Length; i++) str += str_ar[i]+";";
                sw.WriteLine(str);
            }
            //str = "Файл; Операционная система; Пакет обновления ОС; Имя компьютера;Имя пользователя; SMTP -адрес e - mail;Вход в домен;Тип ЦП; Системная плата;Системная память; Монитор; Дисковый накопитель;Первичный адрес IP;Первичный адрес MAC";
            //sw.WriteLine(str);
            //}

            //получаем список файлов и обрабатіваем их

            // open file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Text Files|*.htm";
            openFileDialog1.Title = "файл для обработки";
            openFileDialog1.FileName = "*.htm";
            openFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
           
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {                
                string a = "";
               foreach (String file in openFileDialog1.FileNames)
                {
                    a += file+ Environment.NewLine;
                    string[] file_html = File.ReadAllLines(file,Encoding.GetEncoding(1251));
                    //читаем содержимое файла
                    try
                    {   //sw.Write(file);//вместо єтого расшифровать имя файла в стобцы
                        string[] data_s = file.Split('~');
                        //из имени файла извлекаем номер кабинета
                        int n_slash = data_s[0].LastIndexOf(@"\");
                        data_s[0] = data_s[0].Substring(n_slash+1, data_s[0].Length - n_slash-1);
                        str = "";
                        for (int i = 0; i < data_s.Length; i++) str += data_s[i] + ";";
                        sw.Write(str);
                        foreach (string word_n in tab0)
                        {//перебор ключевых слов их таблицы
                            //sw.Write(",");
                            //sw.Write(";");
                            //string file_body = "";                          
                            //
                            foreach (string file_body in file_html)
                                {//читаем строку из файла-отчета
                                 // sw.Write(","+file_body);
                                 //string res = textfromtag(file_body, word_n, Environment.NewLine);
                                 //  textBox1.Text+= res;
                                 //string res = Regex.Match(file_body, word_n+@"(.*?| (\s ?|\S ?))"+Environment.NewLine).ToString();
                                 // string file_body0 = Regex.Replace(file_body, "<[^>]+>", string.Empty);
                                string word_n2 = word_n + "&nbsp;&nbsp;";
                                int begn = file_body.IndexOf(word_n2);
                                //int begn = file_body.LastIndexOf(word_n);
                                int ends = file_body.Length;

                                if (begn > 0)
                                {
                                    string res = file_body.Substring(begn, ends-begn);
                                //    begn = res.LastIndexOf(word_n);
                                  //  ends = res.Length;
                                    //res = res.Substring(begn, ends - begn);
                                    //if (res.IndexOf(word_n) > 2) res = file_body.Substring(res.IndexOf(word_n), ends - begn);
                                    
                                    //MessageBox.Show(res.LastIndexOf(word_n).ToString());
                                    res = Regex.Replace(res, "<[^>]+>", string.Empty);
                                    res = Regex.Replace(res, ",", " ");
                                    begn = res.LastIndexOf("&nbsp;&nbsp;")+ ("&nbsp;&nbsp;").Length;
                                    ends = res.Length;
                                    res = res.Substring(begn, ends - begn);
                                    res = Regex.Replace(res, "&nbsp;", "");
                                    res = Regex.Replace(res, "&gt;", "");
                                    res = Regex.Replace(res, "&lt;", "");
                                    //res.Replace(word_n, "_");
                                    //sw.Write("," +res.ToString());
                                   // MessageBox.Show(res + "=" + word_n);
                                    sw.Write(res.ToString());
                                    break;
                                }}
                            sw.Write(";"); //                            
                        }
                        sw.WriteLine();
                    }
                    catch (Exception ex)
                    {//тут можем обработать вывод об ощибке открытия                       
                        MessageBox.Show(ex.Message);
                    }
                }
                textBox1.Text+= a;
                sw.Close();
                //открываем папку с результами
                   Process.Start(path);
                //или так чтобы папка должна быть выбрана в проводнике....
                //Process PrFolder = new Process();
                //ProcessStartInfo psi = new ProcessStartInfo();
                //psi.CreateNoWindow = true;
                //psi.WindowStyle = ProcessWindowStyle.Normal;
                //psi.FileName = "explorer";
                //psi.Arguments = @"/n, /select, " + path;
                //PrFolder.StartInfo = psi;
                //PrFolder.Start();
            }
            Application.Exit();
        }        
    }
}
