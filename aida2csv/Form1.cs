using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace aida2csv
{
    public partial class Form1 : Form
    {
        private object openFileDialog1;

        public Form1()
        {
            InitializeComponent();
        }
        //функция поиска в тексте
        public string textfromtag(string source, string begino, string finalo)
        {
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

            string path = @"c:\!\";
            string[] tab0 = File.ReadAllLines(path+"words.txt", Encoding.UTF8);

            //using (
            //FileStream fs = File.Open(path + "out.csv",FileMode.Create);
            FileStream fs = File.Open(path + "out.csv", FileMode.Append);
            FileInfo filer = new FileInfo(path + "out.csv");
            //  )
            //{
            StreamWriter sw = new StreamWriter(fs,Encoding.GetEncoding("Windows-1251"));
            
            string str = "Файл,Операционная система, Пакет обновления ОС, Имя компьютера,Имя пользователя, SMTP -адрес e - mail,Вход в домен,Тип ЦП,Системная плата,Системная память, Монитор, Дисковый накопитель,Первичный адрес IP,Первичный адрес MAC";
            if (filer.Length==0) sw.WriteLine(str);
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
                    try
                    {
                        sw.Write(file);
                        foreach(string word_n in tab0)
                        {
                            sw.Write(",");
                            //string file_body = "";
                            
                            //
                       
                                foreach (string file_body in file_html)
                                {
                                // sw.Write(","+file_body);
                                //string res = textfromtag(file_body, word_n, Environment.NewLine);
                                //  textBox1.Text+= res;
                                //string res = Regex.Match(file_body, word_n+@"(.*?| (\s ?|\S ?))"+Environment.NewLine).ToString();
                               // string file_body0 = Regex.Replace(file_body, "<[^>]+>", string.Empty);
                                int begn = file_body.IndexOf(word_n);
                                int ends = file_body.Length;

                                if (begn > 0)
                                {
                                    
                                    string res = file_body.Substring(begn, ends-begn);
                                    res = Regex.Replace(res, "<[^>]+>", string.Empty);
                                    res = Regex.Replace(res, ",", " ");
                                    begn = res.LastIndexOf("&nbsp;&nbsp;")+ ("&nbsp;&nbsp;").Length;
                                    ends = res.Length;
                                    res = res.Substring(begn, ends - begn);
                                    res.Replace(word_n, "_");
                                    //sw.Write("," +res.ToString());
                                    sw.Write(res.ToString());
                                    break;
                                }
                                }
                            
                            
                            //

                            
                        }


                        sw.WriteLine();
                        //  PictureBox pb = new PictureBox();
                        //  Image loadedImage = Image.FromFile(file);
                        //  pb.Height = loadedImage.Height;
                        //   pb.Width = loadedImage.Width;
                        //    pb.Image = loadedImage;
                        //  flowLayoutPanel1.Controls.Add(pb);
                    }
                    catch (Exception ex)
                    {
                        // Could not load the image - probably related to Windows file system permissions.
                        //      MessageBox.Show("Cannot display the image: " + file.Substring(file.LastIndexOf('\\'))
                        //        + ". You may not have permission to read the file, or " +
                        //       "it may be corrupt.\n\nReported error: " + ex.Message);
                    }
                }
                textBox1.Text+= a;
                sw.Close();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.SelectionStart = textBox1.Text.Length;
            textBox1.ScrollToCaret();
        }
    }
}
