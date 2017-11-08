using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
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

        private void button1_Click(object sender, EventArgs e)
        {//получаем список полей для вібора из хтмл

            string path = @"c:\!\";
            string[] tab0 = File.ReadAllLines(path+"words.txt", Encoding.UTF8);

            //using (
            FileStream fs = File.Open(path + "out.txt",FileMode.Create);
              //  )
                //{
                StreamWriter sw = new StreamWriter(fs);             
                
            //}

            //получаем список файлов и обрабатіваем их
           
            // open file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Text Files|*.html";
            openFileDialog1.Title = "файл для обработки";
            openFileDialog1.FileName = "*.html";
            openFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
           
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                
                string a = "";
               foreach (String file in openFileDialog1.FileNames)
                {
                    a += file+ Environment.NewLine;
                    try
                    {
                        sw.WriteLine(file);
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
                textBox1.Text = a;
                sw.Close();
            }
        }
    }
}
