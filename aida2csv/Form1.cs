using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        {
            string[] in_words = { "розклад", "деканфак", "занятьфак", "______" };
            string[] out_words = { "num", "lastname", "name", "sex", "lastnameen", "nameen" };
            string data = @"<?xml version=""1.0"" encoding=""utf-8""?><documents>";
            // open file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Text Files|*.html";
            openFileDialog1.Title = "файл для обработки";
            openFileDialog1.FileName = "*.html";
            openFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
            //   string path = file_path + "Order_2.xml";
            //file_path= file_path + "2.xml";
          //  DialogResult dr = this.openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Create a PictureBox.
                string a = "";
               foreach (String file in openFileDialog1.FileNames)
                {
                    a += file+ Environment.NewLine;
                    try
                    {
                        //  PictureBox pb = new PictureBox();
                        //  Image loadedImage = Image.FromFile(file);
                        //  pb.Height = loadedImage.Height;
                        //   pb.Width = loadedImage.Width;
                        //    pb.Image = loadedImage;
                        //  flowLayoutPanel1.Controls.Add(pb);
                    }
                    catch (SecurityException ex)
                    {
                        // The user lacks appropriate permissions to read files, discover paths, etc.
                        MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                            "Error message: " + ex.Message + "\n\n" +
                            "Details (send to Support):\n\n" + ex.StackTrace
                        );
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
            }
        }
    }
}
