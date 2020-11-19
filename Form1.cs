using System;
using System.Windows.Forms;
using System.IO;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Globalization;
using Spire.Doc;

namespace PhotoTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string folderName = folderBrowserDialog1.SelectedPath;//шлях до папки
            int count = Convert.ToInt32(new DirectoryInfo(folderName).GetFiles().Length.ToString());//кількість фото в папці
        }

        public void button2_Click(object sender, EventArgs e)
        {



            string folderName = folderBrowserDialog1.SelectedPath;//шлях до папки
            int count = Convert.ToInt32(new DirectoryInfo(folderName).GetFiles().Length.ToString());//кількість фото в папці
            string n = textBox1.Text;
            string day = comboBox1.Text;
            string month = comboBox2.Text;
            string year = comboBox3.Text;
            string[] files = Directory.GetFiles(folderName);
            string[] allfiles = Directory.GetFiles(folderName);



            // создаём документ
            DocX document = DocX.Create(folderName + "/Фототаблиця № " + n + " від " + day + " " + month + " " + year + " року" + ".docx");


            // создаём таблицу с 2 строками и 1 столбцом
            Xceed.Document.NET.Table table = document.AddTable(count, 1);
            table.AutoFit = AutoFit.Window;
            table.Alignment = Alignment.center;
            table.Design = TableDesign.TableGrid;
           
                for (int j = count - 1; j >= 0; j--)
                {
                    if (j % 2 == 0)
                    {
                        table.Rows[j].Cells[0].Paragraphs[0].Append("\n" + "Фототаблиця № " + ((j + 2) / 2) + " до автотоварознавчого" + "\n" + " дослідження № " + n + " від «" + day + "» " + month + " " + year + " року.\n");
                        table.Rows[j].Cells[0].Paragraphs[0].FontSize(14);
                        table.Rows[j].Cells[0].Paragraphs[0].Bold();
                        table.Rows[j].Cells[0].Paragraphs[0].Italic();
                        table.Rows[j].Cells[0].Paragraphs[0].Font("Times New Roman");
                        table.Rows[j].Cells[0].Paragraphs[0].Alignment = Alignment.center;

                    }
                    else
                    {

                        // Create a Picture (Custom View) of this Image.
                        Picture p = document.AddImage(allfiles[j - 1]).CreatePicture();
                        Picture x = document.AddImage(allfiles[j]).CreatePicture();
                        p.Width = 368;
                        x.Width = 368;
                        p.Height = 246;
                        x.Height = 246;
                        string fototitle;
                        if (checkBox1.Checked == true)
                        {
                            fototitle = "Вигляд пошкоджень";
                        }
                        else { fototitle = ""; }
                        table.Rows[j].Cells[0].Paragraphs[0].Append("\n                                                                                                                                  Фото № " + (j)).Bold().Italic().AppendPicture(p).Append("\n" + "\n" + fototitle + "\n").Italic().Append("                                                                                                                                  Фото № " + (j + 1) + "\n").Bold().Italic().AppendPicture(x).Append("\n" + fototitle).Italic().Append("\n" + "\n" + "Експерт :                                   ________________     B.В.Вербов.\n\n").Italic().Bold();
                        table.Rows[j].Cells[0].Paragraphs[0].FontSize(12);
                        table.Rows[j].Cells[0].Paragraphs[0].Font("Times New Roman");
                        table.Rows[j].Cells[0].Paragraphs[0].Alignment = Alignment.center;


                    }
                }
            //підрахунок кількості натискань на кнопку
            int button = 0;
            button++;
            if (button >= 3) 
            {
                button2.Visible = false;
                MessageBox.Show("Перевищено ліміт використання");
            }
            else
            {
                document.InsertParagraph().InsertTableAfterSelf(table);
            }


                // сохраняем документ
                document.Save();
            }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //Form2 f = new Form2();
            //f.Show();
        }
    }
    }
