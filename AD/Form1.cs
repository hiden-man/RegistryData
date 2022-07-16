using Microsoft.Win32;
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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace AD
{
    public partial class Form1 : Form
    {
        Size sSize = Screen.PrimaryScreen.Bounds.Size;
        RegistryKey key;
        InformationWindow IW;
        FileStream FS;

        string path = @"Software\Microsoft\Windows\CurrentVersion\Run",
            path2 = @"C:\RegSaver\Sources\RegistryData.xlsx";

        string sym, sym2;

        int count = 1, countI = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Location = new System.Drawing.Point((sSize.Width/2)-(Size.Width/2),
                (sSize.Height/2)-(Size.Height/2));
            Directory.CreateDirectory(@"C:\RegSaver\Sources");
            WriteData();
            comboBox1.SelectedIndex = 0;
            IW = new InformationWindow();
        }

        // Save Button
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Registry.CurrentUser.CreateSubKey(path).SetValue(textBox1.Text + " (NEW)", textBox2.Text);
            textBox1.Text = "";
            textBox2.Text = "";
            ClearComboList();
            WriteData();

        }
        
        // Select file Button
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            sym2 = "";
            OpenDialogFiles();
        }
       
        // Show Button
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            count = 1;
            countI = 0;
            if (!File.Exists(path2))
            {
                CreateNewFile();
                Excell exc = new Excell(path2, 1);
                key = Registry.CurrentUser.OpenSubKey(path);
                foreach (string valueNames in key.GetValueNames())
                {
                    exc.WriteToCell(countI, 0, valueNames);
                    countI++;
                }

                for (int i = 0; i < countI; i++)
                {
                    richTextBox1.Text += $"{count++}: {exc.ReadCell(i, 0)}\n";
                }
                exc.Save();
                exc.CloseFile();
            }
            else
            {
                countI = 0;
                File.Delete(path2);
                CreateNewFile();
                Excell exc = new Excell(path2, 1);
                key = Registry.CurrentUser.OpenSubKey(path);
                foreach (string valueNames in key.GetValueNames())
                {
                    exc.WriteToCell(countI, 0, valueNames);
                    countI++;
                }
                for (int i = 0; i < countI; i++)
                {
                    richTextBox1.Text += $"{count++}: {exc.ReadCell(i, 0)}\n";
                }
                exc.Save();
                exc.CloseFile();
            }
        }

        // Delete button
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteParameters();
                ClearComboList();
                WriteData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Select parameter", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        // Close button
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Close();
        }

        // Information button
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            IW.Show();
        }


        // Delete parameters Method
        private void DeleteParameters()
        {
            key = Registry.CurrentUser.CreateSubKey(path);
            key.DeleteValue($"{comboBox1.SelectedItem}");

        }
        
        // Open File Dialog Method
        private void OpenDialogFiles()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select file";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "All files (*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.ShowDialog();
            if (ofd.FileName != "")
            {
                char[] ch = ofd.FileName.ToCharArray();
                for (int i = ofd.FileName.Length - 1; i > 0; i--)
                {
                    if (ch[i] != '\\')
                        sym += ch[i];
                    else
                        break;
                }

                char[] ch2 = sym.ToCharArray();

                for (int i = sym.Length - 1; i > 0; i--)
                {
                    if (ch2[i] != '.')
                        sym2 += ch2[i];
                    else
                        break;
                }

                textBox1.Text = sym2;
                textBox2.Text = ofd.FileName;
            }
            else
            {
                textBox1.Text = "File not selected";
                textBox2.Text = "File not selected";
            }
        }
        //------------------------------------------------------------------
        public void OpenFile()
        {
            Excell exc = new Excell(path2, 1);
        }


        private void ClearComboList()
        {
            comboBox1.Items.Clear();
        }

        public void WriteData()
        {
            if (!File.Exists(path2))
            {
                CreateNewFile();
                Excell exc = new Excell(path2, 1);
                key = Registry.CurrentUser.OpenSubKey(path);
                foreach (string valueNames in key.GetValueNames())
                {
                    exc.WriteToCell(countI, 0, valueNames);
                    countI++;
                }

                //Excell exc = new Excell(path5, 1);
                //записує данні у спеціальний клас для роботи з текстовими рядками
                //exc.ReadCell(0, 0);
                //видаляє один символ(1) на першому позиції(0) в тексті

                for (int i = 0; i < countI; i++)
                {
                    comboBox1.Items.Add(exc.ReadCell(i, 0));
                }
                //string[] str = { "dersten", "elder", "ypypy" };
                //for (int i = 0; i < str.Length; i++)
                //    exc.WriteToCell(i, 0, str[i]);
                //exc.WriteToCell(0,0, textBox1.Text);
                exc.Save();
                //exc.SaveAs(path2);
                exc.CloseFile();
            }
            else
            {
                countI = 0;
                File.Delete(path2);
                CreateNewFile();
                Excell exc = new Excell(path2, 1);
                key = Registry.CurrentUser.OpenSubKey(path);
                foreach (string valueNames in key.GetValueNames())
                {
                    exc.WriteToCell(countI, 0, valueNames);
                    countI++;
                }
                for (int i = 0; i < countI; i++)
                {
                    comboBox1.Items.Add(exc.ReadCell(i, 0));
                }
                exc.Save();
                //exc.SaveAs(path2);
                exc.CloseFile();
            }
        }
                //string[] str = { "dersten", "elder", "ypypy" };
                //for (int i = 0; i < str.Length; i++)
                //    exc.WriteToCell(i, 0, str[i]);
                //exc.WriteToCell(0,0, textBox1.Text);
        public void CreateNewFile()
        {
            if (!File.Exists(path2))
            {
                Excell excel = new Excell();
                excel.CreateNewFile();
                excel.SaveAs(path2);
                excel.CloseFile();
            }
            else
            {
                MessageBox.Show("File Exists!");
            }
        }
        //------------------------------------------------------------------


    }

    class Excell
    {
        // простий конструктор 
        public Excell()
        {

        }
        // змінна для зберігання шляху до файлу
        string path = "";

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        // конструктор класу який приймає параметри (шлях та номер листа)
        public Excell(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];

        }
        // функція для зчитування данних з файлу
        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }
        // запис в комірку
        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }
        // зберігання файлу
        public void Save()
        {
            wb.Save();
        }
        // зберігання файлу с іншим іменем
        public void SaveAs(string path1)
        {
            wb.SaveAs(path1);
        }
        // закриття файлу
        public void CloseFile()
        {
            wb.Close();
        }
        // створення нового файлу
        public void CreateNewFile()
        {
            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }
    }
}
