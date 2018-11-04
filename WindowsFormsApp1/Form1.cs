using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<List<Color>> map = new List<List<Color>>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    textBox1.Text = file;
                    pictureBox1.Image = Image.FromFile(file);
                    Bitmap img = new Bitmap(file);
                    for (int i = 0; i < img.Width; i++)
                    {
                        List<Color> row = new List<Color>();
                        for (int j = 0; j < img.Height; j++)
                        {
                            row.Add(img.GetPixel(i, j));
                        }
                        map.Add(row);
                    }
                }
                catch (IOException)
                {
                    MessageBox.Show("Failed to find file");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something went wrong. Error: " + ex.Message);
                }
            }
            
            Console.WriteLine("Path: " + textBox1.Text);
            Console.WriteLine(result); // <-- For debugging use.
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                var worKbooK = excel.Workbooks.Add(Type.Missing);


                var worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;

                int total = map.Count * map[0].Count;

                int soFar = 0;

                for (int i = 0; i < map.Count; i++)
                {
                    for (int j = 0; j < map[i].Count; j++)
                    {
                        soFar++;
                        label2.Text = "Progress: " + soFar + " / " + total;
                        worKsheeT.Cells[i+1, j+1].Interior.Color = System.Drawing.ColorTranslator.ToOle(map[i][j]);// map[i][j];
                    }
                }


                //worKbooK.SaveAs(@"C:\Images\hello.xls");
                worKbooK.SaveAs(Path.ChangeExtension(textBox1.Text, ".xlsx"));
                worKbooK.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worKbooK);
                Marshal.ReleaseComObject(excel);
                MessageBox.Show("Done");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong. Error: " + ex.Message);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
