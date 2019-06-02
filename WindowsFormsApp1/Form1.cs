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
        //List<Dictionary<Color,int>> palette = new List<Dictionary<Color,int>>();
        Dictionary<Color, int> palette = new Dictionary<Color, int>();
        Dictionary<Color, Color> cache;
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
                    map.Clear();
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

        private Color findClosestColor(Color c)
        {
            if (cache.ContainsKey(c)) return cache[c];
            var r = c.R;
            var g = c.G;
            var b = c.B;

            int diff = 255*3;
            Color closest = Color.FromArgb(255, 255, 255, 255);
            foreach(var kvp in palette)
            {
                var pR = kvp.Key.R;
                var pG = kvp.Key.G;
                var pB = kvp.Key.B;

                var currDiff = Math.Abs(r - pR) + Math.Abs(g - pG) + Math.Abs(b - pB);
                if (currDiff < diff)
                {
                    diff = currDiff;
                    closest = kvp.Key;
                }
            }
            cache.Add(c, closest);
            return closest;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            //{
            //    label3.Text = "Please select a palette.";
            //    return;
            //}

            try
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                var worKbooK = excel.Workbooks.Add(Type.Missing);


                var worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;

                int total = map.Count * map[0].Count;

                int soFar = 0;
                cache = new Dictionary<Color, Color>();
                for (int i = 0; i < map.Count; i++)
                {
                    for (int j = 0; j < map[i].Count; j++)
                    {
                        soFar++;
                        label2.Text = "Progress: " + soFar + " / " + total;
                        Color toUse = findClosestColor(map[i][j]);
                        worKsheeT.Cells[j + 1, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(toUse);// map[i][j];
                        worKsheeT.Cells[j + 1, i + 1] = palette[toUse];
                    }
                }

                //worKbooK.SaveAs(@"C:\Images\hello.xls");
                worKbooK.SaveAs(Path.ChangeExtension(textBox1.Text, ".xlsx"));
                worKbooK.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worKbooK);
                Marshal.ReleaseComObject(excel);
                label3.Text = "Done";
            }
            catch (Exception ex)
            {
                label3.Text = "Something went wrong. Error: " + ex.Message;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            #region defpalette
            palette.Add(Color.FromArgb(255, 255, 255, 255), 0);
            palette.Add(Color.FromArgb(255, 1, 1, 0), 2);
            palette.Add(Color.FromArgb(255, 246, 176, 76), 3);
            palette.Add(Color.FromArgb(255, 237, 139, 0), 4);
            palette.Add(Color.FromArgb(255, 225, 6, 0), 5);
            palette.Add(Color.FromArgb(255, 186, 12, 47), 6);
            palette.Add(Color.FromArgb(255, 241, 167, 220), 7);
            palette.Add(Color.FromArgb(255, 255, 52, 179), 8);
            palette.Add(Color.FromArgb(255, 219, 33, 82), 9);
            palette.Add(Color.FromArgb(255, 242, 240, 161), 10);
            palette.Add(Color.FromArgb(255, 255, 209, 0), 11);
            palette.Add(Color.FromArgb(255, 173, 220, 145), 12);
            palette.Add(Color.FromArgb(255, 135, 216, 57), 13);
            palette.Add(Color.FromArgb(255, 36, 158, 107), 14);
            palette.Add(Color.FromArgb(255, 0, 124, 88), 15);
            palette.Add(Color.FromArgb(255, 255, 106, 19), 16);
            palette.Add(Color.FromArgb(255, 255, 103, 31), 17);
            palette.Add(Color.FromArgb(255, 170, 220, 235), 18);
            palette.Add(Color.FromArgb(255, 65, 182, 230), 19);
            palette.Add(Color.FromArgb(255, 0, 144, 218), 20);
            palette.Add(Color.FromArgb(255, 0, 51, 153), 21);
            palette.Add(Color.FromArgb(255, 252, 191, 169), 22);
            palette.Add(Color.FromArgb(255, 204, 153, 102), 23);
            palette.Add(Color.FromArgb(255, 255, 231, 128), 24);
            palette.Add(Color.FromArgb(255, 167, 123, 202), 25);
            palette.Add(Color.FromArgb(255, 160, 94, 181), 26);
            palette.Add(Color.FromArgb(255, 51, 0, 114), 27);
            palette.Add(Color.FromArgb(255, 180, 126, 0), 28);
            palette.Add(Color.FromArgb(255, 164, 73, 61), 29);
            palette.Add(Color.FromArgb(255, 122, 62, 44), 30);
            palette.Add(Color.FromArgb(255, 123, 77, 53), 31);
            palette.Add(Color.FromArgb(255, 92, 71, 56), 32);
            palette.Add(Color.FromArgb(255, 155, 155, 155), 33);
            palette.Add(Color.FromArgb(255, 118, 119, 119), 34);
            palette.Add(Color.FromArgb(255, 160, 159, 157), 35);
            palette.Add(Color.FromArgb(255, 201, 128, 158), 36);
            palette.Add(Color.FromArgb(255, 20, 123, 209), 37);
            palette.Add(Color.FromArgb(255, 105, 179, 231), 38);
            palette.Add(Color.FromArgb(255, 153, 214, 234), 39);
            palette.Add(Color.FromArgb(255, 206, 220, 0), 40);
            palette.Add(Color.FromArgb(255, 246, 235, 97), 41);
            palette.Add(Color.FromArgb(255, 250, 224, 83), 42);
            palette.Add(Color.FromArgb(255, 165, 0, 52), 43);
            palette.Add(Color.FromArgb(255, 255, 163, 139), 44);
            palette.Add(Color.FromArgb(255, 93, 219, 93), 45);
            palette.Add(Color.FromArgb(255, 243, 234, 93), 46);
            palette.Add(Color.FromArgb(255, 243, 207, 179), 47);
            palette.Add(Color.FromArgb(255, 255, 199, 44), 48);
            palette.Add(Color.FromArgb(255, 236, 134, 208), 49);
            palette.Add(Color.FromArgb(255, 197, 180, 227), 50);
            palette.Add(Color.FromArgb(255, 252, 251, 205), 51);
            palette.Add(Color.FromArgb(255, 74, 31, 135), 52);
            palette.Add(Color.FromArgb(255, 115, 211, 60), 53);
            palette.Add(Color.FromArgb(255, 0, 178, 169), 54);
            palette.Add(Color.FromArgb(255, 108, 194, 74), 55);
            palette.Add(Color.FromArgb(255, 136, 139, 141), 56);
            palette.Add(Color.FromArgb(255, 188, 4, 35), 57);
            palette.Add(Color.FromArgb(255, 5, 8, 73), 58);
            palette.Add(Color.FromArgb(255, 83, 26, 35), 59);
            palette.Add(Color.FromArgb(255, 158, 229, 176), 60);
            palette.Add(Color.FromArgb(255, 241, 235, 156), 61);
            palette.Add(Color.FromArgb(255, 252, 63, 63), 62);
            palette.Add(Color.FromArgb(255, 234, 190, 219), 63);
            palette.Add(Color.FromArgb(255, 165, 0, 80), 64);
            palette.Add(Color.FromArgb(255, 239, 129, 46), 65);
            palette.Add(Color.FromArgb(255, 252, 108, 133), 66);
            palette.Add(Color.FromArgb(255, 177, 78, 181), 67);
            palette.Add(Color.FromArgb(255, 105, 19, 238), 68);
            palette.Add(Color.FromArgb(255, 35, 40, 43), 69);
            palette.Add(Color.FromArgb(255, 24, 48, 40), 70);
            palette.Add(Color.FromArgb(255, 234, 170, 0), 71);
            palette.Add(Color.FromArgb(255, 255, 197, 110), 72);
            palette.Add(Color.FromArgb(255, 184, 97, 37), 73);
            palette.Add(Color.FromArgb(255, 205, 178, 119), 74);
            palette.Add(Color.FromArgb(255, 181, 129, 80), 75);
            palette.Add(Color.FromArgb(255, 255, 109, 106), 76);
            palette.Add(Color.FromArgb(255, 170, 87, 97), 77);
            palette.Add(Color.FromArgb(255, 92, 19, 27), 78);
            palette.Add(Color.FromArgb(255, 89, 213, 216), 79);
            palette.Add(Color.FromArgb(255, 0, 174, 199), 80);
            palette.Add(Color.FromArgb(255, 72, 169, 197), 81);
            palette.Add(Color.FromArgb(255, 0, 174, 214), 82);
            palette.Add(Color.FromArgb(255, 0, 133, 173), 83);
            palette.Add(Color.FromArgb(255, 155, 188, 17), 84);
            palette.Add(Color.FromArgb(255, 153, 155, 48), 85);
            palette.Add(Color.FromArgb(255, 0, 133, 34), 86);
            palette.Add(Color.FromArgb(255, 239, 239, 239), 87);
            palette.Add(Color.FromArgb(255, 209, 209, 209), 88);
            palette.Add(Color.FromArgb(255, 187, 188, 188), 89);
            palette.Add(Color.FromArgb(255, 72, 73, 85), 90);
            palette.Add(Color.FromArgb(255, 22, 185, 71), 91);
            palette.Add(Color.FromArgb(255, 218, 182, 152), 92);
            palette.Add(Color.FromArgb(255, 244, 169, 153), 93);
            palette.Add(Color.FromArgb(255, 238, 125, 103), 94);
            palette.Add(Color.FromArgb(255, 240, 134, 97), 95);
            palette.Add(Color.FromArgb(255, 212, 114, 42), 96);
            palette.Add(Color.FromArgb(255, 100, 172, 223), 97);
            palette.Add(Color.FromArgb(255, 100, 194, 220), 98);
            palette.Add(Color.FromArgb(255, 79, 159, 179), 99);
            palette.Add(Color.FromArgb(255, 49, 150, 221), 100);
            palette.Add(Color.FromArgb(255, 27, 108, 182), 101);
            palette.Add(Color.FromArgb(255, 8, 57, 128), 102);
            palette.Add(Color.FromArgb(255, 10, 102, 139), 103);
            palette.Add(Color.FromArgb(255, 8, 91, 110), 104);
            palette.Add(Color.FromArgb(255, 0, 78, 120), 105);
            palette.Add(Color.FromArgb(255, 0, 85, 116), 106);
            palette.Add(Color.FromArgb(255, 204, 190, 128), 107);
            palette.Add(Color.FromArgb(255, 164, 147, 80), 108);
            palette.Add(Color.FromArgb(255, 158, 136, 60), 109);
            palette.Add(Color.FromArgb(255, 118, 108, 43), 110);
            palette.Add(Color.FromArgb(255, 121, 95, 38), 111);
            palette.Add(Color.FromArgb(255, 186, 184, 162), 112);
            palette.Add(Color.FromArgb(255, 114, 140, 84), 113);
            palette.Add(Color.FromArgb(255, 126, 124, 68), 114);
            palette.Add(Color.FromArgb(255, 100, 105, 46), 115);
            palette.Add(Color.FromArgb(255, 78, 88, 44), 116);
            palette.Add(Color.FromArgb(255, 74, 94, 45), 117);
            palette.Add(Color.FromArgb(255, 113, 196, 182), 118);
            palette.Add(Color.FromArgb(255, 102, 204, 153), 119);
            palette.Add(Color.FromArgb(255, 86, 154, 131), 120);
            palette.Add(Color.FromArgb(255, 20, 194, 91), 121);
            palette.Add(Color.FromArgb(255, 2, 168, 24), 122);
            palette.Add(Color.FromArgb(255, 4, 85, 46), 123);
            palette.Add(Color.FromArgb(255, 19, 107, 90), 124);
            palette.Add(Color.FromArgb(255, 5, 70, 65), 125);
            palette.Add(Color.FromArgb(255, 217, 182, 214), 126);
            palette.Add(Color.FromArgb(255, 173, 98, 164), 127);
            palette.Add(Color.FromArgb(255, 230, 140, 163), 128);
            palette.Add(Color.FromArgb(255, 222, 84, 121), 129);
            palette.Add(Color.FromArgb(255, 158, 130, 186), 130);
            palette.Add(Color.FromArgb(255, 232, 65, 107), 131);
            palette.Add(Color.FromArgb(255, 183, 56, 143), 132);
            palette.Add(Color.FromArgb(255, 88, 31, 126), 133);
            palette.Add(Color.FromArgb(255, 140, 163, 212), 134);
            palette.Add(Color.FromArgb(255, 154, 154, 204), 135);
            palette.Add(Color.FromArgb(255, 89, 129, 193), 136);
            palette.Add(Color.FromArgb(255, 65, 102, 176), 137);
            palette.Add(Color.FromArgb(255, 71, 95, 171), 138);
            palette.Add(Color.FromArgb(255, 55, 69, 147), 139);
            palette.Add(Color.FromArgb(255, 61, 86, 165), 140);
            palette.Add(Color.FromArgb(255, 41, 66, 135), 141);
            palette.Add(Color.FromArgb(255, 37, 38, 138), 142);
            palette.Add(Color.FromArgb(255, 26, 47, 111), 143);
            palette.Add(Color.FromArgb(255, 211, 201, 93), 144);
            palette.Add(Color.FromArgb(255, 81, 9, 24), 145);
            palette.Add(Color.FromArgb(255, 100, 179, 158), 146);
            palette.Add(Color.FromArgb(255, 99, 67, 56), 147);
            palette.Add(Color.FromArgb(255, 237, 211, 158), 148);
            palette.Add(Color.FromArgb(255, 105, 99, 171), 149);
            palette.Add(Color.FromArgb(255, 43, 63, 31), 150);
            palette.Add(Color.FromArgb(255, 151, 145, 197), 151);
            palette.Add(Color.FromArgb(255, 184, 189, 224), 152);
            palette.Add(Color.FromArgb(255, 249, 200, 152), 153);
            palette.Add(Color.FromArgb(255, 195, 144, 105), 154);
            palette.Add(Color.FromArgb(255, 68, 80, 91), 155);
            palette.Add(Color.FromArgb(255, 62, 73, 85), 156);
            palette.Add(Color.FromArgb(255, 32, 40, 48), 157);
            #endregion
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
                Console.WriteLine("palette 3");
            else
                Console.WriteLine("Unselect palette 3");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                Console.WriteLine("palette 1");
            else
                Console.WriteLine("Unselect palette 1");
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
                Console.WriteLine("palette 2");
            else
                Console.WriteLine("Unselect palette 2");
        }
    }
}
