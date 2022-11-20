using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace fordito
{
    public partial class Form1 : Form
    {

        string path;
        string[,] data;
        string input, stack, rules;
        TreeNode tree;
        List<string> list;
        int maxHeight = 0;
        Graphics g;
        const string START_SIGN = "E#";
        bool draw = false;

        public Form1()
        {
            InitializeComponent();
            listBox1.Items.Clear();
        }

        private string converter(string input)
        {
            return Regex.Replace(input, "[0-9]+", "i") + "#";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox2.Text = converter(textBox1.Text);
            input = textBox2.Text;
            stack = START_SIGN;
            rules = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            input = textBox2.Text;
            stack = START_SIGN;
            rules = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            path = @textBox3.Text;
            Cursor.Current = Cursors.WaitCursor;
            if (path.Length == 0)
                return;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rows = xlRange.Rows.Count;
            int columns = xlRange.Columns.Count;

            data = new string[columns, rows];

            this.dataGridView1.ColumnCount = columns;
            this.dataGridView1.RowCount = rows;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    if (xlRange.Cells[i, j].Value != null)
                    {
                        this.dataGridView1[j - 1, i - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                        data[j - 1, i - 1] = xlRange.Cells[i, j].Value2.ToString();
                    }
                    else
                    {
                        this.dataGridView1[j - 1, i - 1].Value = "";
                        data[j - 1, i - 1] = "";
                    }
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Cursor.Current = Cursors.Default;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            input = textBox2.Text;
            stack = "E#";
            rules = "";
            list = new List<string>();
            tree = new TreeNode(stack[0].ToString(), null);
            solver();
            for (int i = 0; i < list.Count; i++)
            {
                listBox1.Items.Add(list[i]);
            }
            buildTree(tree, 1);
            listBox1.Items.Add(maxHeight.ToString());
            draw = true;
            pictureBox1.Invalidate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            path = @textBox3.Text;
            if (path.Length == 0)
                return;
            Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            for (int i = 1; i <= this.dataGridView1.RowCount; i++)
            {
                for (int j = 1; j <= this.dataGridView1.ColumnCount; j++)
                {
                    if (this.dataGridView1[j - 1, i - 1].Value != null)
                    {
                        xlWorksheet.Cells[i, j] = this.dataGridView1[j - 1, i - 1].Value;
                    }
                    else
                    {
                        xlWorksheet.Cells[i, j] = "";
                    }
                }
            }
            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Save();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Cursor.Current = Cursors.Default;
        }

        private void solver()
        {
            while (true)
            {
                if (dataGridView1.RowCount < 2 || dataGridView1.ColumnCount < 2)
                {
                    MessageBox.Show("Üres táblázat!");
                    break;
                }

                if (input == "#" && stack == "#")
                {
                    listBox1.Items.Add(styleConvert());
                    break;
                }

                int x = 0;
                string temp;

                if (stack[1].ToString() == "'")
                {
                    temp = stack[0].ToString() + stack[1].ToString();
                }
                else
                {
                    temp = stack[0].ToString();
                }

                while (input[0].ToString() != data[x, 0])
                {
                    x++;
                }

                int y = 0;
                while (temp.ToString() != data[0, y])
                {
                    y++;
                }

                if (data[x, y] == "")
                {
                    MessageBox.Show("Hiba!" + x.ToString() + y.ToString());
                    break;
                }

                calculateStep(x, y);

                listBox1.Items.Add(styleConvert());


            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private string styleConvert()
        {
            return "(  " + input + ",  " + stack + ",  " + rules + "  )";
        }

        private void calculateStep(int x, int y)
        {
            string[] step = Regex.Replace(data[x, y], @"[()\s]", "").Split(',');
            if (stack[1].ToString() == "'")
            {
                stack = stack.Substring(2);
            }
            else
            {
                stack = stack.Substring(1);
            }
            if (step[0] == "e")
            {
                //stack = step[0] + stack;  //ha "e" epszilon, akkor ez a sor nem kell
                rules = rules + step[1];
            }
            else if (step[0].ToLower() == "pop")
            {
                input = input.Substring(1);
            }
            else
            {
                stack = step[0] + stack;
                rules = rules + step[1];
            }
            list.Add(step[0]);
        }

        private void buildTree(TreeNode node, int height)
        {
            if (list.Count == 0)
                return;

            if (height > maxHeight)
            {
                maxHeight = height;
            }
            string value, word = list[0];
            list.RemoveAt(0);
            if (word.ToLower() == "pop")
            {
                return;
            }

            for (int i = 0; i < word.Length; i++)
            {
                if (i < word.Length - 1 && word[i + 1].ToString() == "'")
                {
                    value = word[i].ToString() + word[i + 1].ToString();
                    i++;
                }
                else
                {
                    value = word[i].ToString();
                }

                if (node.left == null)
                {
                    node.left = new TreeNode(value, node);
                    buildTree(node.left, height + 1);
                }
                else if (node.middle == null)
                {
                    node.middle = new TreeNode(value, node);
                    buildTree(node.middle, height + 1);
                }
                else if (node.right == null)
                {
                    node.right = new TreeNode(value, node);
                    buildTree(node.right, height + 1);
                }
            }

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            if (draw)
            {
                drawGraph(g, tree, 1, 1, 1, new PointF(0, 0));
                MessageBox.Show("rajz");
                draw = false;
            }
        }

        private void drawGraph(Graphics g, TreeNode node, int height, double width, double number, PointF point)
        {
            float incY = pictureBox1.Height / (maxHeight + 2);
            float y = incY * height;

            double incX = (pictureBox1.Width) / (width + 1);
            float x = (float)(incX * number);

            if (point.X != 0 && point.Y != 0)
            {
                g.DrawLine(new Pen(Color.Black), point, new PointF(x, y));
                g.DrawEllipse(new Pen(Color.Black), point.X - 2, point.Y - 2, 12, 12);
                g.FillEllipse(new SolidBrush(Color.White), point.X - 2, point.Y - 2, 12, 12);
                g.DrawString(node.parent.value, new Font("Arial", 6), Brushes.Black, point);
            }

            g.DrawEllipse(new Pen(Color.Black), x - 2, y - 2, 12, 12);
            g.DrawString(node.value, new Font("Arial", 6), Brushes.Black, new PointF(x, y));




            if (node.left != null)
            {
                drawGraph(g, node.left, height + 1, width * 3, 3 * number - 2, new PointF(x, y));
            }
            if (node.middle != null)
            {
                drawGraph(g, node.middle, height + 1, width * 3, 3 * number - 1, new PointF(x, y));
            }
            if (node.right != null)
            {
                drawGraph(g, node.right, height + 1, width * 3, 3 * number, new PointF(x, y));
            }
        }



        internal class TreeNode
        {
            public string value;
            public TreeNode left, middle, right, parent;

            public TreeNode(string value, TreeNode parent)
            {
                this.value = value;
                this.parent = parent;
            }
        }

    }
}
