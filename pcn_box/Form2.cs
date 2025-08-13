using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pcn_box
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
            Program.fast_str = listBox1.Text;
            Close();

        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            if (Program.comments == 1)
            {
                if (File.Exists("comments1.txt"))
                {
                    listBox1.Items.AddRange((object[])File.ReadAllLines("comments1.txt", Encoding.GetEncoding(1251)));
                    
                }

            }
            if (Program.comments == 2)
            {
                if (File.Exists("comments2.txt"))
                {
                    listBox1.Items.AddRange((object[])File.ReadAllLines("comments2.txt", Encoding.GetEncoding(1251)));

                }


            }

        }
    }
}
