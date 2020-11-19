using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PhotoTable
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string user = "1111";
            string admin = "allyourbasearebelongtous";
            if (textBox11.Text == user)
            {
                Form3 f = new Form3();
                f.Show();
            }
            else if (textBox11.Text == admin)
            {
                Form4 f = new Form4();
                f.Show();
            }
            else MessageBox.Show("Некоректний пароль");
        }
    }
}
