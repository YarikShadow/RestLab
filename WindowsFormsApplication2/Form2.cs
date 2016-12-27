using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApplication2
{
    public partial class Form2 : Form
    {
        

        public Form2()
        {
            InitializeComponent();
            Change();

         }

        
        public bool language;
        private void Change()
        {
            this.language = Form1.language;
            if (language == true) { 
            label2.Text = "Версія 1.1.0";
            label3.Text = "Автоматийзуй(с) Генрі Форд";
            label1.Text = "       Програма разроблена студентом програмної инженерії\n     для студентів та робітників харчової промисловості";
            
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
