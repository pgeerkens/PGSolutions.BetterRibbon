using System;
using System.Windows.Forms;

namespace Debug_Form
{
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            var s = "friend";
            var test = new PGSolutions.RibbonDispatcher.Main();
            var t = "enemy";
        }
    }
}
