using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Specialized;
using System.Configuration;

namespace YouTrade.Winform
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }
        

        private void Click_LogOut(object sender, EventArgs e)
        {
            DialogResult dlr = MessageBox.Show("Bạn có thật sự muốn đăng xuất ?", "Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlr == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void Click_Login(object sender, EventArgs e)
        {
          //  var connection = ConnectionFactory.GetConnection(ConfigurationManager.ConnectionStrings["Test"].ConnectionString, DataBaseProvider);
            if (tbuser.Text == "" || tbpass.Text == "")
            {
                this.tt.ForeColor = Color.Red;
                this.tt.Text = "Bạn chưa nhập tên hoặc mật khẩu! Vui lòng thử lại";
            }
            else
            {
                string username= ConfigurationManager.AppSettings["username"];
                string password = ConfigurationManager.AppSettings["password"];

                if (tbuser.Text.Trim() == username && tbpass.Text == password)
                {
                    //MessageBox.Show("Login success", "Thông Báo");
                    MainForm mForm = new MainForm();
                    this.Hide();
                    mForm.ShowDialog();
                    this.Close();
                }
                else
                {
                    this.tt.ForeColor = Color.Red;
                    this.tt.Text = "Tên đăng nhập / Mật khẩu bạn đã nhập không chính xác! Vui lòng thử lại";
                    this.tbuser.Clear();
                    this.tbuser.Focus();
                    this.tbpass.Clear();
                }
            }
        }
    }
}
