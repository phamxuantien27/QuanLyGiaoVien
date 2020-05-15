using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DTO;
using QLGV;
namespace WindowsFormsApp1
{
    public partial class DangNhap : Form
    {
        Bus_DangNhap dangNhap = new Bus_DangNhap();
        
        public DangNhap()
        {

            InitializeComponent();
            txtMatKhau.PasswordChar = '*';
            txtMatKhau.MaxLength = 20;
        }

        private void DangNhap_Load(object sender, EventArgs e)
        {
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(dangNhap.DangNhap(txtTaiKhoan.Text,txtMatKhau.Text))
                {
                    FormMain f = new FormMain();
                    this.Hide();
                f.TenTaiKhoan = txtTaiKhoan.Text;
                f.ShowDialog();
                   
                    this.Close();
                    
                
            }
            else
            {
                MessageBox.Show("Tên tài khoản hoặc mật khẩu không đúng");
            }

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtTaiKhoan_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtMatKhau_TextChanged(object sender, EventArgs e)
        {

        }

        //private void txtTaiKhoan_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if(e.KeyChar>= (char) 48&&e.KeyChar<=(char)57|| e.KeyChar >= (char)65 && e.KeyChar <= (char)90|| e.KeyChar >= (char)97 && e.KeyChar <= (char)122)
        //    {
        //        e.Handled = false;
        //    }
        //    else
        //    {
        //        txtTaiKhoan.Text = "";
        //        MessageBox.Show("Tài khoản không chứa các ký tự khác chữ và số");
        //    }
        //}
    }
}
