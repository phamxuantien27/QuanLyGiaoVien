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
using Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class DoiMatKhau : Form
    {
        Bus_DangNhap dangNhap = new Bus_DangNhap();
        List<string> l = new List<string>();
        private string tenTK;

        public string TenTK { get => tenTK; set => tenTK = value; }

        public DoiMatKhau()
        {
            InitializeComponent();
        }
        
        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void DoiMatKhau_Load(object sender, EventArgs e)
        {
            txtTenTk.Text = tenTK;
            
        }

        private void btnCapNhat_Click(object sender, EventArgs e)
        {
            
            if(txtMKcu.Text!="" && txtMKmoi.Text!="" && txtNhapLaiMK.Text!="")
            {
                if(txtMKmoi.Text == txtNhapLaiMK.Text)
                {
                    if (dangNhap.doiMatKhau(txtTenTk.Text, txtMKmoi.Text))
                    {
                        MessageBox.Show("ĐỔi thành công");
                        this.Close();
                        DangNhap f = new DangNhap();
                        f.ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show("Mật khẩu mới chưa chính xác!");
                }
            }
            else
            {
                MessageBox.Show("Nhập đầy đủ các thông tin!");
            }

            
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            FormMain f = new FormMain();
            this.Close();
            f.ShowDialog();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
