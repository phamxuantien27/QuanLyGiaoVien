using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QLGV;
using DTO;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class FormQuanLy : Form
    {
        Bus_GiaoVien giaoVien = new Bus_GiaoVien();
        string s;
        OpenFileDialog op = new OpenFileDialog();
        private int loaiTaiKhoan;

        public int LoaiTaiKhoan { get => loaiTaiKhoan; set => loaiTaiKhoan = value; }

        public FormQuanLy()
        {
            InitializeComponent();
            loadDtgvGV2();
            loadMa_Khoa();
            
            loadHocHam();
            loadHocVi();
        }
        //void loadDtgvGV()
        //{
        //    dgvgiaovien.DataSource =giaoVien.getGiaoVien();
        //    dgvgiaovien.Columns[0].HeaderText = "Mã giáo viên";
        //    dgvgiaovien.Columns[1].HeaderText = "Họ tên";
        //    dgvgiaovien.Columns[2].HeaderText = "Ngày sinh";
        //    dgvgiaovien.Columns[3].HeaderText = "Giới tính";
        //    dgvgiaovien.Columns[4].HeaderText = "Email";
        //    dgvgiaovien.Columns[5].HeaderText = "Học vị";
        //    dgvgiaovien.Columns[6].HeaderText = "Học hàm";
        //    dgvgiaovien.Columns[7].HeaderText = "Tên môn giảng dạy";
        //    dgvgiaovien.Columns[8].HeaderText = "Chức vụ";
        //    dgvgiaovien.Columns[9].HeaderText = "Bộ môn";
        //    dgvgiaovien.Columns[10].HeaderText = "Khoa";
            

        //}
        void loadDtgvGV2()
        {
            dgvgiaovien.DataSource = giaoVien.getGiaoVien2();
            dgvgiaovien.Columns[0].HeaderText = "Mã giáo viên";
            dgvgiaovien.Columns[1].HeaderText = "Họ tên";
            dgvgiaovien.Columns[2].HeaderText = "Ngày sinh";
            dgvgiaovien.Columns[3].HeaderText = "Giới tính";
            dgvgiaovien.Columns[4].HeaderText = "Email";
            dgvgiaovien.Columns[5].HeaderText = "Mã Học vị";
            dgvgiaovien.Columns[6].HeaderText = "Mã Học Hàm";
            dgvgiaovien.Columns[7].HeaderText = "Mã Khoa";
            dgvgiaovien.Columns[8].HeaderText = "Mã Bộ Môn";
            dgvgiaovien.Columns[9].HeaderText = "Mã Môn Học";
           
            dgvgiaovien.Columns[11].HeaderText = "Dịa Chỉ";
            dgvgiaovien.Columns[10].HeaderText = "Chức Vụ";


        }
        void loadMa_Khoa()
        {
            cmbkhoa.DataSource = giaoVien.getMa_Khoa();
            cmbkhoa.DisplayMember = "Ten_Khoa";
            cmbkhoa.ValueMember = "Ma_Khoa";
        }
        void loadHocHam()
        {
            cmbHocHam.DataSource = giaoVien.getHocHam();
            cmbHocHam.DisplayMember = "Ma_HocHam";
            cmbHocHam.ValueMember = "";
        }
        void loadHocVi()
        {
            cmbHocVi.DataSource = giaoVien.getHocVi();
            cmbHocVi.DisplayMember = "Ma_HocVi";
            cmbHocVi.ValueMember = "";
        }

        public List<DTO_GiaoVien> loadGiaoVien()
        {
            return giaoVien.load_GV();
        }
        bool checkMaGV()
        {
            bool kq = true;
            for(int i = 0;i< loadGiaoVien().Count;i++)
            {
                if (txtmagv.Text == loadGiaoVien()[i].Ma_GiaoVien)
                {
                    kq = false;
                }
                   
             
            }
            return kq;
        }
        void loadBoMon(string s)
        {
            cmbbomon.DataSource = giaoVien.getMa_BoMon(s);
            cmbbomon.DisplayMember = "Ma_BoMon";
            cmbbomon.ValueMember = "";
        }
        void getMonHoc(string s)
        {
            cmbMaMonHoc.DataSource = giaoVien.getMonHoc(s);
            cmbMaMonHoc.DisplayMember = "Ma_MonHoc";
            cmbMaMonHoc.ValueMember = "";
        }
        private void FormQuanLy_Load(object sender, EventArgs e)
        {
            btnThem.Visible = false;
            btnSua.Visible = false;
            btnXoa.Visible = false;
            if (loaiTaiKhoan==0)
            {
                btnThem.Visible = true;
                btnSua.Visible = true;
                btnXoa.Visible = true;
            }

            cmbLuaChonTimKiem.Items.Add("Theo mã giáo viên");
            cmbLuaChonTimKiem.Items.Add("Theo họ tên giáo viên");
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnFresh_Click(object sender, EventArgs e)
        {
            btnThem.Visible = true;
            
            btnXoa.Visible = true;
        }

        private void dgvgiaovien_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            txtmagv.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[0].Value);
            txthoten.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[1].Value);
            dateNS.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[2].Value);
            txtGioiTin.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[3].Value);
            txtEmail.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[4].Value);
            cmbHocVi.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[5].Value);
            cmbHocHam.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[6].Value);
            cmbkhoa.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[7].Value);
            cmbbomon.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[8].Value);
            cmbMaMonHoc.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[9].Value);
            txtdiachi.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[11].Value);
            txtchucvu.Text = Convert.ToString(dgvgiaovien.CurrentRow.Cells[10].Value);
            

            ThongTinGiaoVien f = new ThongTinGiaoVien();
                f.MaGiaVien = Convert.ToString(dgvgiaovien.CurrentRow.Cells[0].Value);


            System.Data.DataTable dt = giaoVien.getPath(Convert.ToString(dgvgiaovien.CurrentRow.Cells[0].Value));
                if (dt.Rows[0]["paths"].ToString() == "")
                {
                     pbanhDaiDien.Image = Image.FromFile(@"D:\KI 5\C#\QLGV\WindowsFormsApp1\image\manh.jpg");
                      f.DuongDanAnh = @"D:\KI 5\C#\QLGV\WindowsFormsApp1\image\manh.jpg";
                
                }
                else
                {
                    pbanhDaiDien.Image = Image.FromFile(dt.Rows[0]["paths"].ToString());
                     f.DuongDanAnh = dt.Rows[0]["paths"].ToString();

                 }
            f.ShowDialog();





        }
   
        private void cmbkhoa_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadBoMon(this.cmbkhoa.Text);
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            DTO_GiaoVien gv = new DTO_GiaoVien();
            gv.Ma_GiaoVien = txtmagv.Text;
            gv.HoTen = txthoten.Text;
            gv.HocVi = cmbHocVi.Text;
            gv.HocHam = cmbHocHam.Text;
            gv.NgaySinh = DateTime.Parse(dateNS.Text);
            gv.Ma_BoMon = cmbbomon.Text;
            gv.Ma_Khoa = cmbkhoa.Text;
            gv.DiaChi1 = txtdiachi.Text;
            gv.Email = txtEmail.Text;
            gv.ChucVu = txtchucvu.Text;
            gv.GioiTinh = txtGioiTin.Text;
            gv.Ma_MonHoc1 = cmbMaMonHoc.Text;
            gv.DuongDanAnh = s;
            if (gv.Ma_GiaoVien != "" && gv.HoTen != "" && gv.HocVi != "" && gv.HocHam != "" && gv.Ma_BoMon != "" && gv.Ma_Khoa != "" && gv.Ma_MonHoc1 != "" && gv.ChucVu != "" && gv.DiaChi1 != "" && gv.Email != "" && gv.GioiTinh != "")
            {
                if(checkMaGV()==true)
                {
                    if (giaoVien.Them(gv) == true)
                    {
                        MessageBox.Show("thêm thành công", "thông báo");
                        loadDtgvGV2();

                    }
                    else
                    {
                        MessageBox.Show("thêm thất bại", "thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("thêm thất bại,mã giáo viên trùng ", "thông báo");
                }
            }
            else if (gv.Ma_GiaoVien == "")
            {
                MessageBox.Show("Thiếu mã giáo viên", "Thông báo");
                
            }
           else if(gv.HoTen=="") MessageBox.Show("Thiếu tên giáo viên", "Thông báo");
            else if (gv.ChucVu == "") MessageBox.Show("Thiếu chức vụ ", "Thông báo");
            else if (gv.GioiTinh == "") MessageBox.Show("Nhập thiếu giới tính giáo viên", "Thông báo");
            else if (gv.DiaChi1 == "") MessageBox.Show("Thiếu địa chỉ giáo viên", "Thông báo");
            else if (gv.Email == "") MessageBox.Show("Thiếu email giáo viên", "Thông báo");
            else if (gv.NgaySinh.ToString() == "") MessageBox.Show("Ngay sinh giáo viên trống", "Thông báo");

            else  MessageBox.Show("Thiếu thông tin giáo viên", "Thông báo");
        }
       
        private void btnChonAnh_Click(object sender, EventArgs e)
        {

            if (op.ShowDialog() == DialogResult.OK)
            {
                s = op.FileName;
                pbanhDaiDien.Image = Image.FromFile(s);
            }
           
               
            
        }
        
        private void btnThayAnh_Click(object sender, EventArgs e)
        {
            if(giaoVien.ThemAnh(s,txtmagv.Text)==true)
            {
                MessageBox.Show(" thành công");
            }
            else
            {
                MessageBox.Show("thay đổi thất bại");
            }
        }

        private void cmbbomon_SelectedIndexChanged(object sender, EventArgs e)
        {
            getMonHoc(cmbbomon.Text);
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
           
            if (txtmagv.Text != "")
            {
                if(giaoVien.xoa(txtmagv.Text))
                {
                    MessageBox.Show("Xoa thanh cong");
                }
                else
                {
                    MessageBox.Show("Xoa that bai");
                }
            }
            else
            {
                MessageBox.Show("Xoa that bai");
            }
            loadDtgvGV2();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DTO_GiaoVien gv = new DTO_GiaoVien();
            gv.Ma_GiaoVien = txtmagv.Text;
            gv.HoTen = txthoten.Text;
            gv.HocVi = cmbHocVi.Text;
            gv.HocHam = cmbHocHam.Text;
            gv.NgaySinh = DateTime.Parse(dateNS.Text);
            gv.Ma_BoMon = cmbbomon.Text;
            gv.Ma_Khoa = cmbkhoa.Text;
            gv.DiaChi1 = txtdiachi.Text;
            gv.Email = txtEmail.Text;
            gv.ChucVu = txtchucvu.Text;
            gv.GioiTinh = txtGioiTin.Text;
            gv.Ma_MonHoc1 = cmbMaMonHoc.Text;
            if(giaoVien.sua(gv))
            {
                MessageBox.Show("Thay đổi thông tin giáo viên thành công", "Thông báo");
            }
            else
            {
                MessageBox.Show("Thay đổi thông tin giáo viên That bai", "Thông báo");
            }
            loadDtgvGV2();
        }
        //hàm in ra excel
        public static void ExportFile(string Header, DataGridView dgv)
        {
            // Tạo đối tượng mở Explorer
            SaveFileDialog fsave = new SaveFileDialog();
            // Chỉ ra đuôi của tệp tin
            fsave.Filter = "(Tất cả các tệp)|*.*|(Các tệp excel)|*.xlsx";
            fsave.ShowDialog();

            if (fsave.FileName != "")
            {
                // Tạo Excel App
                Excel.Application app = new Excel.Application();
               
                Excel.Workbook wb = app.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet = null;
                try
                {
                    // Đọc dữ liệu
                    sheet = wb.ActiveSheet;
                    sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, dgv.ColumnCount]].Merge();
                    sheet.Cells[1, 1].Value = Header;
                    sheet.Cells[1, 1].Font.Name = "Times New Roman";
                    sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[1, 1].Font.Size = 20;
                    sheet.Cells[1, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    //Sinh tiêu đề
                    for (int i = 1, k = 1; i <= dgv.Columns.Count; i++)
                    {
                        if (dgv.Columns[i - 1].Visible == false) continue;
                        sheet.Cells[2, k] = dgv.Columns[i - 1].HeaderText;
                        sheet.Cells[2, k].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sheet.Cells[2, k].Font.Name = "Times New Roman";
                        sheet.Cells[2, k].Font.Bold = true;
                        sheet.Cells[2, k].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        k++;
                    }
                    // Sinh dữ liệu
                    for (int i = 1; i <= dgv.RowCount - 1; i++)
                    {
                        if (dgv.Columns[0].Visible == false) continue;
                        sheet.Cells[i + 2, 1] = dgv.Rows[i - 1].Cells[0].Value;
                        sheet.Cells[i + 2, 1].Font.Name = "Times New Roman";
                        sheet.Cells[i + 2, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        for (int j = 2, k = 2; j <= dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j - 1].Visible == false) continue;
                            sheet.Cells[i + 2, k] = dgv.Rows[i - 1].Cells[j - 1].Value;
                            sheet.Cells[i + 2, k].Font.Name = "Times New Roman";
                            sheet.Cells[i + 2, k].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            k++;
                        }
                    }
                    sheet.Columns.AutoFit();
                    wb.SaveAs(fsave.FileName);
                    MessageBox.Show("Ghi thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    app.Quit();
                    wb = null;
                }

            }
            else
            {
                MessageBox.Show("Bạn không chọn tệp tin nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnIn_Click(object sender, EventArgs e)
        {
            ExportFile("Danh Sách giáo viên", dgvgiaovien);
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(cmbLuaChonTimKiem.Text=="Theo mã giáo viên")
            {
                dgvgiaovien.DataSource = giaoVien.timtheoma(txtTimKiem.Text);
                dgvgiaovien.Columns[0].HeaderText = "Mã giáo viên";
                dgvgiaovien.Columns[1].HeaderText = "Họ tên";
                dgvgiaovien.Columns[2].HeaderText = "Ngày sinh";
                dgvgiaovien.Columns[3].HeaderText = "Giới tính";
                dgvgiaovien.Columns[4].HeaderText = "Email";
                dgvgiaovien.Columns[5].HeaderText = "Mã Học vị";
                dgvgiaovien.Columns[6].HeaderText = "Mã Học Hàm";
                dgvgiaovien.Columns[7].HeaderText = "Mã Khoa";
                dgvgiaovien.Columns[8].HeaderText = "Mã Bộ Môn";
                dgvgiaovien.Columns[9].HeaderText = "Mã Môn Học";
                dgvgiaovien.Columns[11].HeaderText = "Dịa Chỉ";
                dgvgiaovien.Columns[10].HeaderText = "Chức Vụ";
            }
            if(cmbLuaChonTimKiem.Text=="Theo họ tên giáo viên")
            {
                dgvgiaovien.DataSource = giaoVien.timtheoten(txtTimKiem.Text);
                dgvgiaovien.Columns[0].HeaderText = "Mã giáo viên";
                dgvgiaovien.Columns[1].HeaderText = "Họ tên";
                dgvgiaovien.Columns[2].HeaderText = "Ngày sinh";
                dgvgiaovien.Columns[3].HeaderText = "Giới tính";
                dgvgiaovien.Columns[4].HeaderText = "Email";
                dgvgiaovien.Columns[5].HeaderText = "Mã Học vị";
                dgvgiaovien.Columns[6].HeaderText = "Mã Học Hàm";
                dgvgiaovien.Columns[7].HeaderText = "Mã Khoa";
                dgvgiaovien.Columns[8].HeaderText = "Mã Bộ Môn";
                dgvgiaovien.Columns[9].HeaderText = "Mã Môn Học";
                dgvgiaovien.Columns[11].HeaderText = "Dịa Chỉ";
                dgvgiaovien.Columns[10].HeaderText = "Chức Vụ";
            }
            if (cmbLuaChonTimKiem.Text == "" && txtTimKiem.Text == "")
                loadDtgvGV2();
        }

        private void txtmagv_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }
    }
}
