using System;
using System.Drawing;
using System.Windows.Forms;

namespace ImportData
{
    public enum SyncMode
    {
        LastNDays,
        DateRange,
        All
    }

    public class SyncOptions
    {
        public SyncMode Mode { get; set; } = SyncMode.LastNDays;
        public int Days { get; set; } = 7;
        public DateTime StartDate { get; set; } = DateTime.Today.AddDays(-7);
        public DateTime EndDate { get; set; } = DateTime.Today;
        public bool OverwriteExisting { get; set; } = false;
    }

    public class SyncHistoryDialog : Form
    {
        public SyncOptions Options { get; private set; } = new SyncOptions();

        private RadioButton radRecent = null!;
        private RadioButton radRange = null!;
        private RadioButton radAll = null!;
        private NumericUpDown numDays = null!;
        private DateTimePicker dtpFrom = null!;
        private DateTimePicker dtpTo = null!;
        private Button btnOk = null!;
        private Button btnCancel = null!;
        private Label lblWarning = null!;
        private CheckBox chkOverwrite = null!;

        public SyncHistoryDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Cấu hình đồng bộ lịch sử";
            this.Size = new Size(460, 390);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(25, 25, 25);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);

            // Panel tiêu đề
            var pnlTitle = new Panel
            {
                Dock = DockStyle.Top,
                Height = 45,
                BackColor = Color.FromArgb(35, 35, 35)
            };
            var lblTitle = new Label
            {
                Text = "⚡ CẤU HÌNH PHẠM VI ĐỒNG BỘ LỊCH SỬ",
                Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
                ForeColor = Color.Lime,
                Location = new Point(15, 12),
                AutoSize = true
            };
            pnlTitle.Controls.Add(lblTitle);
            this.Controls.Add(pnlTitle);

            // Container Panel cho nội dung
            var pnlContent = new Panel
            {
                Location = new Point(0, 45),
                Size = new Size(460, 270),
                BackColor = Color.FromArgb(20, 20, 20)
            };
            this.Controls.Add(pnlContent);

            // --- Option 1: Số ngày gần đây ---
            radRecent = new RadioButton
            {
                Text = "Đồng bộ theo số ngày gần đây",
                Location = new Point(25, 25),
                Size = new Size(220, 24),
                Checked = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.Lime
            };
            radRecent.CheckedChanged += RadMode_CheckedChanged;
            pnlContent.Controls.Add(radRecent);

            numDays = new NumericUpDown
            {
                Location = new Point(260, 25),
                Size = new Size(60, 24),
                Minimum = 1,
                Maximum = 365,
                Value = 7,
                BackColor = Color.Black,
                ForeColor = Color.Lime,
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlContent.Controls.Add(numDays);

            var lblDays = new Label
            {
                Text = "ngày",
                Location = new Point(325, 27),
                AutoSize = true,
                ForeColor = Color.White
            };
            pnlContent.Controls.Add(lblDays);

            // --- Option 2: Khoảng thời gian ---
            radRange = new RadioButton
            {
                Text = "Đồng bộ theo khoảng thời gian",
                Location = new Point(25, 75),
                Size = new Size(250, 24),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.Lime
            };
            radRange.CheckedChanged += RadMode_CheckedChanged;
            pnlContent.Controls.Add(radRange);

            var lblFrom = new Label
            {
                Text = "Từ ngày:",
                Location = new Point(50, 112),
                AutoSize = true,
                ForeColor = Color.DarkGray
            };
            pnlContent.Controls.Add(lblFrom);

            dtpFrom = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Location = new Point(120, 110),
                Size = new Size(110, 24),
                BackColor = Color.Black,
                ForeColor = Color.White,
                Enabled = false
            };
            pnlContent.Controls.Add(dtpFrom);

            var lblTo = new Label
            {
                Text = "Đến ngày:",
                Location = new Point(245, 112),
                AutoSize = true,
                ForeColor = Color.DarkGray
            };
            pnlContent.Controls.Add(lblTo);

            dtpTo = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Location = new Point(315, 110),
                Size = new Size(110, 24),
                BackColor = Color.Black,
                ForeColor = Color.White,
                Enabled = false
            };
            pnlContent.Controls.Add(dtpTo);

            // --- Option 3: Tất cả lịch sử ---
            radAll = new RadioButton
            {
                Text = "Đồng bộ toàn bộ lịch sử (Quét toàn bộ thư mục)",
                Location = new Point(25, 160),
                Size = new Size(380, 24),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.Lime
            };
            radAll.CheckedChanged += RadMode_CheckedChanged;
            pnlContent.Controls.Add(radAll);

            chkOverwrite = new CheckBox
            {
                Text = "Ghi đè/Nạp lại dữ liệu của các tệp đã tồn tại",
                Location = new Point(25, 195),
                Size = new Size(400, 24),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.Yellow,
                Checked = false
            };
            pnlContent.Controls.Add(chkOverwrite);

            lblWarning = new Label
            {
                Text = "⚠️ Cảnh báo: Quét toàn bộ lịch sử có thể mất nhiều thời gian\nnếu thư mục gốc chứa hàng ngàn file Excel.",
                Location = new Point(50, 225),
                Size = new Size(380, 40),
                ForeColor = Color.FromArgb(255, 200, 100),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                Visible = false
            };
            pnlContent.Controls.Add(lblWarning);

            // --- Footer Panel: Chứa nút bấm ---
            var pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            this.Controls.Add(pnlFooter);

            btnOk = new Button
            {
                Text = "▶ Bắt đầu",
                Location = new Point(220, 15),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Black,
                ForeColor = Color.Lime,
                Cursor = Cursors.Hand
            };
            btnOk.FlatAppearance.BorderColor = Color.Lime;
            btnOk.FlatAppearance.BorderSize = 1;
            btnOk.Click += BtnOk_Click;
            btnOk.MouseEnter += (s, e) => btnOk.BackColor = Color.FromArgb(0, 50, 0);
            btnOk.MouseLeave += (s, e) => btnOk.BackColor = Color.Black;
            pnlFooter.Controls.Add(btnOk);

            btnCancel = new Button
            {
                Text = "✕ Hủy bỏ",
                Location = new Point(330, 15),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Black,
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderColor = Color.DarkGray;
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.Click += BtnCancel_Click;
            btnCancel.MouseEnter += (s, e) => btnCancel.BackColor = Color.FromArgb(40, 40, 40);
            btnCancel.MouseLeave += (s, e) => btnCancel.BackColor = Color.Black;
            pnlFooter.Controls.Add(btnCancel);
        }

        private void RadMode_CheckedChanged(object sender, EventArgs e)
        {
            numDays.Enabled = radRecent.Checked;
            dtpFrom.Enabled = radRange.Checked;
            dtpTo.Enabled = radRange.Checked;
            lblWarning.Visible = radAll.Checked;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            Options.OverwriteExisting = chkOverwrite.Checked;

            if (radRecent.Checked)
            {
                Options.Mode = SyncMode.LastNDays;
                Options.Days = (int)numDays.Value;
                Options.StartDate = DateTime.Today.AddDays(-Options.Days);
                Options.EndDate = DateTime.Today;
            }
            else if (radRange.Checked)
            {
                if (dtpFrom.Value.Date > dtpTo.Value.Date)
                {
                    MessageBox.Show("Ngày bắt đầu không được lớn hơn ngày kết thúc!", "Lỗi cấu hình", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                Options.Mode = SyncMode.DateRange;
                Options.StartDate = dtpFrom.Value;
                Options.EndDate = dtpTo.Value;
            }
            else
            {
                Options.Mode = SyncMode.All;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
