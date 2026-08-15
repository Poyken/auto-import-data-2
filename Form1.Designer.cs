namespace ImportData
{
    /// <summary>
    /// File thiết kế giao diện (Tự động tạo bởi Visual Studio).
    /// Chứa các thiết lập về kích thước, màu sắc và thuộc tính của các Control trên Form.
    /// </summary>
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Giải phóng bộ nhớ khi Form bị đóng.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Khởi tạo các thành phần giao diện.
        /// Cấu hình Layout, Font chữ và màu sắc Matrix (Đen - Xanh lá).
        /// </summary>
        private void InitializeComponent()
        {
            // Khởi tạo các đối tượng điều khiển (Controls) trên giao diện.
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnChangeFolder = new System.Windows.Forms.ToolStripButton();
            this.btnSyncHistory = new System.Windows.Forms.ToolStripButton();
            this.lblTeamName = new System.Windows.Forms.ToolStripLabel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip(); 
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel(); 
            this.lstLogs = new System.Windows.Forms.ListBox(); 
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();

            // --- THIẾT LẬP THANH CÔNG CỤ (toolStrip1) ---
            // Đặt ở phía trên cùng của cửa sổ, chứa nút "Đổi thư mục" và tên đội nhóm.
            this.toolStrip1.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.btnChangeFolder,
                this.btnSyncHistory,
                this.lblTeamName
            });
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(784, 30);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Renderer = new DarkToolStripRenderer();

            // --- NÚT ĐỔI THƯ MỤC (btnChangeFolder) ---
            // Nút bấm cho phép người dùng chọn thư mục quét máy đo mới.
            this.btnChangeFolder.ForeColor = System.Drawing.Color.White;
            this.btnChangeFolder.Image = null;
            this.btnChangeFolder.Name = "btnChangeFolder";
            this.btnChangeFolder.Size = new System.Drawing.Size(90, 22);
            this.btnChangeFolder.Text = "📁 Đổi thư mục";
            this.btnChangeFolder.Click += new System.EventHandler(this.BtnChangeFolder_Click);

            // --- NÚT ĐỒNG BỘ LỊCH SỬ (btnSyncHistory) ---
            // Nút bấm kích hoạt tiến trình quét ngầm toàn bộ dữ liệu cũ.
            this.btnSyncHistory.ForeColor = System.Drawing.Color.White;
            this.btnSyncHistory.Image = null;
            this.btnSyncHistory.Name = "btnSyncHistory";
            this.btnSyncHistory.Size = new System.Drawing.Size(120, 22);
            this.btnSyncHistory.Text = "🔄 Đồng bộ lịch sử";
            this.btnSyncHistory.Click += new System.EventHandler(this.BtnSyncHistory_Click);

            // --- NHÃN TÊN ĐỘI (lblTeamName) ---
            // Hiển thị tên đội phát triển bên phải thanh công cụ.
            this.lblTeamName.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.lblTeamName.ForeColor = System.Drawing.Color.FromArgb(120, 120, 120);
            this.lblTeamName.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblTeamName.Name = "lblTeamName";
            this.lblTeamName.Size = new System.Drawing.Size(140, 22);
            this.lblTeamName.Text = "Vietnam Develop EA Team";

            // --- THIẾT LẬP THANH TRẠNG THÁI (statusStrip1) ---
            // Đặt màu nền tối (Gần đen) cho thanh trạng thái nằm dưới chân ứng dụng.
            this.statusStrip1.BackColor = System.Drawing.Color.FromArgb(30, 30, 30); 
            
            // Gắn nhãn chữ (lblStatus) vào thanh trạng thái.
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus});
            
            // Vị trí đặt thanh trạng thái (Tọa độ X=0, Y=419).
            this.statusStrip1.Location = new System.Drawing.Point(0, 419); 
            this.statusStrip1.Name = "statusStrip1";
            
            // Kích thước thanh ngang (Rộng=784, Cao=22).
            this.statusStrip1.Size = new System.Drawing.Size(784, 22); 
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            
            // --- THIẾT LẬP NHÃN CHỮ TRẠNG THÁI (lblStatus) ---
            // Đặt màu chữ là trắng mặc định.
            this.lblStatus.ForeColor = System.Drawing.Color.White; 
            this.lblStatus.Name = "lblStatus";
            
            // Kích thước và nội dung chữ hiển thị ban đầu.
            this.lblStatus.Size = new System.Drawing.Size(115, 17);
            this.lblStatus.Text = "Đang chờ thư mục: " + DateTime.Now.ToString("yyyy-MM-dd");
            this.lblStatus.Margin = new System.Windows.Forms.Padding(5, 3, 0, 2);
            
            // --- THIẾT LẬP BẢNG NHẬT KÝ (lstLogs) ---
            // Đặt màu nền đen sâu thẳm giống màn hình Hacker.
            this.lstLogs.BackColor = System.Drawing.Color.Black; 
            
            // Loại bỏ đường viền của bảng để giao diện trông phẳng và hiện đại hơn.
            this.lstLogs.BorderStyle = System.Windows.Forms.BorderStyle.None; 
            
            // Dock=Fill: Yêu cầu bảng này tự động "Uống chiếm" toàn bộ không gian còn trống của cửa sổ.
            this.lstLogs.Dock = System.Windows.Forms.DockStyle.Fill; 
            
            // Đặt phông chữ Consolas (Chữ chuyên dụng cho code) cỡ 10.
            this.lstLogs.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            
            // Màu chữ là Xanh Lá (Lime) - Màu kinh điển của các thiết bị công nghiệp.
            this.lstLogs.ForeColor = System.Drawing.Color.Lime; 
            this.lstLogs.FormattingEnabled = true;
            
            // Cho phép xuất hiện thanh cuộn ngang nếu dòng log quá dài.
            this.lstLogs.HorizontalScrollbar = true; 
            this.lstLogs.ItemHeight = 22;
            this.lstLogs.Location = new System.Drawing.Point(0, 25);
            this.lstLogs.Name = "lstLogs";
            this.lstLogs.Size = new System.Drawing.Size(784, 394);
            this.lstLogs.TabIndex = 1;

            // --- THIẾT LẬP CỬA SỔ CHÍNH (Form1) ---
            // Cấu hình tỷ lệ hiển thị phông chữ của Windows. (7x15 pixels).
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            
            this.BackColor = System.Drawing.Color.Black; // Nền cửa sổ màu đen.
            
            // Kích thước ban đầu của cửa sổ (Rộng=784, Cao=441).
            this.ClientSize = new System.Drawing.Size(784, 441); 
            
            // Gắn các điều khiển đã tạo ở trên vào khung cửa sổ.
            this.Controls.Add(this.lstLogs);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.statusStrip1);
            this.Name = "Form1";
            
            // CenterScreen: Yêu cầu khi bật App lên thì nó tự nhảy vào chính giữa màn hình máy tính.
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vietnam Develop EA Team";
            
            // Tiếp tục vẽ lại giao diện sau khi đã thiết lập xong xuôi.
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        // Khai báo các biến đại diện cho các thành phần trên giao diện.
        private System.Windows.Forms.ToolStrip toolStrip1;               // Thanh công cụ phía trên.
        private System.Windows.Forms.ToolStripButton btnChangeFolder;    // Nút đổi thư mục quét.
        private System.Windows.Forms.ToolStripButton btnSyncHistory;     // Nút đồng bộ lịch sử.
        private System.Windows.Forms.ToolStripLabel lblTeamName;         // Nhãn tên đội phát triển (bên phải).
        private System.Windows.Forms.StatusStrip statusStrip1;           // Thanh ngang chứa trạng thái ở dưới.
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;     // Dòng chữ hiển thị kết quả (OK/Lỗi).
        private System.Windows.Forms.ListBox lstLogs;                    // Bảng hiển thị danh sách nhật ký nạp file.
    }

    /// <summary>
    /// Lớp tùy chỉnh Renderer cho ToolStrip: Giúp thanh công cụ có nền tối đồng bộ với giao diện Hacker.
    /// Xóa bỏ đường viền mặc định xấu xí của Windows.
    /// </summary>
    public class DarkToolStripRenderer : System.Windows.Forms.ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBorder(System.Windows.Forms.ToolStripRenderEventArgs e)
        {
            // Không vẽ đường viền → Thanh công cụ nằm phẳng liền mạch với giao diện.
        }

        protected override void OnRenderButtonBackground(System.Windows.Forms.ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected || e.Item.Pressed)
            {
                // Vẽ nền khi di chuột qua hoặc nhấn nút (Màu xanh đen nhẹ).
                using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(60, 60, 70)))
                {
                    e.Graphics.FillRectangle(brush, 0, 0, e.Item.Width, e.Item.Height);
                }
            }
        }
    }
}
