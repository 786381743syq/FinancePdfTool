using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Globalization;

namespace FinancePdfApp
{
    public class ImageItem
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public int OrigWidth { get; set; }
        public int OrigHeight { get; set; }
        public long FileSizeBytes { get; set; }
        public int Rotation { get; set; } // 0, 90, 180, 270

        public int CurrentWidth
        {
            get { return (Rotation % 180 == 0) ? OrigWidth : OrigHeight; }
        }

        public int CurrentHeight
        {
            get { return (Rotation % 180 == 0) ? OrigHeight : OrigWidth; }
        }

        public bool IsLandscape
        {
            get { return CurrentWidth > CurrentHeight; }
        }
    }

    public class MainForm : Form
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        private List<ImageItem> items = new List<ImageItem>();
        private ListView listView;
        private PictureBox previewBox;
        private Label lblPreviewInfo;
        private RadioButton rbGov;
        private RadioButton rbLossless;
        private ComboBox cmbPaperSize;
        private ComboBox cmbOrientation;
        private CheckBox chkKeepRatio;
        private CheckBox chkMargin;
        private TextBox txtOutputPath;
        private Button btnBrowseOutput;
        private Button btnGenerate;
        private ProgressBar progressBar;
        private Label lblStatus;
        private BackgroundWorker worker;

        public MainForm(string[] args)
        {
            InitializeComponent();
            if (args != null && args.Length > 0)
            {
                this.Shown += delegate {
                    HandleDroppedPaths(args);
                };
            }
        }

        private void InitializeComponent()
        {
            this.Text = "财务专用 PDF 合成与智能压缩助手";
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // 固定窗口大小，禁止随意缩放导致变形，保证所有元素100%完整可见
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.ClientSize = new Size(1020, 730);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;
            this.BackColor = Color.FromArgb(245, 247, 250);

            // ================== 1. 顶部工具栏 (固定高度 60，按钮高度 36，上下均有充裕内边距) ==================
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1020, 60),
                BackColor = Color.White
            };

            Button btnAddFiles = CreateButton("➕ 添加图片", 105, 36, Color.FromArgb(37, 99, 235), Color.White);
            btnAddFiles.Location = new Point(15, 12);
            btnAddFiles.Click += BtnAddFiles_Click;

            Button btnAddFolder = CreateButton("📁 添加文件夹", 115, 36, Color.FromArgb(71, 85, 105), Color.White);
            btnAddFolder.Location = new Point(130, 12);
            btnAddFolder.Click += BtnAddFolder_Click;

            Button btnMoveUp = CreateButton("⬆ 上移", 75, 36, Color.White, Color.FromArgb(51, 65, 85));
            btnMoveUp.Location = new Point(265, 12);
            btnMoveUp.Click += delegate { MoveItem(-1); };

            Button btnMoveDown = CreateButton("⬇ 下移", 75, 36, Color.White, Color.FromArgb(51, 65, 85));
            btnMoveDown.Location = new Point(350, 12);
            btnMoveDown.Click += delegate { MoveItem(1); };

            Button btnRotate = CreateButton("🔄 旋转90°", 100, 36, Color.White, Color.FromArgb(51, 65, 85));
            btnRotate.Location = new Point(435, 12);
            btnRotate.Click += BtnRotate_Click;

            Button btnRemove = CreateButton("❌ 移除选中", 100, 36, Color.White, Color.FromArgb(220, 38, 38));
            btnRemove.Location = new Point(555, 12);
            btnRemove.Click += BtnRemove_Click;

            Button btnClear = CreateButton("清空列表", 90, 36, Color.White, Color.FromArgb(100, 116, 139));
            btnClear.Location = new Point(665, 12);
            btnClear.Click += delegate {
                items.Clear();
                UpdateListView();
                ClearPreview();
            };

            topBar.Controls.AddRange(new Control[] {
                btnAddFiles, btnAddFolder,
                btnMoveUp, btnMoveDown, btnRotate,
                btnRemove, btnClear
            });

            // 分界线
            Panel dividerTop = new Panel
            {
                Location = new Point(0, 59),
                Size = new Size(1020, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            topBar.Controls.Add(dividerTop);

            // ================== 2. 中间内容区 (高度 455) ==================
            // 左侧：列表
            Panel pnlList = new Panel
            {
                Location = new Point(15, 70),
                Size = new Size(610, 435),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            listView.Columns.Add("页码", 55, HorizontalAlignment.Center);
            listView.Columns.Add("文件名", 230, HorizontalAlignment.Left);
            listView.Columns.Add("分辨率", 110, HorizontalAlignment.Center);
            listView.Columns.Add("版式", 65, HorizontalAlignment.Center);
            listView.Columns.Add("旋转", 60, HorizontalAlignment.Center);
            listView.Columns.Add("原始大小", 85, HorizontalAlignment.Right);
            listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
            pnlList.Controls.Add(listView);

            // 右侧：大图实时预览
            Panel pnlPreview = new Panel
            {
                Location = new Point(635, 70),
                Size = new Size(370, 435),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            lblPreviewInfo = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(71, 85, 105),
                Text = "点击左侧文件预览画面"
            };

            previewBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(241, 245, 249)
            };

            pnlPreview.Controls.Add(previewBox);
            pnlPreview.Controls.Add(lblPreviewInfo);

            // ================== 3. 底部配置与生成区 (固定高度 215) ==================
            Panel bottomPanel = new Panel
            {
                Location = new Point(0, 515),
                Size = new Size(1020, 215),
                BackColor = Color.White
            };

            Panel dividerBottom = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1020, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            bottomPanel.Controls.Add(dividerBottom);

            // 3.1 画质与压缩选项 (左半部)
            GroupBox grpMode = new GroupBox
            {
                Text = "画质与压缩选项",
                Location = new Point(15, 6),
                Size = new Size(485, 96),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            rbGov = new RadioButton
            {
                Text = "⭐ 政务申报模式 (压缩约70%，公章发票极清)",
                Location = new Point(15, 27),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            rbLossless = new RadioButton
            {
                Text = "💎 原画无损模式 (100% 原画直出，零重编码)",
                Location = new Point(15, 59),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            grpMode.Controls.Add(rbGov);
            grpMode.Controls.Add(rbLossless);
            bottomPanel.Controls.Add(grpMode);

            // 3.2 纸张与页面版式设置 (右半部)
            GroupBox grpPaper = new GroupBox
            {
                Text = "纸张大小与页面版式",
                Location = new Point(510, 6),
                Size = new Size(495, 96),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblPaper = new Label
            {
                Text = "纸张：",
                Location = new Point(12, 31),
                AutoSize = true,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            cmbPaperSize = new ComboBox
            {
                Location = new Point(66, 28),
                Size = new Size(150, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbPaperSize.Items.AddRange(new object[] {
                "标准 A4 (常用推荐)",
                "适合原图 (无白边)",
                "标准 A3 (宽幅大表)"
            });
            cmbPaperSize.SelectedIndex = 0;

            Label lblOrient = new Label
            {
                Text = "方向：",
                Location = new Point(234, 31),
                AutoSize = true,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            cmbOrientation = new ComboBox
            {
                Location = new Point(288, 28),
                Size = new Size(190, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbOrientation.Items.AddRange(new object[] {
                "统一竖版 (纵向居中)",
                "智能感应 (自动横竖)",
                "统一横版 (全部横向)"
            });
            cmbOrientation.SelectedIndex = 0;

            chkKeepRatio = new CheckBox
            {
                Text = "保持比例居中（防拉伸）",
                Location = new Point(15, 64),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            chkMargin = new CheckBox
            {
                Text = "预留打印边距（防贴边）",
                Location = new Point(245, 64),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            cmbPaperSize.SelectedIndexChanged += delegate {
                bool isFixedPaper = cmbPaperSize.SelectedIndex != 1;
                cmbOrientation.Enabled = isFixedPaper;
                chkKeepRatio.Enabled = isFixedPaper;
                chkMargin.Enabled = isFixedPaper;
                UpdatePreview();
            };
            cmbOrientation.SelectedIndexChanged += delegate { UpdatePreview(); };
            chkKeepRatio.CheckedChanged += delegate { UpdatePreview(); };
            chkMargin.CheckedChanged += delegate { UpdatePreview(); };

            grpPaper.Controls.AddRange(new Control[] { lblPaper, cmbPaperSize, lblOrient, cmbOrientation, chkKeepRatio, chkMargin });
            bottomPanel.Controls.Add(grpPaper);

            // 3.2 输出路径
            Label lblOut = new Label
            {
                Text = "输出文件：",
                Location = new Point(15, 110),
                Size = new Size(75, 26),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            txtOutputPath = new TextBox
            {
                Location = new Point(90, 109),
                Size = new Size(815, 28),
                Font = new Font("Microsoft YaHei UI", 9.5F)
            };

            btnBrowseOutput = CreateButton("浏览...", 85, 28, Color.White, Color.FromArgb(51, 65, 85));
            btnBrowseOutput.Location = new Point(915, 108);
            btnBrowseOutput.Click += BtnBrowseOutput_Click;

            bottomPanel.Controls.AddRange(new Control[] { lblOut, txtOutputPath, btnBrowseOutput });

            // 3.3 进度条与“一键生成”大按钮 (绝对定位，完美显眼)
            lblStatus = new Label
            {
                Text = "就绪。可直接拖拽图片文件或文件夹到窗口中。",
                Location = new Point(15, 147),
                Size = new Size(780, 22),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(71, 85, 105)
            };

            progressBar = new ProgressBar
            {
                Location = new Point(15, 174),
                Size = new Size(780, 24)
            };

            btnGenerate = CreateButton("🚀 一键生成 PDF", 190, 52, Color.FromArgb(16, 185, 129), Color.White);
            btnGenerate.Font = new Font("Microsoft YaHei UI", 11.5F, FontStyle.Bold);
            btnGenerate.Location = new Point(810, 146);
            btnGenerate.Click += BtnGenerate_Click;

            bottomPanel.Controls.AddRange(new Control[] { lblStatus, progressBar, btnGenerate });

            // 将各主区域加入窗体
            this.Controls.Add(bottomPanel);
            this.Controls.Add(pnlPreview);
            this.Controls.Add(pnlList);
            this.Controls.Add(topBar);

            // BackgroundWorker for background processing
            worker = new BackgroundWorker { WorkerReportsProgress = true };
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private Button CreateButton(string text, int width, int height, Color backColor, Color foreColor)
        {
            Button btn = new Button
            {
                Text = text,
                Width = width,
                Height = height,
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = (backColor == Color.White) ? 1 : 0;
            btn.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            return btn;
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            HandleDroppedPaths(paths);
        }

        private void HandleDroppedPaths(string[] paths)
        {
            List<string> collected = new List<string>();
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                {
                    string[] extPatterns = new string[] { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp", "*.tif", "*.tiff" };
                    foreach (string ext in extPatterns)
                    {
                        collected.AddRange(Directory.GetFiles(p, ext, SearchOption.TopDirectoryOnly));
                    }
                }
                else if (File.Exists(p) && IsImageFile(p))
                {
                    collected.Add(p);
                }
            }

            collected.Sort(StrCmpLogicalW);
            AddImagePaths(collected);
        }

        private bool IsImageFile(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".webp" || ext == ".tif" || ext == ".tiff";
        }

        private void BtnAddFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "选择图片文件（可多选）";
                ofd.Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.webp;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.webp;*.tif;*.tiff|所有文件 (*.*)|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    List<string> files = new List<string>(ofd.FileNames);
                    files.Sort(StrCmpLogicalW);
                    AddImagePaths(files);
                }
            }
        }

        private void BtnAddFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "请选择包含图片的文件夹：";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    HandleDroppedPaths(new string[] { fbd.SelectedPath });
                }
            }
        }

        private void AddImagePaths(List<string> files)
        {
            foreach (string file in files)
            {
                try
                {
                    FileInfo fi = new FileInfo(file);
                    using (Image img = Image.FromFile(file))
                    {
                        items.Add(new ImageItem
                        {
                            FilePath = file,
                            FileName = fi.Name,
                            OrigWidth = img.Width,
                            OrigHeight = img.Height,
                            FileSizeBytes = fi.Length,
                            Rotation = 0
                        });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error reading " + file + ": " + ex.Message);
                }
            }

            UpdateListView();

            // 自动设定默认输出文件名
            if (string.IsNullOrEmpty(txtOutputPath.Text) && items.Count > 0)
            {
                string dir = Path.GetDirectoryName(items[0].FilePath);
                string folderName = new DirectoryInfo(dir).Name;
                txtOutputPath.Text = Path.Combine(dir, folderName + "_合并.pdf");
            }
        }

        private void UpdateListView()
        {
            listView.BeginUpdate();
            listView.Items.Clear();
            for (int i = 0; i < items.Count; i++)
            {
                ImageItem it = items[i];
                ListViewItem lvi = new ListViewItem((i + 1).ToString());
                lvi.SubItems.Add(it.FileName);
                lvi.SubItems.Add(it.CurrentWidth + " x " + it.CurrentHeight);
                lvi.SubItems.Add(it.IsLandscape ? "横版" : "竖版");
                lvi.SubItems.Add(it.Rotation + "°");
                lvi.SubItems.Add(FormatFileSize(it.FileSizeBytes));
                lvi.Tag = it;
                listView.Items.Add(lvi);
            }
            listView.EndUpdate();
            if (listView.Items.Count > 0 && listView.SelectedItems.Count == 0)
            {
                listView.Items[0].Selected = true;
            }
            lblStatus.Text = "共加载 " + items.Count + " 个文件，总大小约 " + FormatFileSize(GetTotalSize());
        }

        private long GetTotalSize()
        {
            long t = 0;
            foreach (var it in items) t += it.FileSizeBytes;
            return t;
        }

        private string FormatFileSize(long bytes)
        {
            if (bytes > 1024 * 1024)
                return (bytes / 1048576.0).ToString("F2") + " MB";
            return (bytes / 1024.0).ToString("F1") + " KB";
        }

        private void ListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (listView.SelectedItems.Count == 0 || items.Count == 0)
            {
                ClearPreview();
                return;
            }

            ImageItem it = (ImageItem)listView.SelectedItems[0].Tag;
            int curIdx = listView.SelectedIndices[0];

            try
            {
                using (Image src = Image.FromFile(it.FilePath))
                {
                    using (Bitmap rawBmp = new Bitmap(src))
                    {
                        if (it.Rotation == 90) rawBmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        else if (it.Rotation == 180) rawBmp.RotateFlip(RotateFlipType.Rotate180FlipNone);
                        else if (it.Rotation == 270) rawBmp.RotateFlip(RotateFlipType.Rotate270FlipNone);

                        int imgW = rawBmp.Width;
                        int imgH = rawBmp.Height;

                        int paperMode = cmbPaperSize != null ? cmbPaperSize.SelectedIndex : 0;
                        int orientMode = cmbOrientation != null ? cmbOrientation.SelectedIndex : 0;
                        bool keepRatio = chkKeepRatio != null ? chkKeepRatio.Checked : true;
                        bool hasMargin = chkMargin != null ? chkMargin.Checked : true;

                        double pageW, pageH, drawX, drawY, drawW, drawH;
                        CalculateLayout(imgW, imgH, paperMode, orientMode, keepRatio, hasMargin,
                            out pageW, out pageH, out drawX, out drawY, out drawW, out drawH);

                        int boxW = previewBox.ClientSize.Width;
                        int boxH = previewBox.ClientSize.Height;
                        if (boxW <= 0) boxW = 360;
                        if (boxH <= 0) boxH = 400;

                        Bitmap canvas = new Bitmap(boxW, boxH);
                        using (Graphics g = Graphics.FromImage(canvas))
                        {
                            g.SmoothingMode = SmoothingMode.HighQuality;
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                            // 桌面底色
                            g.Clear(Color.FromArgb(238, 242, 246));

                            // 拟真纸张缩放到预览区域 (留出16像素边距)
                            float pad = 16f;
                            float availW = boxW - pad * 2;
                            float availH = boxH - pad * 2;
                            float pageScale = (float)Math.Min(availW / pageW, availH / pageH);

                            float paperPixelW = (float)(pageW * pageScale);
                            float paperPixelH = (float)(pageH * pageScale);
                            float paperPixelX = (boxW - paperPixelW) / 2f;
                            float paperPixelY = (boxH - paperPixelH) / 2f;

                            RectangleF paperRect = new RectangleF(paperPixelX, paperPixelY, paperPixelW, paperPixelH);

                            // 纸张投影
                            RectangleF shadowRect = new RectangleF(paperPixelX + 4, paperPixelY + 4, paperPixelW, paperPixelH);
                            using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(45, 0, 0, 0)))
                            {
                                g.FillRectangle(shadowBrush, shadowRect);
                            }

                            // 纸张纯白背景与边框
                            g.FillRectangle(Brushes.White, paperRect);
                            using (Pen borderPen = new Pen(Color.FromArgb(203, 213, 225), 1))
                            {
                                g.DrawRectangle(borderPen, paperRect.X, paperRect.Y, paperRect.Width, paperRect.Height);
                            }

                            // 在纸张内渲染实际内容位置
                            // 在 PDF 中 (drawX, drawY) 为左下角；在 GDI+ 中 Y 轴向下，距离顶部为 pageH - drawY - drawH
                            float imgPixelX = paperPixelX + (float)(drawX * pageScale);
                            float imgPixelY = paperPixelY + (float)((pageH - drawY - drawH) * pageScale);
                            float imgPixelW = (float)(drawW * pageScale);
                            float imgPixelH = (float)(drawH * pageScale);

                            g.DrawImage(rawBmp, imgPixelX, imgPixelY, imgPixelW, imgPixelH);

                            // 若非原图且内容小于纸张，绘制极其微弱的蓝色辅助定位虚线框
                            if (paperMode != 1 && (imgPixelW < paperPixelW * 0.96f || imgPixelH < paperPixelH * 0.96f))
                            {
                                using (Pen outlinePen = new Pen(Color.FromArgb(70, 59, 130, 246), 1))
                                {
                                    outlinePen.DashStyle = DashStyle.Dash;
                                    g.DrawRectangle(outlinePen, imgPixelX, imgPixelY, imgPixelW, imgPixelH);
                                }
                            }
                        }

                        if (previewBox.Image != null) previewBox.Image.Dispose();
                        previewBox.Image = canvas;

                        string paperName = (paperMode == 1) ? "原图大小" : (paperMode == 2 ? "A3" : "A4");
                        string orientName = (pageW > pageH) ? "横版" : "竖版";
                        string layoutInfo = (paperMode == 1) ? "1:1无白边" : (drawW < pageW * 0.95 ? "居中显示" : "满幅适配");
                        lblPreviewInfo.Text = "第 " + (curIdx + 1) + " 页 | " + paperName + " " + orientName + " (" + layoutInfo + ") | " + it.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                lblPreviewInfo.Text = "预览加载失败: " + ex.Message;
            }
        }

        private static void CalculateLayout(
            int imgW, int imgH, int paperMode, int orientMode, bool keepRatio, bool hasMargin,
            out double pageW, out double pageH,
            out double drawX, out double drawY, out double drawW, out double drawH)
        {
            if (paperMode == 1) // 适合原图尺寸 (1:1，无白边)
            {
                pageW = imgW;
                pageH = imgH;
                drawX = 0;
                drawY = 0;
                drawW = pageW;
                drawH = pageH;
                return;
            }

            double shortSide = (paperMode == 2) ? 841.89 : 595.28;  // A3 or A4
            double longSide  = (paperMode == 2) ? 1190.55 : 841.89;

            bool isLandscape;
            if (orientMode == 0) // 统一纵向 (推荐)
            {
                isLandscape = false;
            }
            else if (orientMode == 2) // 统一横向
            {
                isLandscape = true;
            }
            else // 智能自适应
            {
                isLandscape = imgW > imgH;
            }

            pageW = isLandscape ? longSide : shortSide;
            pageH = isLandscape ? shortSide : longSide;

            double margin = hasMargin ? 30.0 : 0.0; // 30 pt (~10.5 mm 舒适边距)
            double availW = Math.Max(10.0, pageW - margin * 2.0);
            double availH = Math.Max(10.0, pageH - margin * 2.0);

            if (keepRatio)
            {
                double scale = Math.Min(availW / imgW, availH / imgH);
                drawW = imgW * scale;
                drawH = imgH * scale;
                drawX = (pageW - drawW) / 2.0;
                drawY = (pageH - drawH) / 2.0;
            }
            else
            {
                drawW = availW;
                drawH = availH;
                drawX = (pageW - drawW) / 2.0;
                drawY = (pageH - drawH) / 2.0;
            }
        }

        private void ClearPreview()
        {
            if (previewBox.Image != null)
            {
                previewBox.Image.Dispose();
                previewBox.Image = null;
            }
            lblPreviewInfo.Text = "点击左侧文件预览画面";
        }

        private void MoveItem(int direction)
        {
            if (listView.SelectedIndices.Count == 0) return;
            int idx = listView.SelectedIndices[0];
            int targetIdx = idx + direction;
            if (targetIdx < 0 || targetIdx >= items.Count) return;

            ImageItem it = items[idx];
            items.RemoveAt(idx);
            items.Insert(targetIdx, it);

            UpdateListView();
            listView.Items[targetIdx].Selected = true;
            listView.Items[targetIdx].EnsureVisible();
        }

        private void BtnRotate_Click(object sender, EventArgs e)
        {
            if (listView.SelectedIndices.Count == 0) return;
            int idx = listView.SelectedIndices[0];
            ImageItem it = items[idx];
            it.Rotation = (it.Rotation + 90) % 360;

            UpdateListView();
            listView.Items[idx].Selected = true;
            UpdatePreview();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (listView.SelectedIndices.Count == 0) return;
            int idx = listView.SelectedIndices[0];
            items.RemoveAt(idx);
            UpdateListView();
            ClearPreview();
        }

        private void BtnBrowseOutput_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "保存 PDF 文件至";
                sfd.Filter = "PDF 文件 (*.pdf)|*.pdf";
                if (!string.IsNullOrEmpty(txtOutputPath.Text))
                {
                    try
                    {
                        sfd.InitialDirectory = Path.GetDirectoryName(txtOutputPath.Text);
                        sfd.FileName = Path.GetFileName(txtOutputPath.Text);
                    }
                    catch { }
                }
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtOutputPath.Text = sfd.FileName;
                }
            }
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            if (items.Count == 0)
            {
                MessageBox.Show("请先添加需要转换的图片！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtOutputPath.Text))
            {
                MessageBox.Show("请指定输出 PDF 的文件路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnGenerate.Enabled = false;
            progressBar.Value = 0;
            lblStatus.Text = "准备生成 PDF...";

            bool isLossless = rbLossless.Checked;
            int quality = rbGov.Checked ? 75 : 100;
            int maxDim = rbGov.Checked ? 1920 : 0;

            GenerateParams p = new GenerateParams
            {
                Items = new List<ImageItem>(items),
                OutputPath = txtOutputPath.Text,
                IsLossless = isLossless,
                Quality = quality,
                MaxDimension = maxDim,
                PaperSizeMode = cmbPaperSize.SelectedIndex,
                OrientationMode = cmbOrientation.SelectedIndex,
                KeepAspectRatio = chkKeepRatio.Checked,
                HasMargin = chkMargin.Checked
            };

            worker.RunWorkerAsync(p);
        }

        private class GenerateParams
        {
            public List<ImageItem> Items;
            public string OutputPath;
            public bool IsLossless;
            public int Quality;
            public int MaxDimension;
            public int PaperSizeMode;
            public int OrientationMode;
            public bool KeepAspectRatio;
            public bool HasMargin;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            GenerateParams p = (GenerateParams)e.Argument;
            byte[] pdfBytes = BuildPdf(p, worker);
            File.WriteAllBytes(p.OutputPath, pdfBytes);
            e.Result = new FileInfo(p.OutputPath).Length;
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblStatus.Text = (string)e.UserState;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnGenerate.Enabled = true;
            progressBar.Value = 100;

            if (e.Error != null)
            {
                lblStatus.Text = "生成失败: " + e.Error.Message;
                MessageBox.Show("生成 PDF 时发生错误：\n" + e.Error.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            long finalSize = (long)e.Result;
            string sizeStr = FormatFileSize(finalSize);
            lblStatus.Text = "✅ 成功生成 PDF！文件大小: " + sizeStr;

            string msg = "PDF 文件生成成功！\n\n输出路径：" + txtOutputPath.Text + "\n文件体积：" + sizeStr + "\n总页数：" + items.Count + " 页\n\n是否立即打开该文件？";

            DialogResult dr = MessageBox.Show(msg, "处理完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(txtOutputPath.Text);
                }
                catch { }
            }
        }

        private byte[] BuildPdf(GenerateParams p, BackgroundWorker bw)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                List<long> offsets = new List<long>();
                offsets.Add(0); // Object 0 dummy

                // Header
                byte[] header = Encoding.ASCII.GetBytes("%PDF-1.4\r\n%\xE2\xE3\xCF\xD3\r\n");
                ms.Write(header, 0, header.Length);

                int n = p.Items.Count;

                // Object 1: Catalog
                offsets.Add(ms.Position);
                WriteString(ms, "1 0 obj\r\n<< /Type /Catalog /Pages 2 0 R >>\r\nendobj\r\n");

                // Object 2: Pages
                offsets.Add(ms.Position);
                StringBuilder kids = new StringBuilder();
                for (int i = 0; i < n; i++)
                {
                    kids.Append(3 + 3 * i).Append(" 0 R ");
                }
                WriteString(ms, "2 0 obj\r\n<< /Type /Pages /Kids [" + kids.ToString() + "] /Count " + n + " >>\r\nendobj\r\n");

                for (int i = 0; i < n; i++)
                {
                    ImageItem it = p.Items[i];
                    int progress = (int)((i + 0.5f) / n * 100);
                    bw.ReportProgress(progress, "正在处理第 " + (i + 1) + "/" + n + " 页: " + it.FileName + " ...");

                    byte[] imgBytes;
                    int finalW = it.CurrentWidth;
                    int finalH = it.CurrentHeight;

                    bool canUseDirectLossless = p.IsLossless && it.Rotation == 0 &&
                        (it.FilePath.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || it.FilePath.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase));

                    if (canUseDirectLossless)
                    {
                        imgBytes = File.ReadAllBytes(it.FilePath);
                    }
                    else
                    {
                        using (Image src = Image.FromFile(it.FilePath))
                        {
                            Bitmap bmp = new Bitmap(src);
                            if (it.Rotation == 90) bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
                            else if (it.Rotation == 180) bmp.RotateFlip(RotateFlipType.Rotate180FlipNone);
                            else if (it.Rotation == 270) bmp.RotateFlip(RotateFlipType.Rotate270FlipNone);

                            int targetW = bmp.Width;
                            int targetH = bmp.Height;

                            if (p.MaxDimension > 0)
                            {
                                int maxD = Math.Max(targetW, targetH);
                                if (maxD > p.MaxDimension)
                                {
                                    float scale = (float)p.MaxDimension / maxD;
                                    targetW = (int)(targetW * scale);
                                    targetH = (int)(targetH * scale);
                                }
                            }

                            using (Bitmap resized = new Bitmap(targetW, targetH))
                            {
                                using (Graphics g = Graphics.FromImage(resized))
                                {
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                    g.SmoothingMode = SmoothingMode.HighQuality;
                                    g.Clear(Color.White);
                                    g.DrawImage(bmp, 0, 0, targetW, targetH);
                                }

                                ImageCodecInfo encoder = GetEncoder(ImageFormat.Jpeg);
                                EncoderParameters encParams = new EncoderParameters(1);
                                encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)p.Quality);

                                using (MemoryStream imgMs = new MemoryStream())
                                {
                                    resized.Save(imgMs, encoder, encParams);
                                    imgBytes = imgMs.ToArray();
                                }
                                finalW = targetW;
                                finalH = targetH;
                            }
                            bmp.Dispose();
                        }
                    }

                    double pageW, pageH, drawX, drawY, drawW, drawH;
                    CalculateLayout(finalW, finalH, p.PaperSizeMode, p.OrientationMode, p.KeepAspectRatio, p.HasMargin,
                        out pageW, out pageH, out drawX, out drawY, out drawW, out drawH);

                    string pageWStr = pageW.ToString("F2", CultureInfo.InvariantCulture);
                    string pageHStr = pageH.ToString("F2", CultureInfo.InvariantCulture);
                    string drawWStr = drawW.ToString("F2", CultureInfo.InvariantCulture);
                    string drawHStr = drawH.ToString("F2", CultureInfo.InvariantCulture);
                    string drawXStr = drawX.ToString("F2", CultureInfo.InvariantCulture);
                    string drawYStr = drawY.ToString("F2", CultureInfo.InvariantCulture);

                    int pageObjId = 3 + 3 * i;
                    int imgObjId = 4 + 3 * i;
                    int contentObjId = 5 + 3 * i;

                    // Page Object
                    offsets.Add(ms.Position);
                    WriteString(ms, pageObjId + " 0 obj\r\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + pageWStr + " " + pageHStr + "] /Resources << /XObject << /Im1 " + imgObjId + " 0 R >> >> /Contents " + contentObjId + " 0 R >>\r\nendobj\r\n");

                    // Image Object
                    offsets.Add(ms.Position);
                    WriteString(ms, imgObjId + " 0 obj\r\n<< /Type /XObject /Subtype /Image /Width " + finalW + " /Height " + finalH + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + imgBytes.Length + " >>\r\nstream\r\n");
                    ms.Write(imgBytes, 0, imgBytes.Length);
                    WriteString(ms, "\r\nendstream\r\nendobj\r\n");

                    // Content Object
                    string contentStream = "q\r\n" + drawWStr + " 0 0 " + drawHStr + " " + drawXStr + " " + drawYStr + " cm\r\n/Im1 Do\r\nQ\r\n";
                    byte[] contentBytes = Encoding.ASCII.GetBytes(contentStream);

                    offsets.Add(ms.Position);
                    WriteString(ms, contentObjId + " 0 obj\r\n<< /Length " + contentBytes.Length + " >>\r\nstream\r\n" + contentStream + "endstream\r\nendobj\r\n");
                }

                // Cross-reference table
                long xrefOffset = ms.Position;
                WriteString(ms, "xref\r\n0 " + offsets.Count + "\r\n0000000000 65535 f \r\n");
                for (int i = 1; i < offsets.Count; i++)
                {
                    WriteString(ms, offsets[i].ToString("D10") + " 00000 n \r\n");
                }

                // Trailer
                WriteString(ms, "trailer\r\n<< /Size " + offsets.Count + " /Root 1 0 R >>\r\nstartxref\r\n" + xrefOffset + "\r\n%%EOF\r\n");

                return ms.ToArray();
            }
        }

        private static void WriteString(Stream s, string str)
        {
            byte[] b = Encoding.ASCII.GetBytes(str);
            s.Write(b, 0, b.Length);
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                    return codec;
            }
            return null;
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                SetProcessDPIAware();
            }
            catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args));
        }
    }
}
