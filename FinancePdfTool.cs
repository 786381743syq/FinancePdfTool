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
using System.Threading;

using System.Reflection;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

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

        // ================== 主选项卡 ==================
        private TabControl mainTabControl;
        private TabPage tabPageImgToPdf;
        private TabPage tabPagePdfToImg;

        // ================== Tab 1：图片合成 PDF 相关控件 ==================
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
        private BackgroundWorker workerImgToPdf;

        // ================== Tab 2：PDF 提取图片 相关控件 ==================
        private string currentPdfPath = null;
        private IPdfEngine currentPdfEngine = null;
        private int currentPdfPageCount = 0;
        private List<SizeF> currentPdfPageSizes = new List<SizeF>();
        private int currentSelectedPageIndex = -1;

        private Label lblPdfFileInfo;
        private ListView listViewPdfPages;
        private PictureBox previewBoxPdf;
        private Label lblCurrentPageInfo;
        private Button btnPrevPage;
        private Button btnNextPage;
        private RadioButton rbFormatPng;
        private RadioButton rbFormatJpg;
        private CheckBox chkAutoTrimPdf;
        private ComboBox cmbDpi;
        private RadioButton rbRangeAll;
        private RadioButton rbRangeChecked;
        private RadioButton rbRangeCurrent;
        private TextBox txtPdfOutputDir;
        private Button btnBrowsePdfOutputDir;
        private Button btnExportImages;
        private ProgressBar progressBarPdf;
        private Label lblPdfStatus;
        private BackgroundWorker workerPdfToImg;

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
            this.Text = "财务专用 PDF 转换器 - 图片与 PDF 互转神器";
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // 固定窗口大小，防止缩放变形
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.ClientSize = new Size(1020, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;
            this.BackColor = Color.FromArgb(245, 247, 250);

            // ================== 主 Tab 容器 ==================
            mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold),
                ItemSize = new Size(180, 36),
                SizeMode = TabSizeMode.Fixed
            };

            tabPageImgToPdf = new TabPage("📄 图片合成 PDF")
            {
                BackColor = Color.FromArgb(245, 247, 250),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular)
            };

            tabPagePdfToImg = new TabPage("🖼️ PDF 提取图片")
            {
                BackColor = Color.FromArgb(245, 247, 250),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular)
            };

            BuildTabImgToPdf();
            BuildTabPdfToImg();

            mainTabControl.TabPages.Add(tabPageImgToPdf);
            mainTabControl.TabPages.Add(tabPagePdfToImg);
            this.Controls.Add(mainTabControl);

            // 后台任务工作者
            workerImgToPdf = new BackgroundWorker { WorkerReportsProgress = true };
            workerImgToPdf.DoWork += WorkerImgToPdf_DoWork;
            workerImgToPdf.ProgressChanged += WorkerImgToPdf_ProgressChanged;
            workerImgToPdf.RunWorkerCompleted += WorkerImgToPdf_RunWorkerCompleted;

            workerPdfToImg = new BackgroundWorker { WorkerReportsProgress = true };
            workerPdfToImg.DoWork += WorkerPdfToImg_DoWork;
            workerPdfToImg.ProgressChanged += WorkerPdfToImg_ProgressChanged;
            workerPdfToImg.RunWorkerCompleted += WorkerPdfToImg_RunWorkerCompleted;
        }

        private void ShowHelpDialog(int initialTab = 0)
        {
            using (HelpForm dlg = new HelpForm(initialTab))
            {
                dlg.ShowDialog(this);
            }
        }

        #region ================== Tab 1：图片合成 PDF UI 与逻辑 ==================

        private void BuildTabImgToPdf()
        {
            // 1. 顶部工具栏
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 54),
                BackColor = Color.White
            };

            Button btnAddFiles = CreateButton("➕ 添加图片", 105, 34, Color.FromArgb(37, 99, 235), Color.White);
            btnAddFiles.Location = new Point(15, 10);
            btnAddFiles.Click += BtnAddFiles_Click;

            Button btnAddFolder = CreateButton("📁 添加文件夹", 115, 34, Color.FromArgb(71, 85, 105), Color.White);
            btnAddFolder.Location = new Point(130, 10);
            btnAddFolder.Click += BtnAddFolder_Click;

            Button btnMoveUp = CreateButton("⬆ 上移", 75, 34, Color.White, Color.FromArgb(51, 65, 85));
            btnMoveUp.Location = new Point(265, 10);
            btnMoveUp.Click += delegate { MoveItem(-1); };

            Button btnMoveDown = CreateButton("⬇ 下移", 75, 34, Color.White, Color.FromArgb(51, 65, 85));
            btnMoveDown.Location = new Point(350, 10);
            btnMoveDown.Click += delegate { MoveItem(1); };

            Button btnRotate = CreateButton("🔄 旋转90°", 95, 34, Color.White, Color.FromArgb(51, 65, 85));
            btnRotate.Location = new Point(435, 10);
            btnRotate.Click += BtnRotate_Click;

            Button btnRemove = CreateButton("❌ 移除选中", 95, 34, Color.White, Color.FromArgb(220, 38, 38));
            btnRemove.Location = new Point(540, 10);
            btnRemove.Click += BtnRemove_Click;

            Button btnClear = CreateButton("清空列表", 85, 34, Color.White, Color.FromArgb(100, 116, 139));
            btnClear.Location = new Point(645, 10);
            btnClear.Click += delegate {
                items.Clear();
                UpdateListView();
                ClearPreview();
            };

            Button btnHelpTab1 = CreateButton("💡 使用须知", 105, 34, Color.FromArgb(238, 242, 255), Color.FromArgb(67, 56, 202));
            btnHelpTab1.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
            btnHelpTab1.Location = new Point(892, 10);
            btnHelpTab1.Click += delegate { ShowHelpDialog(0); };

            topBar.Controls.AddRange(new Control[] {
                btnAddFiles, btnAddFolder,
                btnMoveUp, btnMoveDown, btnRotate,
                btnRemove, btnClear, btnHelpTab1
            });

            Panel dividerTop = new Panel
            {
                Location = new Point(0, 53),
                Size = new Size(1012, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            topBar.Controls.Add(dividerTop);

            // 2. 中间内容区 (左侧列表 + 右侧拟真预览)
            Panel pnlList = new Panel
            {
                Location = new Point(15, 64),
                Size = new Size(610, 420),
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

            Panel pnlPreview = new Panel
            {
                Location = new Point(635, 64),
                Size = new Size(360, 420),
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

            // 3. 底部配置与生成区
            Panel bottomPanel = new Panel
            {
                Location = new Point(0, 492),
                Size = new Size(1012, 215),
                BackColor = Color.White
            };

            Panel dividerBottom = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            bottomPanel.Controls.Add(dividerBottom);

            // 3.1 画质压缩选项
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

            // 3.2 纸张版式设置
            GroupBox grpPaper = new GroupBox
            {
                Text = "纸张大小与页面版式",
                Location = new Point(510, 6),
                Size = new Size(485, 96),
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
                Size = new Size(185, 26),
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

            // 3.3 输出路径
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
            btnBrowseOutput.Location = new Point(910, 108);
            btnBrowseOutput.Click += BtnBrowseOutput_Click;

            bottomPanel.Controls.AddRange(new Control[] { lblOut, txtOutputPath, btnBrowseOutput });

            // 3.4 进度条与“一键生成”大按钮
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

            btnGenerate = CreateButton("🚀 一键生成 PDF", 185, 52, Color.FromArgb(16, 185, 129), Color.White);
            btnGenerate.Font = new Font("Microsoft YaHei UI", 11.5F, FontStyle.Bold);
            btnGenerate.Location = new Point(810, 146);
            btnGenerate.Click += BtnGenerate_Click;

            bottomPanel.Controls.AddRange(new Control[] { lblStatus, progressBar, btnGenerate });

            tabPageImgToPdf.Controls.Add(bottomPanel);
            tabPageImgToPdf.Controls.Add(pnlPreview);
            tabPageImgToPdf.Controls.Add(pnlList);
            tabPageImgToPdf.Controls.Add(topBar);
        }

        #endregion

        #region ================== Tab 2：PDF 提取图片 UI 与逻辑 ==================

        private void BuildTabPdfToImg()
        {
            // 1. 顶部操作栏
            Panel topBarPdf = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 54),
                BackColor = Color.White
            };

            Button btnSelectPdf = CreateButton("📂 选择 PDF 文件", 135, 34, Color.FromArgb(37, 99, 235), Color.White);
            btnSelectPdf.Location = new Point(15, 10);
            btnSelectPdf.Click += BtnSelectPdf_Click;

            Button btnOpenPdfDir = CreateButton("📁 打开输出目录", 130, 34, Color.FromArgb(71, 85, 105), Color.White);
            btnOpenPdfDir.Location = new Point(160, 10);
            btnOpenPdfDir.Click += BtnOpenPdfDir_Click;

            Button btnClearPdf = CreateButton("清空", 75, 34, Color.White, Color.FromArgb(100, 116, 139));
            btnClearPdf.Location = new Point(300, 10);
            btnClearPdf.Click += delegate { ClearPdfView(); };

            lblPdfFileInfo = new Label
            {
                Location = new Point(390, 10),
                Size = new Size(490, 34),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(30, 41, 59),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                Text = "未加载任何 PDF 文件。请点击左侧按钮或直接拖拽 PDF 到此处。"
            };

            Button btnHelpTab2 = CreateButton("💡 使用须知", 105, 34, Color.FromArgb(238, 242, 255), Color.FromArgb(67, 56, 202));
            btnHelpTab2.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
            btnHelpTab2.Location = new Point(892, 10);
            btnHelpTab2.Click += delegate { ShowHelpDialog(1); };

            topBarPdf.Controls.AddRange(new Control[] {
                btnSelectPdf, btnOpenPdfDir, btnClearPdf, lblPdfFileInfo, btnHelpTab2
            });

            Panel dividerTop = new Panel
            {
                Location = new Point(0, 53),
                Size = new Size(1012, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            topBarPdf.Controls.Add(dividerTop);

            // 2. 中间内容区 (左侧页面列表 + 右侧页面实时预览)
            Panel pnlPdfPages = new Panel
            {
                Location = new Point(15, 64),
                Size = new Size(420, 420),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Panel pnlPageListHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = Color.FromArgb(248, 250, 252)
            };

            Label lblListTitle = new Label
            {
                Text = "📑 页面列表 (勾选导出)",
                Location = new Point(10, 8),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            Button btnSelectAll = CreateButton("全选", 50, 24, Color.White, Color.FromArgb(51, 65, 85));
            btnSelectAll.Location = new Point(235, 6);
            btnSelectAll.Font = new Font("Microsoft YaHei UI", 8.5F);
            btnSelectAll.Click += delegate { SetAllPagesChecked(true); };

            Button btnDeselectAll = CreateButton("全不选", 55, 24, Color.White, Color.FromArgb(51, 65, 85));
            btnDeselectAll.Location = new Point(292, 6);
            btnDeselectAll.Font = new Font("Microsoft YaHei UI", 8.5F);
            btnDeselectAll.Click += delegate { SetAllPagesChecked(false); };

            Button btnInvertSelect = CreateButton("反选", 50, 24, Color.White, Color.FromArgb(51, 65, 85));
            btnInvertSelect.Location = new Point(355, 6);
            btnInvertSelect.Font = new Font("Microsoft YaHei UI", 8.5F);
            btnInvertSelect.Click += delegate { InvertPagesChecked(); };

            pnlPageListHeader.Controls.AddRange(new Control[] { lblListTitle, btnSelectAll, btnDeselectAll, btnInvertSelect });

            listViewPdfPages = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            listViewPdfPages.Columns.Add("页码", 90, HorizontalAlignment.Left);
            listViewPdfPages.Columns.Add("尺寸", 115, HorizontalAlignment.Center);
            listViewPdfPages.Columns.Add("版式", 65, HorizontalAlignment.Center);
            listViewPdfPages.Columns.Add("画幅规格", 125, HorizontalAlignment.Center);
            listViewPdfPages.SelectedIndexChanged += ListViewPdfPages_SelectedIndexChanged;

            pnlPdfPages.Controls.Add(listViewPdfPages);
            pnlPdfPages.Controls.Add(pnlPageListHeader);

            // 右侧实时大图预览区
            Panel pnlPdfPreview = new Panel
            {
                Location = new Point(445, 64),
                Size = new Size(550, 420),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Panel pnlPreviewNav = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = Color.FromArgb(248, 250, 252)
            };

            btnPrevPage = CreateButton("◀ 上一页", 80, 26, Color.White, Color.FromArgb(51, 65, 85));
            btnPrevPage.Location = new Point(10, 5);
            btnPrevPage.Font = new Font("Microsoft YaHei UI", 9F);
            btnPrevPage.Click += delegate { NavigatePage(-1); };

            lblCurrentPageInfo = new Label
            {
                Location = new Point(100, 5),
                Size = new Size(345, 26),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 65, 85),
                Text = "点击左侧页面预览"
            };

            btnNextPage = CreateButton("下一页 ▶", 80, 26, Color.White, Color.FromArgb(51, 65, 85));
            btnNextPage.Location = new Point(455, 5);
            btnNextPage.Font = new Font("Microsoft YaHei UI", 9F);
            btnNextPage.Click += delegate { NavigatePage(1); };

            pnlPreviewNav.Controls.AddRange(new Control[] { btnPrevPage, lblCurrentPageInfo, btnNextPage });

            previewBoxPdf = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(238, 242, 246)
            };

            pnlPdfPreview.Controls.Add(previewBoxPdf);
            pnlPdfPreview.Controls.Add(pnlPreviewNav);

            // 3. 底部配置与生成区
            Panel bottomPanelPdf = new Panel
            {
                Location = new Point(0, 492),
                Size = new Size(1012, 215),
                BackColor = Color.White
            };

            Panel dividerBottomPdf = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 1),
                BackColor = Color.FromArgb(226, 232, 240)
            };
            bottomPanelPdf.Controls.Add(dividerBottomPdf);

            // 3.1 格式与清晰度 (左半部)
            GroupBox grpExportSettings = new GroupBox
            {
                Text = "导出格式与分辨率",
                Location = new Point(15, 6),
                Size = new Size(540, 96),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblFmt = new Label
            {
                Text = "格式：",
                Location = new Point(12, 27),
                Size = new Size(56, 20),
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            rbFormatPng = new RadioButton
            {
                Text = "⭐ PNG 高清无损",
                Location = new Point(74, 25),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            rbFormatJpg = new RadioButton
            {
                Text = "📁 JPG 通用压缩",
                Location = new Point(218, 25),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            chkAutoTrimPdf = new CheckBox
            {
                Text = "✂️ 智能去白边 (自动裁切)",
                Location = new Point(362, 25),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(79, 70, 229),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            chkAutoTrimPdf.CheckedChanged += delegate {
                if (currentPdfEngine != null && currentSelectedPageIndex >= 0)
                {
                    RenderPdfPagePreview(currentSelectedPageIndex);
                }
            };

            Label lblDpiTitle = new Label
            {
                Text = "清晰度：",
                Location = new Point(12, 59),
                Size = new Size(56, 20),
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            cmbDpi = new ComboBox
            {
                Location = new Point(74, 56),
                Size = new Size(450, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbDpi.Items.AddRange(new object[] {
                "300 DPI - 超清打印 (3x高精 · 票据公章发票极清 · 推荐)",
                "150 DPI - 高清阅读 (1.5x放大 · 适合日常浏览与归档)",
                "96 DPI - 标准轻量 (1x原生 · 适合快速传输与极小体积)"
            });
            cmbDpi.SelectedIndex = 0;

            grpExportSettings.Controls.AddRange(new Control[] {
                lblFmt, rbFormatPng, rbFormatJpg, chkAutoTrimPdf,
                lblDpiTitle, cmbDpi
            });
            bottomPanelPdf.Controls.Add(grpExportSettings);

            // 3.2 导出范围 (右半部)
            GroupBox grpRange = new GroupBox
            {
                Text = "导出范围",
                Location = new Point(565, 6),
                Size = new Size(430, 96),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            rbRangeAll = new RadioButton
            {
                Text = "导出全部页面",
                Location = new Point(15, 25),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            rbRangeChecked = new RadioButton
            {
                Text = "仅导出列表勾选页",
                Location = new Point(15, 58),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            rbRangeCurrent = new RadioButton
            {
                Text = "仅导出当前预览页",
                Location = new Point(175, 58),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            grpRange.Controls.AddRange(new Control[] { rbRangeAll, rbRangeChecked, rbRangeCurrent });
            bottomPanelPdf.Controls.Add(grpRange);

            // 3.3 输出目录
            Label lblOutDir = new Label
            {
                Text = "保存目录：",
                Location = new Point(15, 110),
                Size = new Size(75, 26),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            txtPdfOutputDir = new TextBox
            {
                Location = new Point(90, 109),
                Size = new Size(815, 28),
                Font = new Font("Microsoft YaHei UI", 9.5F)
            };

            btnBrowsePdfOutputDir = CreateButton("浏览...", 85, 28, Color.White, Color.FromArgb(51, 65, 85));
            btnBrowsePdfOutputDir.Location = new Point(910, 108);
            btnBrowsePdfOutputDir.Click += BtnBrowsePdfOutputDir_Click;

            bottomPanelPdf.Controls.AddRange(new Control[] { lblOutDir, txtPdfOutputDir, btnBrowsePdfOutputDir });

            // 3.4 进度条与“一键导出”大按钮
            lblPdfStatus = new Label
            {
                Text = "就绪。请选择或拖拽 PDF 文件到窗口中。",
                Location = new Point(15, 147),
                Size = new Size(760, 22),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(71, 85, 105)
            };

            progressBarPdf = new ProgressBar
            {
                Location = new Point(15, 174),
                Size = new Size(760, 24)
            };

            btnExportImages = CreateButton("🚀 导出为高清图片", 205, 52, Color.FromArgb(16, 185, 129), Color.White);
            btnExportImages.Font = new Font("Microsoft YaHei UI", 11.5F, FontStyle.Bold);
            btnExportImages.Location = new Point(790, 146);
            btnExportImages.Click += BtnExportImages_Click;

            bottomPanelPdf.Controls.AddRange(new Control[] { lblPdfStatus, progressBarPdf, btnExportImages });

            tabPagePdfToImg.Controls.Add(bottomPanelPdf);
            tabPagePdfToImg.Controls.Add(pnlPdfPreview);
            tabPagePdfToImg.Controls.Add(pnlPdfPages);
            tabPagePdfToImg.Controls.Add(topBarPdf);
        }

        private void BtnSelectPdf_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "选择需要转换为图片的 PDF 文件";
                ofd.Filter = "PDF 文件 (*.pdf)|*.pdf";
                ofd.Multiselect = false;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadPdfDocument(ofd.FileName);
                }
            }
        }

        private void BtnOpenPdfDir_Click(object sender, EventArgs e)
        {
            string dir = txtPdfOutputDir.Text;
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
            {
                try { System.Diagnostics.Process.Start("explorer.exe", dir); } catch { }
            }
            else if (!string.IsNullOrEmpty(currentPdfPath) && File.Exists(currentPdfPath))
            {
                try { System.Diagnostics.Process.Start("explorer.exe", Path.GetDirectoryName(currentPdfPath)); } catch { }
            }
            else
            {
                MessageBox.Show("尚未生成图片目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ClearPdfView()
        {
            currentPdfPath = null;
            if (currentPdfEngine != null)
            {
                currentPdfEngine.Close();
            }
            currentPdfPageCount = 0;
            currentPdfPageSizes.Clear();
            currentSelectedPageIndex = -1;

            lblPdfFileInfo.Text = "未加载任何 PDF 文件。请点击左侧按钮或直接拖拽 PDF 到此处。";
            listViewPdfPages.Items.Clear();
            ClearPdfPreview();
            txtPdfOutputDir.Text = "";
            lblPdfStatus.Text = "已清空。";
            progressBarPdf.Value = 0;
            rbRangeAll.Text = "导出全部页面";
        }

        private void ClearPdfPreview()
        {
            if (previewBoxPdf.Image != null)
            {
                previewBoxPdf.Image.Dispose();
                previewBoxPdf.Image = null;
            }
            lblCurrentPageInfo.Text = "点击左侧页面预览";
        }

        private void SetAllPagesChecked(bool check)
        {
            listViewPdfPages.BeginUpdate();
            foreach (ListViewItem item in listViewPdfPages.Items)
            {
                item.Checked = check;
            }
            listViewPdfPages.EndUpdate();
        }

        private void InvertPagesChecked()
        {
            listViewPdfPages.BeginUpdate();
            foreach (ListViewItem item in listViewPdfPages.Items)
            {
                item.Checked = !item.Checked;
            }
            listViewPdfPages.EndUpdate();
        }

        private void NavigatePage(int delta)
        {
            if (currentPdfPageCount <= 0) return;
            int newIdx = currentSelectedPageIndex + delta;
            if (newIdx >= 0 && newIdx < currentPdfPageCount)
            {
                listViewPdfPages.Items[newIdx].Selected = true;
                listViewPdfPages.Items[newIdx].EnsureVisible();
            }
        }

        private void ListViewPdfPages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewPdfPages.SelectedIndices.Count == 0) return;
            int pageIdx = listViewPdfPages.SelectedIndices[0];
            RenderPdfPagePreview(pageIdx);
        }

        private void LoadPdfDocument(string path)
        {
            if (!File.Exists(path)) return;
            try
            {
                currentPdfPath = path;
                FileInfo fi = new FileInfo(path);
                lblPdfFileInfo.Text = "已加载：" + fi.Name + " (" + FormatFileSize(fi.Length) + ")";

                string dir = Path.GetDirectoryName(path);
                string nameWithoutExt = Path.GetFileNameWithoutExtension(path);
                txtPdfOutputDir.Text = Path.Combine(dir, nameWithoutExt + "_图片");

                lblPdfStatus.Text = "正在解析 PDF 页面...";
                Application.DoEvents();

                if (currentPdfEngine == null)
                {
                    currentPdfEngine = PdfEngineFactory.Create();
                }

                bool loaded = currentPdfEngine.Load(path);
                currentPdfPageCount = currentPdfEngine.PageCount;

                if (!loaded || currentPdfPageCount == 0)
                {
                    if (!PdfEngineFactory.IsWin10WinRtAvailable())
                    {
                        MessageBox.Show("提示：当前运行环境为 Windows 7 系统。\n\n由于 Windows 7 未内置 Direct2D 矢量 PDF 解析引擎：\n• 对于扫描版单据、照片发票或图片型 PDF，工具可自动极速提取高清原图；\n• 对于纯矢量排版的发票单据，建议在 Windows 10 或 Windows 11 电脑上运行导出，可享受系统级 300 DPI 超清硬件加速渲染！", "系统兼容提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        lblPdfStatus.Text = "当前系统为 Windows 7，未检测到扫描图片。建议在 Windows 10/11 电脑上运行以解析矢量 PDF。";
                    }
                    else
                    {
                        MessageBox.Show("未能成功解析该 PDF 文件，文件可能损坏或受密码保护。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        lblPdfStatus.Text = "解析 PDF 失败。";
                    }
                    return;
                }

                listViewPdfPages.BeginUpdate();
                listViewPdfPages.Items.Clear();
                currentPdfPageSizes.Clear();

                for (int i = 0; i < currentPdfPageCount; i++)
                {
                    SizeF size = currentPdfEngine.GetPageSize(i);
                    currentPdfPageSizes.Add(size);

                    bool isLandscape = size.Width > size.Height;
                    string orient = isLandscape ? "横版" : "竖版";
                    string dimStr = Math.Round(size.Width) + " x " + Math.Round(size.Height) + " pt";

                    double ratio = size.Height > 0 ? (double)size.Width / size.Height : 1.0;
                    string note = "";
                    if (Math.Abs(ratio - 0.707) < 0.05 || Math.Abs(ratio - 1.414) < 0.05)
                        note = "A4/A3 国际标准";
                    else if (Math.Abs(ratio - 1.0) < 0.05)
                        note = "正方形";
                    else
                        note = "自定义画幅";

                    ListViewItem lvi = new ListViewItem("第 " + (i + 1) + " 页");
                    lvi.SubItems.Add(dimStr);
                    lvi.SubItems.Add(orient);
                    lvi.SubItems.Add(note);
                    lvi.Checked = true;
                    lvi.Tag = i;
                    listViewPdfPages.Items.Add(lvi);
                }
                listViewPdfPages.EndUpdate();

                lblPdfFileInfo.Text = "已加载：" + fi.Name + " (" + FormatFileSize(fi.Length) + " · 共 " + currentPdfPageCount + " 页)";
                lblPdfStatus.Text = "就绪。共 " + currentPdfPageCount + " 页，默认已全选，可随时点击导出。";
                rbRangeAll.Text = "导出全部页面 (共 " + currentPdfPageCount + " 页)";

                if (listViewPdfPages.Items.Count > 0)
                {
                    listViewPdfPages.Items[0].Selected = true;
                }
            }
            catch (Exception ex)
            {
                lblPdfStatus.Text = "加载 PDF 失败: " + ex.Message;
                MessageBox.Show("加载 PDF 失败：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderPdfPagePreview(int pageIndex)
        {
            if (currentPdfEngine == null || pageIndex < 0 || pageIndex >= currentPdfPageCount)
            {
                ClearPdfPreview();
                return;
            }

            try
            {
                currentSelectedPageIndex = pageIndex;
                lblCurrentPageInfo.Text = "第 " + (pageIndex + 1) + " / " + currentPdfPageCount + " 页";

                int boxW = previewBoxPdf.ClientSize.Width;
                int boxH = previewBoxPdf.ClientSize.Height;
                if (boxW <= 0) boxW = 530;
                if (boxH <= 0) boxH = 370;

                SizeF pageSize = currentPdfPageSizes[pageIndex];
                float scale = Math.Min((float)boxW * 1.5f / pageSize.Width, (float)boxH * 1.5f / pageSize.Height);
                if (scale < 0.2f) scale = 0.5f;
                if (scale > 2.5f) scale = 2.5f;

                using (Bitmap rendered = currentPdfEngine.RenderPage(pageIndex, scale))
                {
                    if (rendered == null) return;
                    Image drawImg = rendered;
                    Bitmap previewCropped = null;
                    bool wasTrimmed = false;

                    if (chkAutoTrimPdf != null && chkAutoTrimPdf.Checked)
                    {
                        using (Bitmap tempBmp = new Bitmap(rendered))
                        {
                            Rectangle cropRect = DetectContentBounds(tempBmp);
                            if (cropRect.Width > 30 && cropRect.Height > 30 &&
                                (cropRect.Width < tempBmp.Width * 0.97f || cropRect.Height < tempBmp.Height * 0.97f))
                            {
                                previewCropped = tempBmp.Clone(cropRect, tempBmp.PixelFormat);
                                drawImg = previewCropped;
                                wasTrimmed = true;
                            }
                        }
                    }

                    try
                    {
                        Bitmap canvas = new Bitmap(boxW, boxH);
                        using (Graphics g = Graphics.FromImage(canvas))
                        {
                            g.Clear(Color.FromArgb(238, 242, 246));
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = SmoothingMode.HighQuality;

                            float pad = 12f;
                            float availW = boxW - pad * 2;
                            float availH = boxH - pad * 2;
                            float fitScale = Math.Min(availW / drawImg.Width, availH / drawImg.Height);

                            float drawW = drawImg.Width * fitScale;
                            float drawH = drawImg.Height * fitScale;
                            float drawX = (boxW - drawW) / 2f;
                            float drawY = (boxH - drawH) / 2f;

                            using (SolidBrush shadow = new SolidBrush(Color.FromArgb(40, 0, 0, 0)))
                            {
                                g.FillRectangle(shadow, drawX + 4, drawY + 4, drawW, drawH);
                            }

                            g.FillRectangle(Brushes.White, drawX, drawY, drawW, drawH);
                            g.DrawImage(drawImg, drawX, drawY, drawW, drawH);

                            using (Pen borderPen = new Pen(Color.FromArgb(203, 213, 225), 1))
                            {
                                g.DrawRectangle(borderPen, drawX, drawY, drawW, drawH);
                            }
                        }

                        if (previewBoxPdf.Image != null) previewBoxPdf.Image.Dispose();
                        previewBoxPdf.Image = canvas;

                        string trimStatus = wasTrimmed ? " [✂️已智能去白边]" : "";
                        lblCurrentPageInfo.Text = "第 " + (pageIndex + 1) + " / " + currentPdfPageCount + " 页" + trimStatus;
                    }
                    finally
                    {
                        if (previewCropped != null) previewCropped.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                lblCurrentPageInfo.Text = "预览加载失败: " + ex.Message;
            }
        }

        private void BtnBrowsePdfOutputDir_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "选择导出的图片保存文件夹";
                if (!string.IsNullOrEmpty(txtPdfOutputDir.Text))
                {
                    try { fbd.SelectedPath = txtPdfOutputDir.Text; } catch { }
                }
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPdfOutputDir.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnExportImages_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentPdfPath) || !File.Exists(currentPdfPath))
            {
                MessageBox.Show("请先打开需要转换的 PDF 文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtPdfOutputDir.Text))
            {
                MessageBox.Show("请指定图片保存的输出目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<int> pagesToExport = new List<int>();
            if (rbRangeAll.Checked)
            {
                for (int i = 0; i < currentPdfPageCount; i++) pagesToExport.Add(i);
            }
            else if (rbRangeCurrent.Checked)
            {
                if (currentSelectedPageIndex >= 0 && currentSelectedPageIndex < currentPdfPageCount)
                {
                    pagesToExport.Add(currentSelectedPageIndex);
                }
                else
                {
                    pagesToExport.Add(0);
                }
            }
            else // rbRangeChecked
            {
                foreach (ListViewItem item in listViewPdfPages.Items)
                {
                    if (item.Checked)
                    {
                        pagesToExport.Add((int)item.Tag);
                    }
                }
            }

            if (pagesToExport.Count == 0)
            {
                MessageBox.Show("未选中任何需要导出的页面！请勾选至少一个页面。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnExportImages.Enabled = false;
            progressBarPdf.Value = 0;
            lblPdfStatus.Text = "准备导出高清图片...";

            PdfExportParams p = new PdfExportParams
            {
                PdfPath = currentPdfPath,
                OutputDir = txtPdfOutputDir.Text,
                IsPng = rbFormatPng.Checked,
                DpiMode = cmbDpi.SelectedIndex,
                AutoTrim = chkAutoTrimPdf.Checked,
                PageIndices = pagesToExport
            };

            workerPdfToImg.RunWorkerAsync(p);
        }

        private class PdfExportParams
        {
            public string PdfPath;
            public string OutputDir;
            public bool IsPng;
            public int DpiMode;
            public bool AutoTrim;
            public List<int> PageIndices;
        }

        private void WorkerPdfToImg_DoWork(object sender, DoWorkEventArgs e)
        {
            PdfExportParams p = (PdfExportParams)e.Argument;
            if (!Directory.Exists(p.OutputDir))
            {
                Directory.CreateDirectory(p.OutputDir);
            }

            using (IPdfEngine engine = PdfEngineFactory.Create())
            {
                if (!engine.Load(p.PdfPath))
                {
                    throw new InvalidOperationException("无法解析该 PDF 文件进行导出。");
                }

                float dpiScale = 4.1667f;
                if (p.DpiMode == 1) dpiScale = 2.0833f;
                else if (p.DpiMode == 2) dpiScale = 1.3333f;

                string pdfBaseName = Path.GetFileNameWithoutExtension(p.PdfPath);
                int total = p.PageIndices.Count;
                int digits = total > 99 ? 3 : 2;

                List<string> exportedFiles = new List<string>();

                for (int i = 0; i < total; i++)
                {
                    int pageIdx = p.PageIndices[i];
                    int progress = (int)((i + 0.5f) / total * 100);
                    workerPdfToImg.ReportProgress(progress, "正在导出第 " + (i + 1) + "/" + total + " 页 (PDF第 " + (pageIdx + 1) + " 页)...");

                    using (Bitmap rendered = engine.RenderPage(pageIdx, dpiScale))
                    {
                        if (rendered == null) continue;

                        using (Bitmap finalBmp = new Bitmap(rendered.Width, rendered.Height, PixelFormat.Format32bppRgb))
                        {
                            using (Graphics g = Graphics.FromImage(finalBmp))
                            {
                                g.Clear(Color.White);
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                g.DrawImage(rendered, 0, 0, rendered.Width, rendered.Height);
                            }

                            string pageNumStr = (pageIdx + 1).ToString().PadLeft(digits, '0');
                            string ext = p.IsPng ? ".png" : ".jpg";
                            string outFileName = pdfBaseName + "_第" + pageNumStr + "页" + ext;
                            string outFilePath = Path.Combine(p.OutputDir, outFileName);

                            Bitmap saveBmp = finalBmp;
                            Bitmap croppedBmp = null;
                            if (p.AutoTrim)
                            {
                                Rectangle contentRect = DetectContentBounds(finalBmp);
                                if (contentRect.Width > 50 && contentRect.Height > 50 &&
                                    (contentRect.Width < finalBmp.Width * 0.97f || contentRect.Height < finalBmp.Height * 0.97f))
                                {
                                    croppedBmp = finalBmp.Clone(contentRect, finalBmp.PixelFormat);
                                    saveBmp = croppedBmp;
                                }
                            }

                            try
                            {
                                if (p.IsPng)
                                {
                                    saveBmp.Save(outFilePath, ImageFormat.Png);
                                }
                                else
                                {
                                    ImageCodecInfo encoder = GetEncoder(ImageFormat.Jpeg);
                                    EncoderParameters encParams = new EncoderParameters(1);
                                    encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 92L);
                                    saveBmp.Save(outFilePath, encoder, encParams);
                                }
                            }
                            finally
                            {
                                if (croppedBmp != null) croppedBmp.Dispose();
                            }

                            exportedFiles.Add(outFilePath);
                        }
                    }
                }

                e.Result = new object[] { exportedFiles.Count, p.OutputDir };
            }
        }

        private void WorkerPdfToImg_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarPdf.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblPdfStatus.Text = (string)e.UserState;
        }

        private void WorkerPdfToImg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExportImages.Enabled = true;
            progressBarPdf.Value = 100;

            if (e.Error != null)
            {
                lblPdfStatus.Text = "导出失败: " + e.Error.Message;
                MessageBox.Show("导出图片时发生错误：\n" + e.Error.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            object[] res = (object[])e.Result;
            int count = (int)res[0];
            string outDir = (string)res[1];

            lblPdfStatus.Text = "✅ 成功导出 " + count + " 张高清图片至：" + outDir;

            string msg = "图片导出完成！\n\n共成功导出：" + count + " 张高清图片\n保存目录：" + outDir + "\n\n是否立即打开该文件夹？";
            DialogResult dr = MessageBox.Show(msg, "处理完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", outDir);
                }
                catch { }
            }
        }

        #endregion



        #region ================== 全局拖拽与通用事件 ==================

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths != null && paths.Length > 0)
                {
                    HandleDroppedPaths(paths);
                }
            }
        }

        private void HandleDroppedPaths(string[] paths)
        {
            if (paths == null || paths.Length == 0) return;

            string firstPdf = null;
            foreach (string p in paths)
            {
                if (File.Exists(p) && p.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    firstPdf = p;
                    break;
                }
            }

            if (firstPdf != null)
            {
                mainTabControl.SelectedTab = tabPagePdfToImg;
                LoadPdfDocument(firstPdf);
            }
            else
            {
                mainTabControl.SelectedTab = tabPageImgToPdf;
                AddFilesOrDirectories(paths);
            }
        }

        #endregion

        #region ================== 图片合成 PDF 辅助逻辑 ==================

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
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            btn.FlatAppearance.BorderSize = (backColor == Color.White) ? 1 : 0;
            btn.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            return btn;
        }

        private void BtnAddFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "选择图片文件";
                ofd.Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.tif|所有文件 (*.*)|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    AddFilesOrDirectories(ofd.FileNames);
                }
            }
        }

        private void BtnAddFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "选择包含票据/凭证图片的文件夹";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    AddFilesOrDirectories(new string[] { fbd.SelectedPath });
                }
            }
        }

        private void AddFilesOrDirectories(string[] paths)
        {
            List<string> candidateFiles = new List<string>();
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                {
                    string[] exts = new string[] { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif" };
                    foreach (string ext in exts)
                    {
                        try
                        {
                            candidateFiles.AddRange(Directory.GetFiles(p, ext, SearchOption.TopDirectoryOnly));
                        }
                        catch { }
                    }
                }
                else if (File.Exists(p))
                {
                    string ext = Path.GetExtension(p).ToLowerInvariant();
                    if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".tiff" || ext == ".tif")
                    {
                        candidateFiles.Add(p);
                    }
                }
            }

            candidateFiles.Sort(new Comparison<string>(StrCmpLogicalW));

            foreach (string file in candidateFiles)
            {
                try
                {
                    using (Image img = Image.FromFile(file))
                    {
                        items.Add(new ImageItem
                        {
                            FilePath = file,
                            FileName = Path.GetFileName(file),
                            OrigWidth = img.Width,
                            OrigHeight = img.Height,
                            FileSizeBytes = new FileInfo(file).Length,
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

                            g.Clear(Color.FromArgb(238, 242, 246));

                            float pad = 16f;
                            float availW = boxW - pad * 2;
                            float availH = boxH - pad * 2;
                            float pageScale = (float)Math.Min(availW / pageW, availH / pageH);

                            float paperPixelW = (float)(pageW * pageScale);
                            float paperPixelH = (float)(pageH * pageScale);
                            float paperPixelX = (boxW - paperPixelW) / 2f;
                            float paperPixelY = (boxH - paperPixelH) / 2f;

                            RectangleF paperRect = new RectangleF(paperPixelX, paperPixelY, paperPixelW, paperPixelH);

                            RectangleF shadowRect = new RectangleF(paperPixelX + 4, paperPixelY + 4, paperPixelW, paperPixelH);
                            using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(45, 0, 0, 0)))
                            {
                                g.FillRectangle(shadowBrush, shadowRect);
                            }

                            g.FillRectangle(Brushes.White, paperRect);
                            using (Pen borderPen = new Pen(Color.FromArgb(203, 213, 225), 1))
                            {
                                g.DrawRectangle(borderPen, paperRect.X, paperRect.Y, paperRect.Width, paperRect.Height);
                            }

                            float imgPixelX = paperPixelX + (float)(drawX * pageScale);
                            float imgPixelY = paperPixelY + (float)((pageH - drawY - drawH) * pageScale);
                            float imgPixelW = (float)(drawW * pageScale);
                            float imgPixelH = (float)(drawH * pageScale);

                            g.DrawImage(rawBmp, imgPixelX, imgPixelY, imgPixelW, imgPixelH);

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
            if (paperMode == 1)
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
            if (orientMode == 0)
            {
                isLandscape = false;
            }
            else if (orientMode == 2)
            {
                isLandscape = true;
            }
            else
            {
                isLandscape = imgW > imgH;
            }

            pageW = isLandscape ? longSide : shortSide;
            pageH = isLandscape ? shortSide : longSide;

            double margin = hasMargin ? 30.0 : 0.0;
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

            workerImgToPdf.RunWorkerAsync(p);
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

        private void WorkerImgToPdf_DoWork(object sender, DoWorkEventArgs e)
        {
            GenerateParams p = (GenerateParams)e.Argument;
            byte[] pdfBytes = BuildPdf(p, workerImgToPdf);
            File.WriteAllBytes(p.OutputPath, pdfBytes);
            e.Result = new FileInfo(p.OutputPath).Length;
        }

        private void WorkerImgToPdf_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblStatus.Text = (string)e.UserState;
        }

        private void WorkerImgToPdf_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
                offsets.Add(0);

                byte[] header = Encoding.ASCII.GetBytes("%PDF-1.4\r\n%\xE2\xE3\xCF\xD3\r\n");
                ms.Write(header, 0, header.Length);

                int n = p.Items.Count;

                offsets.Add(ms.Position);
                WriteString(ms, "1 0 obj\r\n<< /Type /Catalog /Pages 2 0 R >>\r\nendobj\r\n");

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

                    offsets.Add(ms.Position);
                    WriteString(ms, pageObjId + " 0 obj\r\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + pageWStr + " " + pageHStr + "] /Resources << /XObject << /Im1 " + imgObjId + " 0 R >> >> /Contents " + contentObjId + " 0 R >>\r\nendobj\r\n");

                    offsets.Add(ms.Position);
                    WriteString(ms, imgObjId + " 0 obj\r\n<< /Type /XObject /Subtype /Image /Width " + finalW + " /Height " + finalH + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + imgBytes.Length + " >>\r\nstream\r\n");
                    ms.Write(imgBytes, 0, imgBytes.Length);
                    WriteString(ms, "\r\nendstream\r\nendobj\r\n");

                    string contentStream = "q\r\n" + drawWStr + " 0 0 " + drawHStr + " " + drawXStr + " " + drawYStr + " cm\r\n/Im1 Do\r\nQ\r\n";
                    byte[] contentBytes = Encoding.ASCII.GetBytes(contentStream);

                    offsets.Add(ms.Position);
                    WriteString(ms, contentObjId + " 0 obj\r\n<< /Length " + contentBytes.Length + " >>\r\nstream\r\n" + contentStream + "endstream\r\nendobj\r\n");
                }

                long xrefOffset = ms.Position;
                WriteString(ms, "xref\r\n0 " + offsets.Count + "\r\n0000000000 65535 f \r\n");
                for (int i = 1; i < offsets.Count; i++)
                {
                    WriteString(ms, offsets[i].ToString("D10") + " 00000 n \r\n");
                }

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

        /// <summary>
        /// 智能扫描图片内容边界，自动去除四周空白留白（纯白/灰白背景）。
        /// 专用于银行回单、电子发票等在 A4 扫描或打印底版上的去白边。
        /// </summary>
        private static Rectangle DetectContentBounds(Bitmap bmp, int threshold = 240)
        {
            if (bmp == null) return Rectangle.Empty;

            int width = bmp.Width;
            int height = bmp.Height;

            BitmapData data = null;
            try
            {
                data = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                int stride = Math.Abs(data.Stride);
                byte[] buffer = new byte[stride * height];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);

                int minX = width, minY = height, maxX = -1, maxY = -1;

                // 逐行快速扫描
                for (int y = 0; y < height; y++)
                {
                    int rowOffset = y * stride;
                    for (int x = 0; x < width; x++)
                    {
                        int idx = rowOffset + (x * 4);
                        byte b = buffer[idx];
                        byte g = buffer[idx + 1];
                        byte r = buffer[idx + 2];
                        byte a = buffer[idx + 3];

                        // 如果是不透明且非纯白（任意通道低于阈值，即视为内容）
                        if (a > 30 && (r < threshold || g < threshold || b < threshold))
                        {
                            if (x < minX) minX = x;
                            if (x > maxX) maxX = x;
                            if (y < minY) minY = y;
                            if (y > maxY) maxY = y;
                        }
                    }
                }

                // 如果整张图都是空白，或扫描无效，保持原图
                if (maxX < minX || maxY < minY)
                {
                    return new Rectangle(0, 0, width, height);
                }

                int contentW = maxX - minX + 1;
                int contentH = maxY - minY + 1;

                // 如果内容本身已经占满 97% 以上宽和高，说明无需去白边
                if (contentW > width * 0.97 && contentH > height * 0.97)
                {
                    return new Rectangle(0, 0, width, height);
                }

                // 预留舒适自然的安全内边距（自适应，约1.5%）
                int pad = Math.Max(12, (int)(Math.Min(width, height) * 0.015));
                int cropX = Math.Max(0, minX - pad);
                int cropY = Math.Max(0, minY - pad);
                int cropW = Math.Min(width - cropX, contentW + (pad * 2));
                int cropH = Math.Min(height - cropY, contentH + (pad * 2));

                return new Rectangle(cropX, cropY, cropW, cropH);
            }
            catch
            {
                return new Rectangle(0, 0, width, height);
            }
            finally
            {
                if (data != null)
                {
                    bmp.UnlockBits(data);
                }
            }
        }

        #endregion

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                {
                    try { MessageBox.Show("程序运行发生异常：\n" + (e.ExceptionObject != null ? e.ExceptionObject.ToString() : "未知错误"), "运行提示", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                };
            }
            catch { }

            try
            {
                SetProcessDPIAware();
            }
            catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new MainForm(args));
            }
            catch (Exception ex)
            {
                MessageBox.Show("程序启动错误：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    #region ================== 跨版本兼容 PDF 引擎架构 (支持 Win7/8/10/11) ==================

    public interface IPdfEngine : IDisposable
    {
        bool Load(string pdfPath);
        int PageCount { get; }
        SizeF GetPageSize(int pageIndex);
        Bitmap RenderPage(int pageIndex, float scale);
        void Close();
    }

    public class DynamicEngineWrapper : IPdfEngine
    {
        private object rawEngine;
        private MethodInfo mLoad;
        private MethodInfo mGetPageCount;
        private MethodInfo mGetPageSize;
        private MethodInfo mRenderPage;
        private MethodInfo mClose;

        public DynamicEngineWrapper(object raw)
        {
            this.rawEngine = raw;
            Type t = raw.GetType();
            mLoad = t.GetMethod("Load");
            mGetPageCount = t.GetMethod("GetPageCount");
            mGetPageSize = t.GetMethod("GetPageSize");
            mRenderPage = t.GetMethod("RenderPage");
            mClose = t.GetMethod("Close");
        }

        public bool Load(string pdfPath)
        {
            try { return (bool)mLoad.Invoke(rawEngine, new object[] { pdfPath }); }
            catch { return false; }
        }

        public int PageCount
        {
            get
            {
                try { return (int)mGetPageCount.Invoke(rawEngine, null); }
                catch { return 0; }
            }
        }

        public SizeF GetPageSize(int pageIndex)
        {
            try { return (SizeF)mGetPageSize.Invoke(rawEngine, new object[] { pageIndex }); }
            catch { return SizeF.Empty; }
        }

        public Bitmap RenderPage(int pageIndex, float scale)
        {
            try { return (Bitmap)mRenderPage.Invoke(rawEngine, new object[] { pageIndex, scale }); }
            catch { return null; }
        }

        public void Close()
        {
            try { mClose.Invoke(rawEngine, null); }
            catch { }
        }

        public void Dispose()
        {
            Close();
        }
    }

    public static class PdfEngineFactory
    {
        private static Type cachedWinRtType = null;
        private static bool winRtInitAttempted = false;

        public static bool IsWin10WinRtAvailable()
        {
            try
            {
                if (Environment.OSVersion.Version.Major < 10 && !(Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 2))
                    return false;

                string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                string winData = Path.Combine(winDir, @"System32\WinMetadata\Windows.Data.winmd");
                return File.Exists(winData);
            }
            catch { return false; }
        }

        public static IPdfEngine Create()
        {
            if (IsWin10WinRtAvailable())
            {
                try
                {
                    if (!winRtInitAttempted)
                    {
                        winRtInitAttempted = true;
                        InitWinRtEngine();
                    }
                    if (cachedWinRtType != null)
                    {
                        object inst = Activator.CreateInstance(cachedWinRtType);
                        return new DynamicEngineWrapper(inst);
                    }
                }
                catch { }
            }

            return new FallbackPdfEngine();
        }

        private static void InitWinRtEngine()
        {
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string netDir = Path.Combine(winDir, @"Microsoft.NET\Framework64\v4.0.30319");
            if (!Directory.Exists(netDir)) netDir = Path.Combine(winDir, @"Microsoft.NET\Framework\v4.0.30319");
            string metaDir = Path.Combine(winDir, @"System32\WinMetadata");

            string code = @"
using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.Foundation;

namespace DynamicWinRt
{
    public class Engine
    {
        private PdfDocument doc;
        private List<SizeF> pageSizes = new List<SizeF>();

        public bool Load(string pdfPath)
        {
            Close();
            try
            {
                var storageFile = AwaitOp(StorageFile.GetFileFromPathAsync(pdfPath));
                doc = AwaitOp(PdfDocument.LoadFromFileAsync(storageFile));
                int count = (int)doc.PageCount;
                for (uint i = 0; i < (uint)count; i++)
                {
                    using (var page = doc.GetPage(i))
                    {
                        pageSizes.Add(new SizeF((float)page.Size.Width, (float)page.Size.Height));
                    }
                }
                return true;
            }
            catch
            {
                Close();
                return false;
            }
        }

        public int GetPageCount() { return doc != null ? (int)doc.PageCount : 0; }

        public SizeF GetPageSize(int index)
        {
            if (index >= 0 && index < pageSizes.Count) return pageSizes[index];
            return SizeF.Empty;
        }

        public Bitmap RenderPage(int index, float scale)
        {
            if (doc == null || index < 0 || index >= (int)doc.PageCount) return null;
            try
            {
                using (var page = doc.GetPage((uint)index))
                {
                    uint renderW = (uint)Math.Max(1, (int)Math.Round(page.Size.Width * scale));
                    uint renderH = (uint)Math.Max(1, (int)Math.Round(page.Size.Height * scale));

                    var options = new PdfPageRenderOptions();
                    options.DestinationWidth = renderW;
                    options.DestinationHeight = renderH;

                    using (var stream = new InMemoryRandomAccessStream())
                    {
                        AwaitAction(page.RenderToStreamAsync(stream, options));
                        using (var reader = new DataReader(stream.GetInputStreamAt(0)))
                        {
                            AwaitOp(reader.LoadAsync((uint)stream.Size));
                            byte[] bytes = new byte[stream.Size];
                            reader.ReadBytes(bytes);
                            using (MemoryStream ms = new MemoryStream(bytes))
                            using (Bitmap raw = new Bitmap(ms))
                            {
                                return new Bitmap(raw);
                            }
                        }
                    }
                }
            }
            catch { return null; }
        }

        private static T AwaitOp<T>(IAsyncOperation<T> op)
        {
            ManualResetEvent done = new ManualResetEvent(false);
            T res = default(T);
            Exception err = null;
            op.Completed = new AsyncOperationCompletedHandler<T>((info, status) =>
            {
                try
                {
                    if (status == AsyncStatus.Completed) res = info.GetResults();
                    else err = info.ErrorCode;
                }
                catch (Exception ex) { err = ex; }
                finally { done.Set(); }
            });
            done.WaitOne();
            if (err != null) throw err;
            return res;
        }

        private static void AwaitAction(IAsyncAction action)
        {
            ManualResetEvent done = new ManualResetEvent(false);
            Exception err = null;
            action.Completed = new AsyncActionCompletedHandler((info, status) =>
            {
                try
                {
                    if (status != AsyncStatus.Completed) err = info.ErrorCode;
                }
                catch (Exception ex) { err = ex; }
                finally { done.Set(); }
            });
            done.WaitOne();
            if (err != null) throw err;
        }

        public void Close()
        {
            doc = null;
            pageSizes.Clear();
        }
    }
}";

            var provider = new CSharpCodeProvider();
            var parameters = new CompilerParameters();
            parameters.GenerateInMemory = true;
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add("System.Drawing.dll");
            parameters.ReferencedAssemblies.Add(Path.Combine(netDir, "System.Runtime.dll"));
            parameters.ReferencedAssemblies.Add(Path.Combine(netDir, "System.Runtime.WindowsRuntime.dll"));
            parameters.ReferencedAssemblies.Add(Path.Combine(metaDir, "Windows.Foundation.winmd"));
            parameters.ReferencedAssemblies.Add(Path.Combine(metaDir, "Windows.Data.winmd"));
            parameters.ReferencedAssemblies.Add(Path.Combine(metaDir, "Windows.Storage.winmd"));

            var results = provider.CompileAssemblyFromSource(parameters, code);
            if (!results.Errors.HasErrors)
            {
                cachedWinRtType = results.CompiledAssembly.GetType("DynamicWinRt.Engine");
            }
        }
    }

    public class FallbackPdfEngine : IPdfEngine
    {
        private List<Bitmap> extractedImages = new List<Bitmap>();

        public bool Load(string pdfPath)
        {
            Close();
            try
            {
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);
                int idx = 0;
                while (idx < pdfBytes.Length - 10)
                {
                    // 查找 JPEG 文件头 0xFF, 0xD8, 0xFF
                    if (pdfBytes[idx] == 0xFF && pdfBytes[idx + 1] == 0xD8 && pdfBytes[idx + 2] == 0xFF)
                    {
                        // 查找 JPEG 文件尾 0xFF, 0xD9
                        int endIdx = -1;
                        for (int j = idx + 3; j < pdfBytes.Length - 1; j++)
                        {
                            if (pdfBytes[j] == 0xFF && pdfBytes[j + 1] == 0xD9)
                            {
                                endIdx = j + 2;
                                break;
                            }
                        }

                        if (endIdx > idx)
                        {
                            try
                            {
                                int len = endIdx - idx;
                                byte[] imgData = new byte[len];
                                Array.Copy(pdfBytes, idx, imgData, 0, len);
                                using (MemoryStream ms = new MemoryStream(imgData))
                                {
                                    Bitmap b = new Bitmap(ms);
                                    if (b.Width >= 100 && b.Height >= 100)
                                    {
                                        extractedImages.Add(new Bitmap(b));
                                    }
                                }
                            }
                            catch { }
                            idx = endIdx;
                            continue;
                        }
                    }
                    idx++;
                }
                return extractedImages.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        public int PageCount
        {
            get { return extractedImages.Count; }
        }

        public SizeF GetPageSize(int pageIndex)
        {
            if (pageIndex >= 0 && pageIndex < extractedImages.Count)
                return new SizeF(extractedImages[pageIndex].Width, extractedImages[pageIndex].Height);
            return SizeF.Empty;
        }

        public Bitmap RenderPage(int pageIndex, float scale)
        {
            if (pageIndex >= 0 && pageIndex < extractedImages.Count)
            {
                Bitmap src = extractedImages[pageIndex];
                if (Math.Abs(scale - 1.0f) < 0.05f)
                {
                    return new Bitmap(src);
                }

                int targetW = Math.Max(1, (int)Math.Round(src.Width * scale));
                int targetH = Math.Max(1, (int)Math.Round(src.Height * scale));
                Bitmap res = new Bitmap(targetW, targetH, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(res))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(src, 0, 0, targetW, targetH);
                }
                return res;
            }
            return null;
        }

        public void Close()
        {
            foreach (var b in extractedImages)
            {
                b.Dispose();
            }
            extractedImages.Clear();
        }

        public void Dispose()
        {
            Close();
        }
    }

    #endregion

    public class HelpForm : Form
    {
        private const string GITHUB_URL = "https://github.com/786381743syq/FinancePdfTool";
        private TabControl helpTabControl;

        public HelpForm(int initialTabIndex = 0)
        {
            InitializeComponent();
            if (initialTabIndex >= 0 && initialTabIndex < helpTabControl.TabPages.Count)
            {
                helpTabControl.SelectedIndex = initialTabIndex;
            }
        }

        private void InitializeComponent()
        {
            this.Text = "财务专用 PDF 转换器 - 使用须知与功能指南";
            this.Size = new Size(800, 720);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular);

            // 1. 顶部横幅
            Panel pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                BackColor = Color.FromArgb(15, 23, 42)
            };

            Label lblTitle = new Label
            {
                Text = "📖 财务专用 PDF 转换器 · 使用须知与功能指南",
                Font = new Font("Microsoft YaHei UI", 12.5F, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 14),
                AutoSize = true
            };

            Label lblSub = new Label
            {
                Text = "专为财务发票、银行单据、报销凭证与表格文档定制的高清互转与智能排版工具",
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(148, 163, 184),
                Location = new Point(22, 42),
                AutoSize = true
            };

            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblSub });
            this.Controls.Add(pnlHeader);

            // 2. 底部栏 (GitHub 地址与关闭按钮)
            Panel pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                BackColor = Color.FromArgb(248, 250, 252)
            };

            Panel divFooter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = Color.FromArgb(226, 232, 240)
            };
            pnlFooter.Controls.Add(divFooter);

            Label lblGhTitle = new Label
            {
                Text = "项目开源主页 (GitHub)：",
                Location = new Point(18, 12),
                AutoSize = true,
                ForeColor = Color.FromArgb(100, 116, 139),
                Font = new Font("Microsoft YaHei UI", 8.5F)
            };

            LinkLabel lnkGitHub = new LinkLabel
            {
                Text = GITHUB_URL,
                Location = new Point(18, 33),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                LinkColor = Color.FromArgb(37, 99, 235),
                Cursor = Cursors.Hand
            };
            lnkGitHub.LinkClicked += delegate {
                try { System.Diagnostics.Process.Start(GITHUB_URL); } catch { }
            };

            Button btnCopyGh = new Button
            {
                Text = "📋 复制链接",
                Location = new Point(460, 27),
                Size = new Size(95, 26),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(51, 65, 85),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 8.5F),
                Cursor = Cursors.Hand
            };
            btnCopyGh.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnCopyGh.Click += delegate {
                try
                {
                    Clipboard.SetText(GITHUB_URL);
                    MessageBox.Show("GitHub 项目链接已复制到剪贴板！\n\n" + GITHUB_URL, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("复制失败: " + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            Button btnClose = new Button
            {
                Text = "我知道了",
                Location = new Point(665, 14),
                Size = new Size(105, 36),
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += delegate { this.Close(); };

            pnlFooter.Controls.AddRange(new Control[] { lblGhTitle, lnkGitHub, btnCopyGh, btnClose });
            this.Controls.Add(pnlFooter);

            // 3. 中间 TabControl
            helpTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                ItemSize = new Size(230, 36),
                SizeMode = TabSizeMode.Fixed
            };

            TabPage tab1 = new TabPage("📄 图片合成 PDF · 使用须知");
            tab1.BackColor = Color.White;
            tab1.Font = new Font("Microsoft YaHei UI", 9.5F);

            Panel pnlTab1 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            TextBox txtTab1 = CreateHelpTextBox(GetTab1HelpText());
            pnlTab1.Controls.Add(txtTab1);
            tab1.Controls.Add(pnlTab1);

            TabPage tab2 = new TabPage("🖼️ PDF 提取图片 · 使用须知");
            tab2.BackColor = Color.White;
            tab2.Font = new Font("Microsoft YaHei UI", 9.5F);

            Panel pnlTab2 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            TextBox txtTab2 = CreateHelpTextBox(GetTab2HelpText());
            pnlTab2.Controls.Add(txtTab2);
            tab2.Controls.Add(pnlTab2);

            helpTabControl.TabPages.AddRange(new TabPage[] { tab1, tab2 });
            this.Controls.Add(helpTabControl);

            helpTabControl.BringToFront();
        }

        private TextBox CreateHelpTextBox(string content)
        {
            return new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(30, 41, 59),
                BorderStyle = BorderStyle.None,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular),
                Text = content.Replace("\n", "\r\n")
            };
        }

        private string GetTab1HelpText()
        {
            return 
@"【一、图片导入与页面排序】
1. 支持格式：PNG、JPG、JPEG、BMP、GIF、TIFF 等财务单据与常见图像格式。
2. 批量导入：支持点击「添加图片」多选，或点击「添加文件夹」全量载入，更支持直接从桌面或文件夹拖拽文件到窗口。
3. 智能自然排序：自动识别文件名中的连续数字编号（如 1、2、10...），彻底杜绝系统默认字典序导致的乱序错位。
4. 顺序与朝向校准：点击列表中某一页，可通过「上移」、「下移」微调顺序；点击「旋转90°」可纠正手机拍摄横竖颠倒。

【二、纸张版式与消除白边指南（核心重点）】
1. ⭐【适合原图 (无白边)】（财务单据首选）：
   • 生成的 PDF 页面尺寸 100% 匹配单据图片的实际长宽比，四周完全零白边！
   • 适用场景：银行客户交易回单、电子发票、出租车票、小票据等无需固定 A4 纸张的数字化电子归档。
2. 📄【标准 A4 / A3】（正式文件打印推荐）：
   • 生成符合国际标准的 A4（210×297mm）或 A3 尺寸标准纸张，适合正式报送与装订打印。
3. 🧭【智能感应 (自动横竖)】：
   • 当单据是横向时（如银行电子回单、横版发票），系统会自动将 A4 纸张转为横向，并横向满幅放大居中铺满，避免在竖版 A4 纸上下留下大面积空白！
4. 🔒【保持比例防拉伸】与【预留打印边距】：
   • 始终按原图比例缩放，坚决防止财务公章与文字被压扁或拉长；
   • 默认预留 30pt 安全边距，防止线下打印机物理边缘切字。

【三、画质与压缩模式】
1. ⭐【政务申报模式】（日常推荐）：
   • 体积平均缩减约 70%，针对财务红章、发票代码号码与金额进行高保真边缘增强；
   • 确保印章极清、票号锐利的同时极大压缩体积，轻松满足各类政务/税务/银行系统单文件上传大小限制（如 <10MB）。
2. 💎【原画无损模式】：
   • 100% 原始图像数据直接封装进 PDF 流，零重编码、零画质损失，适合重要审计凭证长期封存。

【四、输出与快捷操作】
• 转换完成自动弹出成功提示，支持一键点击「打开文件所在文件夹」直达归档目录。

======================================================================
🔗 开源项目主页 (GitHub)：https://github.com/786381743syq/FinancePdfTool
欢迎 Star 支持与提交反馈！";
        }

        private string GetTab2HelpText()
        {
            return 
@"【一、原生硬件加速与极速解析】
1. Windows 原生底层：基于 Windows 10/11 系统内置 Direct2D / WinRT 原生硬件加速引擎，无需额外安装任何第三方组件，秒开秒转。
2. 拖拽加载：支持直接将任意 PDF 文件拖拽进窗口即可瞬间载入。
3. 页面管理与大图预览：左侧清晰展示每页尺寸、版式、规格，右侧提供高清纸张拟真大图与翻页预览。

【二、导出格式选择】
1. ⭐【PNG 高清无损】（财务票据推荐）：
   • 绝对无损画质，发票防伪底纹、公章鲜红印泥微观细节、表格细线毫无模糊与压缩杂斑；
   • 财务入账凭证留存、发票报销核对及二次打印的推荐格式。
2. 📁【JPG 通用压缩】：
   • 高保真通用压缩算法，文件体积小巧轻量；
   • 适合用于微信、钉钉快速发送传输，或作为日常邮件附件查阅。

【三、✂️ 智能去白边（自动裁切四周多余空白留白）】
1. 痛点解决：财务客户交易回单、银行电子回单、发票等常常放在标准 A4 页面上，上下左右留下大面积多余白纸。
2. 毫秒级内存扫描：勾选「✂️ 智能去白边」后，系统毫秒级识别单据正文边界，自动剔除四周多余留白，仅提取核心单据主体！
3. 动态舒适留白：裁切时自动保留约 1.5% 自然呼吸安全边距，绝不会切到文字或印章边框；
4. 实时联动预览：勾选或取消该选项时，右侧大图预览即时同步渲染去白边效果，所见即所得。

【四、分辨率 (DPI) 精度选择】
1. 🖨️【300 DPI 超清打印】（默认推荐）：
   • 3x 超高采样率（A4 画幅约 2480×3508 像素），印章与小号字体极度锐利，满足专业印刷与法律存证标准。
2. 💻【150 DPI 高清阅读】：
   • 1.5x 高精采样（A4 画幅约 1240×1754 像素），在清晰度与文件大小之间取得最佳平衡，适合电脑屏幕浏览与归档。
3. ⚡【96 DPI 标准轻量】：
   • 1x 原生屏幕标准尺寸，体积最小，适合极速导出。

【五、导出范围与灵活筛选】
1. 导出全部页面：一键将整本 PDF 的全部页面批量转换为单张图片。
2. 仅导出列表勾选页：通过左侧勾选框，配合「全选 / 全不选 / 反选」快捷功能，自由提取指定页码。
3. 仅导出当前预览页：适合只想单独保存当前在右侧大图查看的那一页。

【六、自动归档与一键查看】
• 默认在 PDF 同级目录下自动建立「<文件名>_图片」专属文件夹，文件按「页码_01.png」自动规范命名；
• 导出完毕后自动提示，并可一键打开输出文件夹。

======================================================================
🔗 开源项目主页 (GitHub)：https://github.com/786381743syq/FinancePdfTool
欢迎 Star 支持与提交反馈！";
        }
    }
}
