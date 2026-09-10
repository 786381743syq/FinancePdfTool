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
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Drawing.Text;

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

        // ================== Tab 3：PDF 提取 Excel 相关控件 ==================
        private TabPage tabPagePdfToExcel;
        private string currentPdfForExcelPath = null;
        private List<TableSheet> currentExtractedSheets = new List<TableSheet>();
        private Label lblPdfExcelFileInfo;
        private ComboBox cmbExcelSheets;
        private RadioButton rbExportXlsx;
        private RadioButton rbExportCsv;
        private CheckBox chkFinancialFormat;
        private TextBox txtExcelOutputDir;
        private Button btnBrowseExcelOutputDir;
        private Button btnExportExcel;
        private ProgressBar progressBarExcel;
        private Label lblExcelStatus;
        private DataGridView dgvPdfExcel;
        private Label lblDgvHeader;
        private BackgroundWorker workerPdfToExcel;

        // ================== Tab 4：Excel 转成 PDF 相关控件 ==================
        private TabPage tabPageExcelToPdf;
        private string currentExcelPath = null;
        private List<ExcelSheetData> currentLoadedExcelSheets = new List<ExcelSheetData>();
        private List<Bitmap> currentExcelPdfPageBitmaps = new List<Bitmap>();
        private int currentExcelPdfPreviewPageIndex = 0;
        private Label lblExcelPdfFileInfo;
        private ComboBox cmbExcelToPdfSheets;
        private ComboBox cmbExcelPdfOrientation;
        private ComboBox cmbExcelPdfTheme;
        private CheckBox chkExcelPdfPageNum;
        private TextBox txtExcelPdfOutputDir;
        private Button btnBrowseExcelPdfOutputDir;
        private Button btnExportExcelPdf;
        private ProgressBar progressBarExcelPdf;
        private Label lblExcelPdfStatus;
        private PictureBox previewBoxExcelPdf;
        private Label lblExcelPdfPageInfo;
        private Button btnPrevExcelPdfPage;
        private Button btnNextExcelPdfPage;
        private BackgroundWorker workerExcelToPdf;

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

            tabPagePdfToExcel = new TabPage("📊 PDF 提取 Excel")
            {
                BackColor = Color.FromArgb(245, 247, 250),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular)
            };

            tabPageExcelToPdf = new TabPage("📑 Excel 转成 PDF")
            {
                BackColor = Color.FromArgb(245, 247, 250),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular)
            };

            BuildTabImgToPdf();
            BuildTabPdfToImg();
            BuildTabPdfToExcel();
            BuildTabExcelToPdf();

            mainTabControl.TabPages.Add(tabPageImgToPdf);
            mainTabControl.TabPages.Add(tabPagePdfToImg);
            mainTabControl.TabPages.Add(tabPagePdfToExcel);
            mainTabControl.TabPages.Add(tabPageExcelToPdf);
            mainTabControl.ItemSize = new Size(220, 38);
            this.Controls.Add(mainTabControl);

            // 后台任务工作者 (Tab 3 & 4)
            workerPdfToExcel = new BackgroundWorker { WorkerReportsProgress = true };
            workerPdfToExcel.DoWork += WorkerPdfToExcel_DoWork;
            workerPdfToExcel.ProgressChanged += WorkerPdfToExcel_ProgressChanged;
            workerPdfToExcel.RunWorkerCompleted += WorkerPdfToExcel_RunWorkerCompleted;

            workerExcelToPdf = new BackgroundWorker { WorkerReportsProgress = true };
            workerExcelToPdf.DoWork += WorkerExcelToPdf_DoWork;
            workerExcelToPdf.ProgressChanged += WorkerExcelToPdf_ProgressChanged;
            workerExcelToPdf.RunWorkerCompleted += WorkerExcelToPdf_RunWorkerCompleted;

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
            string firstExcel = null;
            foreach (string p in paths)
            {
                if (File.Exists(p))
                {
                    string ext = Path.GetExtension(p).ToLowerInvariant();
                    if (ext == ".pdf" && firstPdf == null) firstPdf = p;
                    if ((ext == ".xlsx" || ext == ".xls" || ext == ".csv") && firstExcel == null) firstExcel = p;
                }
            }

            if (firstExcel != null)
            {
                mainTabControl.SelectedTab = tabPageExcelToPdf;
                LoadExcelDocument(firstExcel);
            }
            else if (firstPdf != null)
            {
                if (mainTabControl.SelectedTab == tabPagePdfToExcel)
                {
                    LoadPdfForExcel(firstPdf);
                }
                else
                {
                    mainTabControl.SelectedTab = tabPagePdfToImg;
                    LoadPdfDocument(firstPdf);
                }
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

        #region ================== Tab 3：PDF 提取 Excel UI 与逻辑 ==================

        private void BuildTabPdfToExcel()
        {
            // 1. 顶部工具栏
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 54),
                BackColor = Color.White
            };

            Button btnOpenPdf = CreateButton("📂 选择 PDF 文件...", 140, 34, Color.FromArgb(37, 99, 235), Color.White);
            btnOpenPdf.Location = new Point(15, 10);
            btnOpenPdf.Click += delegate {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "选择需要提取表格的 PDF 文件";
                    ofd.Filter = "PDF 电子文档 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        LoadPdfForExcel(ofd.FileName);
                    }
                }
            };

            lblPdfExcelFileInfo = new Label
            {
                Text = "请点击左侧按钮或直接拖拽 PDF 文件（支持年报/季报财报、申报表、对账单等）至此窗口",
                Location = new Point(165, 17),
                Size = new Size(720, 22),
                ForeColor = Color.FromArgb(71, 85, 105),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular),
                AutoEllipsis = true
            };

            Button btnHelpTab3 = CreateButton("💡 使用须知", 100, 34, Color.FromArgb(243, 244, 246), Color.FromArgb(79, 70, 229));
            btnHelpTab3.Location = new Point(895, 10);
            btnHelpTab3.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            btnHelpTab3.Click += delegate { ShowHelpDialog(2); };

            topBar.Controls.AddRange(new Control[] { btnOpenPdf, lblPdfExcelFileInfo, btnHelpTab3 });
            tabPagePdfToExcel.Controls.Add(topBar);

            // 2. 左侧控制面板
            Panel pnlLeft = new Panel
            {
                Location = new Point(12, 64),
                Size = new Size(420, 626),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 2.1 报表识别与工作表选择
            GroupBox grpSheets = new GroupBox
            {
                Text = "报表识别与工作表选择",
                Location = new Point(14, 12),
                Size = new Size(390, 105),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblSheetSelect = new Label
            {
                Text = "选择预览/导出的工作表：",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            cmbExcelSheets = new ComboBox
            {
                Location = new Point(18, 52),
                Size = new Size(355, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbExcelSheets.Items.Add("尚未载入任何 PDF 报表");
            cmbExcelSheets.SelectedIndex = 0;
            cmbExcelSheets.SelectedIndexChanged += CmbExcelSheets_SelectedIndexChanged;

            grpSheets.Controls.AddRange(new Control[] { lblSheetSelect, cmbExcelSheets });
            pnlLeft.Controls.Add(grpSheets);

            // 2.2 导出格式与数据规则
            GroupBox grpRules = new GroupBox
            {
                Text = "导出格式与数据规则",
                Location = new Point(14, 126),
                Size = new Size(390, 135),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            rbExportXlsx = new RadioButton
            {
                Text = "Excel 工作簿 (*.xlsx) - 现代化 OpenXML 推荐",
                Location = new Point(18, 25),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            rbExportCsv = new RadioButton
            {
                Text = "CSV 表格 (*.csv) - 兼容通用系统",
                Location = new Point(18, 55),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 23, 42)
            };

            chkFinancialFormat = new CheckBox
            {
                Text = "金额写入真实数值 (保留千分位，直接支持SUM求和)",
                Location = new Point(18, 90),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(79, 70, 229),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold)
            };

            grpRules.Controls.AddRange(new Control[] { rbExportXlsx, rbExportCsv, chkFinancialFormat });
            pnlLeft.Controls.Add(grpRules);

            // 2.3 输出路径与生成
            GroupBox grpOut = new GroupBox
            {
                Text = "导出保存路径",
                Location = new Point(14, 270),
                Size = new Size(390, 160),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblPath = new Label
            {
                Text = "保存路径：",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            txtExcelOutputDir = new TextBox
            {
                Location = new Point(18, 48),
                Size = new Size(275, 25),
                ReadOnly = false
            };

            btnBrowseExcelOutputDir = CreateButton("浏览...", 70, 27, Color.FromArgb(241, 245, 249), Color.FromArgb(51, 65, 85));
            btnBrowseExcelOutputDir.Location = new Point(300, 47);
            btnBrowseExcelOutputDir.Click += BtnBrowseExcelOutputDir_Click;

            btnExportExcel = CreateButton("⚡ 立即导出 Excel 表格", 355, 42, Color.FromArgb(16, 185, 129), Color.White);
            btnExportExcel.Location = new Point(18, 92);
            btnExportExcel.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            btnExportExcel.Click += BtnExportExcel_Click;

            grpOut.Controls.AddRange(new Control[] { lblPath, txtExcelOutputDir, btnBrowseExcelOutputDir, btnExportExcel });
            pnlLeft.Controls.Add(grpOut);

            // 进度条与状态
            progressBarExcel = new ProgressBar
            {
                Location = new Point(14, 550),
                Size = new Size(390, 14),
                Style = ProgressBarStyle.Continuous
            };

            lblExcelStatus = new Label
            {
                Text = "就绪。请载入 PDF 开始提取表格。",
                Location = new Point(14, 574),
                Size = new Size(390, 36),
                ForeColor = Color.FromArgb(100, 116, 139),
                Font = new Font("Microsoft YaHei UI", 9F)
            };

            pnlLeft.Controls.AddRange(new Control[] { progressBarExcel, lblExcelStatus });
            tabPagePdfToExcel.Controls.Add(pnlLeft);

            // 3. 右侧实时表格预览面板
            Panel pnlRight = new Panel
            {
                Location = new Point(444, 64),
                Size = new Size(556, 626),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            lblDgvHeader = new Label
            {
                Text = "📋 数据实时表格预览 (所见即所得)",
                Location = new Point(12, 10),
                Size = new Size(530, 24),
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            dgvPdfExcel = new DataGridView
            {
                Location = new Point(12, 40),
                Size = new Size(530, 570),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                GridColor = Color.FromArgb(226, 232, 240),
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            };
            dgvPdfExcel.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 250);
            dgvPdfExcel.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59);
            dgvPdfExcel.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
            dgvPdfExcel.ColumnHeadersHeight = 32;
            dgvPdfExcel.EnableHeadersVisualStyles = false;
            dgvPdfExcel.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 254);
            dgvPdfExcel.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular);
            dgvPdfExcel.DefaultCellStyle.ForeColor = Color.FromArgb(51, 65, 85);
            dgvPdfExcel.RowTemplate.Height = 28;

            // Enable DoubleBuffered on DataGridView via reflection
            try
            {
                typeof(DataGridView).InvokeMember("DoubleBuffered",
                    BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                    null, dgvPdfExcel, new object[] { true });
            }
            catch { }

            pnlRight.Controls.AddRange(new Control[] { lblDgvHeader, dgvPdfExcel });
            tabPagePdfToExcel.Controls.Add(pnlRight);
        }

        public void LoadPdfForExcel(string filePath)
        {
            if (!File.Exists(filePath)) return;
            currentPdfForExcelPath = filePath;
            lblPdfExcelFileInfo.Text = "已载入: " + Path.GetFileName(filePath) + " (正在高精解析报表结构...)";
            progressBarExcel.Value = 30;

            try
            {
                currentExtractedSheets = PdfTableExtractor.ExtractWorkbook(filePath);
                progressBarExcel.Value = 100;

                cmbExcelSheets.Items.Clear();
                if (currentExtractedSheets.Count > 1)
                {
                    cmbExcelSheets.Items.Add("⭐ [全部工作表] (多Sheet合并导出 - 共" + currentExtractedSheets.Count + "个报表)");
                }
                for (int i = 0; i < currentExtractedSheets.Count; i++)
                {
                    var s = currentExtractedSheets[i];
                    cmbExcelSheets.Items.Add(string.Format("{0}. {1} ({2}行 × {3}列)", i + 1, s.Title, s.Rows.Count, s.Headers.Count));
                }

                if (cmbExcelSheets.Items.Count > 0) cmbExcelSheets.SelectedIndex = (currentExtractedSheets.Count > 1) ? 1 : 0;

                string defaultOut = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".xlsx");
                txtExcelOutputDir.Text = defaultOut;

                int totalRows = 0;
                foreach (var s in currentExtractedSheets) totalRows += s.Rows.Count;
                lblPdfExcelFileInfo.Text = string.Format("已载入: {0} (共 {1} 页，检测到 {2} 个财务报表，合计 {3} 行)", Path.GetFileName(filePath), currentExtractedSheets.Count, currentExtractedSheets.Count, totalRows);
                lblExcelStatus.Text = "✅ 解析完成！可点击「立即导出 Excel 表格」或在右侧预览。";
            }
            catch (Exception ex)
            {
                progressBarExcel.Value = 0;
                lblPdfExcelFileInfo.Text = "解析失败: " + ex.Message;
                lblExcelStatus.Text = "❌ 解析出错: " + ex.Message;
                MessageBox.Show("解析 PDF 报表结构时出错：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CmbExcelSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (currentExtractedSheets == null || currentExtractedSheets.Count == 0) return;
            int idx = cmbExcelSheets.SelectedIndex;
            if (currentExtractedSheets.Count > 1)
            {
                if (idx == 0) idx = 1; // display first sheet if "all" selected
                int sheetIdx = idx - 1;
                if (sheetIdx >= 0 && sheetIdx < currentExtractedSheets.Count)
                {
                    DisplaySheetInGrid(currentExtractedSheets[sheetIdx]);
                }
            }
            else if (idx >= 0 && idx < currentExtractedSheets.Count)
            {
                DisplaySheetInGrid(currentExtractedSheets[idx]);
            }
        }

        private void DisplaySheetInGrid(TableSheet sheet)
        {
            dgvPdfExcel.Columns.Clear();
            dgvPdfExcel.Rows.Clear();

            if (sheet == null || sheet.Headers.Count == 0) return;

            if (lblDgvHeader != null)
            {
                if (sheet.HeaderInfo != null && !string.IsNullOrEmpty(sheet.HeaderInfo.Title))
                {
                    lblDgvHeader.Text = "📋 " + sheet.HeaderInfo.Title + (!string.IsNullOrEmpty(sheet.HeaderInfo.CompanyName) ? ("  [" + sheet.HeaderInfo.CompanyName + "]") : "");
                }
                else
                {
                    lblDgvHeader.Text = "📋 " + sheet.Title + " (所见即所得表格预览)";
                }
            }

            for (int c = 0; c < sheet.Headers.Count; c++)
            {
                string hName = sheet.Headers[c];
                var col = new DataGridViewTextBoxColumn
                {
                    HeaderText = hName,
                    Width = (hName.Contains("资产") || hName.Contains("负债") || hName.Contains("项目")) ? 140 :
                            (hName.Contains("金额") || hName.Contains("余额")) ? 115 : 60
                };
                if (hName.Contains("金额") || hName.Contains("余额") || hName.Contains("行次"))
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }
                dgvPdfExcel.Columns.Add(col);
            }

            foreach (var r in sheet.Rows)
            {
                object[] vals = new object[sheet.Headers.Count];
                for (int c = 0; c < sheet.Headers.Count; c++)
                {
                    vals[c] = (c < r.Count) ? r[c].Text : "";
                }
                dgvPdfExcel.Rows.Add(vals);
            }
        }

        private void BtnBrowseExcelOutputDir_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "选择 Excel 保存路径";
                sfd.Filter = rbExportXlsx.Checked ? "Excel 工作簿 (*.xlsx)|*.xlsx" : "CSV 逗号分隔表格 (*.csv)|*.csv";
                if (!string.IsNullOrEmpty(txtExcelOutputDir.Text))
                {
                    try
                    {
                        sfd.InitialDirectory = Path.GetDirectoryName(txtExcelOutputDir.Text);
                        sfd.FileName = Path.GetFileName(txtExcelOutputDir.Text);
                    }
                    catch { }
                }
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtExcelOutputDir.Text = sfd.FileName;
                }
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (currentExtractedSheets == null || currentExtractedSheets.Count == 0)
            {
                MessageBox.Show("请先选择并解析 PDF 报表文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string outPath = txtExcelOutputDir.Text.Trim();
            if (string.IsNullOrEmpty(outPath))
            {
                MessageBox.Show("请指定导出文件的保存路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnExportExcel.Enabled = false;
            progressBarExcel.Value = 20;
            lblExcelStatus.Text = "正在生成 Excel 工作簿...";

            bool isCsv = rbExportCsv.Checked;
            int selectedSheetIdx = cmbExcelSheets.SelectedIndex;

            workerPdfToExcel.RunWorkerAsync(new object[] { outPath, isCsv, selectedSheetIdx });
        }

        private void WorkerPdfToExcel_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] args = (object[])e.Argument;
            string outPath = (string)args[0];
            bool isCsv = (bool)args[1];
            int sheetSelectIdx = (int)args[2];

            List<TableSheet> exportSheets = new List<TableSheet>();
            if (currentExtractedSheets.Count > 1)
            {
                if (sheetSelectIdx == 0)
                {
                    exportSheets.AddRange(currentExtractedSheets);
                }
                else
                {
                    int target = sheetSelectIdx - 1;
                    if (target >= 0 && target < currentExtractedSheets.Count)
                    {
                        exportSheets.Add(currentExtractedSheets[target]);
                    }
                }
            }
            else
            {
                exportSheets.AddRange(currentExtractedSheets);
            }

            workerPdfToExcel.ReportProgress(50, "正在组装 OpenXML 工作簿包...");

            if (isCsv)
            {
                StringBuilder sbCsv = new StringBuilder();
                foreach (var s in exportSheets)
                {
                    sbCsv.AppendLine(string.Join(",", s.Headers.ToArray()));
                    foreach (var r in s.Rows)
                    {
                        var rowItems = new List<string>();
                        foreach (var cell in r) rowItems.Add("\"" + cell.Text.Replace("\"", "\"\"") + "\"");
                        sbCsv.AppendLine(string.Join(",", rowItems.ToArray()));
                    }
                }
                File.WriteAllText(outPath, sbCsv.ToString(), Encoding.UTF8);
            }
            else
            {
                byte[] xlsxBytes = ExcelBuilder.GenerateXlsx(exportSheets);
                File.WriteAllBytes(outPath, xlsxBytes);
            }

            workerPdfToExcel.ReportProgress(100, "导出完成！");
            e.Result = outPath;
        }

        private void WorkerPdfToExcel_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarExcel.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblExcelStatus.Text = (string)e.UserState;
        }

        private void WorkerPdfToExcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExportExcel.Enabled = true;
            progressBarExcel.Value = 100;

            if (e.Error != null)
            {
                lblExcelStatus.Text = "❌ 导出失败: " + e.Error.Message;
                MessageBox.Show("导出 Excel 时发生错误：\n" + e.Error.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string outPath = (string)e.Result;
            lblExcelStatus.Text = "✅ 成功导出 Excel: " + outPath;

            DialogResult dr = MessageBox.Show("Excel 表格导出成功！\n\n保存路径：" + outPath + "\n\n是否立即打开导出的 Excel 文件？", "导出完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                try { System.Diagnostics.Process.Start(outPath); } catch { }
            }
        }

        #endregion

        #region ================== Tab 4：Excel 转成 PDF UI 与逻辑 ==================

        private void BuildTabExcelToPdf()
        {
            // 1. 顶部工具栏
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(1012, 54),
                BackColor = Color.White
            };

            Button btnOpenExcel = CreateButton("📂 选择 Excel 文件...", 155, 34, Color.FromArgb(99, 102, 241), Color.White);
            btnOpenExcel.Location = new Point(15, 10);
            btnOpenExcel.Click += delegate {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "选择需要转换为 PDF 的 Excel 文件";
                    ofd.Filter = "Excel 与 表格文件 (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|所有文件 (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        LoadExcelDocument(ofd.FileName);
                    }
                }
            };

            lblExcelPdfFileInfo = new Label
            {
                Text = "请点击左侧按钮或直接拖拽 Excel (*.xlsx, *.xls, *.csv) 文件至此窗口",
                Location = new Point(180, 17),
                Size = new Size(705, 22),
                ForeColor = Color.FromArgb(71, 85, 105),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular),
                AutoEllipsis = true
            };

            Button btnHelpTab4 = CreateButton("💡 使用须知", 100, 34, Color.FromArgb(243, 244, 246), Color.FromArgb(79, 70, 229));
            btnHelpTab4.Location = new Point(895, 10);
            btnHelpTab4.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            btnHelpTab4.Click += delegate { ShowHelpDialog(3); };

            topBar.Controls.AddRange(new Control[] { btnOpenExcel, lblExcelPdfFileInfo, btnHelpTab4 });
            tabPageExcelToPdf.Controls.Add(topBar);

            // 2. 左侧控制面板
            Panel pnlLeft = new Panel
            {
                Location = new Point(12, 64),
                Size = new Size(420, 626),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 2.1 工作表选择
            GroupBox grpSheets = new GroupBox
            {
                Text = "工作表选择",
                Location = new Point(14, 12),
                Size = new Size(390, 105),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblSelect = new Label
            {
                Text = "选择要转换的工作表：",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = Color.FromArgb(51, 65, 85)
            };

            cmbExcelToPdfSheets = new ComboBox
            {
                Location = new Point(18, 52),
                Size = new Size(355, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbExcelToPdfSheets.Items.Add("尚未载入任何 Excel 工作簿");
            cmbExcelToPdfSheets.SelectedIndex = 0;
            cmbExcelToPdfSheets.SelectedIndexChanged += delegate { RenderExcelPdfPreview(); };

            grpSheets.Controls.AddRange(new Control[] { lblSelect, cmbExcelToPdfSheets });
            pnlLeft.Controls.Add(grpSheets);

            // 2.2 版式与主题
            GroupBox grpLayout = new GroupBox
            {
                Text = "页面版式与风格",
                Location = new Point(14, 126),
                Size = new Size(390, 150),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblOri = new Label { Text = "页面方向：", Location = new Point(15, 26), AutoSize = true, ForeColor = Color.FromArgb(51, 65, 85) };
            cmbExcelPdfOrientation = new ComboBox { Location = new Point(90, 23), Size = new Size(283, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbExcelPdfOrientation.Items.AddRange(new object[] {
                "🧭 智能感应 (列数>5自动横向，否则纵向 · 推荐)",
                "横向 (Landscape - 宽表格推荐)",
                "纵向 (Portrait - 窄表格推荐)"
            });
            cmbExcelPdfOrientation.SelectedIndex = 0;
            cmbExcelPdfOrientation.SelectedIndexChanged += delegate { RenderExcelPdfPreview(); };

            Label lblTheme = new Label { Text = "表格主题：", Location = new Point(15, 66), AutoSize = true, ForeColor = Color.FromArgb(51, 65, 85) };
            cmbExcelPdfTheme = new ComboBox { Location = new Point(90, 63), Size = new Size(283, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbExcelPdfTheme.Items.AddRange(new object[] {
                "经典财务蓝 (精致表头 + 舒适条纹 · 推荐)",
                "极简黑白网格 (正式红头文件/印刷推荐)",
                "现代商务灰 (淡雅沉稳)"
            });
            cmbExcelPdfTheme.SelectedIndex = 0;
            cmbExcelPdfTheme.SelectedIndexChanged += delegate { RenderExcelPdfPreview(); };

            chkExcelPdfPageNum = new CheckBox
            {
                Text = "页脚打印页码与工作表名称 (如: 工作表 第 1 / 3 页)",
                Location = new Point(18, 108),
                AutoSize = true,
                Checked = true,
                ForeColor = Color.FromArgb(79, 70, 229),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold)
            };

            grpLayout.Controls.AddRange(new Control[] { lblOri, cmbExcelPdfOrientation, lblTheme, cmbExcelPdfTheme, chkExcelPdfPageNum });
            pnlLeft.Controls.Add(grpLayout);

            // 2.3 输出路径与生成
            GroupBox grpOut = new GroupBox
            {
                Text = "PDF 导出保存路径",
                Location = new Point(14, 285),
                Size = new Size(390, 160),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            Label lblPdfPath = new Label { Text = "保存路径：", Location = new Point(15, 25), AutoSize = true, ForeColor = Color.FromArgb(51, 65, 85) };
            txtExcelPdfOutputDir = new TextBox { Location = new Point(18, 48), Size = new Size(275, 25) };
            btnBrowseExcelPdfOutputDir = CreateButton("浏览...", 70, 27, Color.FromArgb(241, 245, 249), Color.FromArgb(51, 65, 85));
            btnBrowseExcelPdfOutputDir.Location = new Point(300, 47);
            btnBrowseExcelPdfOutputDir.Click += BtnBrowseExcelPdfOutputDir_Click;

            btnExportExcelPdf = CreateButton("⚡ 立即生成 PDF 文件", 355, 42, Color.FromArgb(99, 102, 241), Color.White);
            btnExportExcelPdf.Location = new Point(18, 92);
            btnExportExcelPdf.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            btnExportExcelPdf.Click += BtnExportExcelPdf_Click;

            grpOut.Controls.AddRange(new Control[] { lblPdfPath, txtExcelPdfOutputDir, btnBrowseExcelPdfOutputDir, btnExportExcelPdf });
            pnlLeft.Controls.Add(grpOut);

            // 进度条与状态
            progressBarExcelPdf = new ProgressBar
            {
                Location = new Point(14, 550),
                Size = new Size(390, 14),
                Style = ProgressBarStyle.Continuous
            };

            lblExcelPdfStatus = new Label
            {
                Text = "就绪。请载入 Excel 文件开始转换。",
                Location = new Point(14, 574),
                Size = new Size(390, 36),
                ForeColor = Color.FromArgb(100, 116, 139),
                Font = new Font("Microsoft YaHei UI", 9F)
            };

            pnlLeft.Controls.AddRange(new Control[] { progressBarExcelPdf, lblExcelPdfStatus });
            tabPageExcelToPdf.Controls.Add(pnlLeft);

            // 3. 右侧 PDF 实时大图效果预览
            Panel pnlRight = new Panel
            {
                Location = new Point(444, 64),
                Size = new Size(556, 626),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lblPreviewHeader = new Label
            {
                Text = "📑 PDF 页面效果预览",
                Location = new Point(12, 10),
                Size = new Size(530, 24),
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            previewBoxExcelPdf = new PictureBox
            {
                Location = new Point(12, 40),
                Size = new Size(530, 520),
                BackColor = Color.FromArgb(238, 242, 246),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle
            };

            Panel pnlNav = new Panel
            {
                Location = new Point(12, 570),
                Size = new Size(530, 42),
                BackColor = Color.FromArgb(248, 250, 252)
            };

            btnPrevExcelPdfPage = CreateButton("◀ 上一页", 90, 30, Color.White, Color.FromArgb(51, 65, 85));
            btnPrevExcelPdfPage.Location = new Point(110, 6);
            btnPrevExcelPdfPage.Click += delegate {
                if (currentExcelPdfPreviewPageIndex > 0)
                {
                    currentExcelPdfPreviewPageIndex--;
                    UpdateExcelPdfPreviewImage();
                }
            };

            lblExcelPdfPageInfo = new Label
            {
                Text = "第 0 / 0 页",
                Location = new Point(210, 10),
                Size = new Size(110, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            btnNextExcelPdfPage = CreateButton("下一页 ▶", 90, 30, Color.White, Color.FromArgb(51, 65, 85));
            btnNextExcelPdfPage.Location = new Point(330, 6);
            btnNextExcelPdfPage.Click += delegate {
                if (currentExcelPdfPreviewPageIndex < currentExcelPdfPageBitmaps.Count - 1)
                {
                    currentExcelPdfPreviewPageIndex++;
                    UpdateExcelPdfPreviewImage();
                }
            };

            pnlNav.Controls.AddRange(new Control[] { btnPrevExcelPdfPage, lblExcelPdfPageInfo, btnNextExcelPdfPage });
            pnlRight.Controls.AddRange(new Control[] { lblPreviewHeader, previewBoxExcelPdf, pnlNav });
            tabPageExcelToPdf.Controls.Add(pnlRight);
        }

        public void LoadExcelDocument(string filePath)
        {
            if (!File.Exists(filePath)) return;
            currentExcelPath = filePath;
            lblExcelPdfFileInfo.Text = "已载入: " + Path.GetFileName(filePath) + " (正在读取工作表...)";

            try
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext == ".csv")
                {
                    currentLoadedExcelSheets = ExcelReader.ReadCsv(filePath);
                }
                else
                {
                    currentLoadedExcelSheets = ExcelReader.ReadXlsx(filePath);
                }

                cmbExcelToPdfSheets.Items.Clear();
                if (currentLoadedExcelSheets.Count > 1)
                {
                    cmbExcelToPdfSheets.Items.Add("⭐ [全部工作表合并导出] (共 " + currentLoadedExcelSheets.Count + " 个工作表)");
                }
                for (int i = 0; i < currentLoadedExcelSheets.Count; i++)
                {
                    var s = currentLoadedExcelSheets[i];
                    cmbExcelToPdfSheets.Items.Add(string.Format("{0}. {1} ({2}行)", i + 1, s.Name, s.Rows.Count));
                }
                if (cmbExcelToPdfSheets.Items.Count > 0) cmbExcelToPdfSheets.SelectedIndex = 0;

                string defaultOut = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".pdf");
                txtExcelPdfOutputDir.Text = defaultOut;

                int totalRows = 0;
                foreach (var s in currentLoadedExcelSheets) totalRows += s.Rows.Count;
                lblExcelPdfFileInfo.Text = string.Format("已载入: {0} (共 {1} 个工作表，{2} 行数据)", Path.GetFileName(filePath), currentLoadedExcelSheets.Count, totalRows);
                lblExcelPdfStatus.Text = "✅ 读取成功！右侧已就绪实时渲染预览。";

                RenderExcelPdfPreview();
            }
            catch (Exception ex)
            {
                lblExcelPdfFileInfo.Text = "读取失败: " + ex.Message;
                lblExcelPdfStatus.Text = "❌ 出错: " + ex.Message;
                MessageBox.Show("读取 Excel 文件失败：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderExcelPdfPreview()
        {
            if (currentLoadedExcelSheets == null || currentLoadedExcelSheets.Count == 0) return;

            // Dispose old bitmaps
            foreach (var b in currentExcelPdfPageBitmaps) b.Dispose();
            currentExcelPdfPageBitmaps.Clear();
            currentExcelPdfPreviewPageIndex = 0;

            List<ExcelSheetData> targetSheets = new List<ExcelSheetData>();
            int idx = cmbExcelToPdfSheets.SelectedIndex;
            if (currentLoadedExcelSheets.Count > 1)
            {
                if (idx == 0) targetSheets.AddRange(currentLoadedExcelSheets);
                else if (idx - 1 >= 0 && idx - 1 < currentLoadedExcelSheets.Count) targetSheets.Add(currentLoadedExcelSheets[idx - 1]);
            }
            else if (idx >= 0 && idx < currentLoadedExcelSheets.Count)
            {
                targetSheets.Add(currentLoadedExcelSheets[idx]);
            }

            int oriMode = cmbExcelPdfOrientation.SelectedIndex;
            int themeMode = cmbExcelPdfTheme.SelectedIndex;
            bool printFooter = chkExcelPdfPageNum.Checked;

            foreach (var s in targetSheets)
            {
                var bmps = ExcelToPdfRenderer.RenderSheetToBitmaps(s, oriMode, themeMode, printFooter);
                currentExcelPdfPageBitmaps.AddRange(bmps);
            }

            UpdateExcelPdfPreviewImage();
        }

        private void UpdateExcelPdfPreviewImage()
        {
            if (currentExcelPdfPageBitmaps.Count == 0)
            {
                previewBoxExcelPdf.Image = null;
                lblExcelPdfPageInfo.Text = "第 0 / 0 页";
                btnPrevExcelPdfPage.Enabled = false;
                btnNextExcelPdfPage.Enabled = false;
                return;
            }

            if (currentExcelPdfPreviewPageIndex < 0) currentExcelPdfPreviewPageIndex = 0;
            if (currentExcelPdfPreviewPageIndex >= currentExcelPdfPageBitmaps.Count) currentExcelPdfPreviewPageIndex = currentExcelPdfPageBitmaps.Count - 1;

            previewBoxExcelPdf.Image = currentExcelPdfPageBitmaps[currentExcelPdfPreviewPageIndex];
            lblExcelPdfPageInfo.Text = string.Format("第 {0} / {1} 页", currentExcelPdfPreviewPageIndex + 1, currentExcelPdfPageBitmaps.Count);

            btnPrevExcelPdfPage.Enabled = (currentExcelPdfPreviewPageIndex > 0);
            btnNextExcelPdfPage.Enabled = (currentExcelPdfPreviewPageIndex < currentExcelPdfPageBitmaps.Count - 1);
        }

        private void BtnBrowseExcelPdfOutputDir_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "选择生成的 PDF 保存路径";
                sfd.Filter = "PDF 电子文档 (*.pdf)|*.pdf";
                if (!string.IsNullOrEmpty(txtExcelPdfOutputDir.Text))
                {
                    try
                    {
                        sfd.InitialDirectory = Path.GetDirectoryName(txtExcelPdfOutputDir.Text);
                        sfd.FileName = Path.GetFileName(txtExcelPdfOutputDir.Text);
                    }
                    catch { }
                }
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPdfOutputDir.Text = sfd.FileName;
                }
            }
        }

        private void BtnExportExcelPdf_Click(object sender, EventArgs e)
        {
            if (currentLoadedExcelSheets == null || currentLoadedExcelSheets.Count == 0)
            {
                MessageBox.Show("请先选择并载入 Excel 文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string outPath = txtExcelPdfOutputDir.Text.Trim();
            if (string.IsNullOrEmpty(outPath))
            {
                MessageBox.Show("请指定 PDF 保存路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnExportExcelPdf.Enabled = false;
            progressBarExcelPdf.Value = 20;
            lblExcelPdfStatus.Text = "正在生成高清 PDF 文档...";

            workerExcelToPdf.RunWorkerAsync(outPath);
        }

        private void WorkerExcelToPdf_DoWork(object sender, DoWorkEventArgs e)
        {
            string outPath = (string)e.Argument;

            workerExcelToPdf.ReportProgress(40, "正在渲染各工作表矢量页面...");

            List<Bitmap> renderBitmaps = new List<Bitmap>();
            List<ExcelSheetData> targetSheets = new List<ExcelSheetData>();
            int idx = 0;
            int oriMode = 0;
            int themeMode = 0;
            bool printFooter = true;

            this.Invoke(new Action(delegate {
                idx = cmbExcelToPdfSheets.SelectedIndex;
                oriMode = cmbExcelPdfOrientation.SelectedIndex;
                themeMode = cmbExcelPdfTheme.SelectedIndex;
                printFooter = chkExcelPdfPageNum.Checked;

                if (currentLoadedExcelSheets.Count > 1)
                {
                    if (idx == 0) targetSheets.AddRange(currentLoadedExcelSheets);
                    else if (idx - 1 >= 0 && idx - 1 < currentLoadedExcelSheets.Count) targetSheets.Add(currentLoadedExcelSheets[idx - 1]);
                }
                else if (idx >= 0 && idx < currentLoadedExcelSheets.Count)
                {
                    targetSheets.Add(currentLoadedExcelSheets[idx]);
                }
            }));

            foreach (var s in targetSheets)
            {
                var bmps = ExcelToPdfRenderer.RenderSheetToBitmaps(s, oriMode, themeMode, printFooter);
                renderBitmaps.AddRange(bmps);
            }

            workerExcelToPdf.ReportProgress(80, "正在封装 PDF 文件流...");

            // Call static BuildPdf
            PdfBuilder.BuildPdfFromBitmaps(renderBitmaps, outPath);

            foreach (var b in renderBitmaps) b.Dispose();

            workerExcelToPdf.ReportProgress(100, "转换完成！");
            e.Result = outPath;
        }

        private void WorkerExcelToPdf_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarExcelPdf.Value = Math.Min(100, Math.Max(0, e.ProgressPercentage));
            lblExcelPdfStatus.Text = (string)e.UserState;
        }

        private void WorkerExcelToPdf_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExportExcelPdf.Enabled = true;
            progressBarExcelPdf.Value = 100;

            if (e.Error != null)
            {
                lblExcelPdfStatus.Text = "❌ 转换失败: " + e.Error.Message;
                MessageBox.Show("转换 PDF 发生错误：\n" + e.Error.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string outPath = (string)e.Result;
            lblExcelPdfStatus.Text = "✅ 成功生成 PDF: " + outPath;

            DialogResult dr = MessageBox.Show("PDF 文件转换生成成功！\n\n保存路径：" + outPath + "\n\n是否立即打开生成的 PDF 文件？", "转换完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                try { System.Diagnostics.Process.Start(outPath); } catch { }
            }
        }

        #endregion


        [STAThread]
        static void Main(string[] args)
        {
            if (args != null && args.Length >= 2 && args[0] == "--test-extract")
            {
                RunTestExtract(args[1]);
                return;
            }
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

        static void RunTestExtract(string pdfPath)
        {
            var sheets = PdfTableExtractor.ExtractWorkbook(pdfPath);
            Console.WriteLine("Extracted " + sheets.Count + " sheets.");
            for (int s = 0; s < sheets.Count; s++)
            {
                var sheet = sheets[s];
                Console.WriteLine(string.Format("\n--- Sheet {0}: {1} (Total Rows: {2}) ---", s + 1, sheet.Title, sheet.Rows.Count));
                int start = Math.Max(0, sheet.Rows.Count - 4);
                for (int r = start; r < sheet.Rows.Count; r++)
                {
                    var parts = new List<string>();
                    foreach (var c in sheet.Rows[r]) parts.Add("'" + c.Text + "'");
                    Console.WriteLine(string.Format("Row {0,2}: {1}", r + 1, string.Join(" | ", parts.ToArray())));
                }
            }

            byte[] xlsxBytes = ExcelBuilder.GenerateXlsx(sheets);
            string fixedOut = Path.Combine(Path.GetDirectoryName(pdfPath), Path.GetFileNameWithoutExtension(pdfPath) + "_已修复.xlsx");
            File.WriteAllBytes(fixedOut, xlsxBytes);
            Console.WriteLine("Saved: " + fixedOut);

            try
            {
                string origOut = Path.Combine(Path.GetDirectoryName(pdfPath), Path.GetFileNameWithoutExtension(pdfPath) + ".xlsx");
                File.WriteAllBytes(origOut, xlsxBytes);
                Console.WriteLine("Overwritten: " + origOut);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Original file busy: " + ex.Message);
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

            helpTabControl.ItemSize = new Size(185, 36);

            TabPage tab1 = new TabPage("📄 图片合成 PDF");
            tab1.BackColor = Color.White;
            tab1.Font = new Font("Microsoft YaHei UI", 9.5F);
            Panel pnlTab1 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            pnlTab1.Controls.Add(CreateHelpTextBox(GetTab1HelpText()));
            tab1.Controls.Add(pnlTab1);

            TabPage tab2 = new TabPage("🖼️ PDF 提取图片");
            tab2.BackColor = Color.White;
            tab2.Font = new Font("Microsoft YaHei UI", 9.5F);
            Panel pnlTab2 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            pnlTab2.Controls.Add(CreateHelpTextBox(GetTab2HelpText()));
            tab2.Controls.Add(pnlTab2);

            TabPage tab3 = new TabPage("📊 PDF 提取 Excel");
            tab3.BackColor = Color.White;
            tab3.Font = new Font("Microsoft YaHei UI", 9.5F);
            Panel pnlTab3 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            pnlTab3.Controls.Add(CreateHelpTextBox(GetTab3HelpText()));
            tab3.Controls.Add(pnlTab3);

            TabPage tab4 = new TabPage("📑 Excel 转成 PDF");
            tab4.BackColor = Color.White;
            tab4.Font = new Font("Microsoft YaHei UI", 9.5F);
            Panel pnlTab4 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 12, 16, 12) };
            pnlTab4.Controls.Add(CreateHelpTextBox(GetTab4HelpText()));
            tab4.Controls.Add(pnlTab4);

            helpTabControl.TabPages.AddRange(new TabPage[] { tab1, tab2, tab3, tab4 });
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

        private string GetTab3HelpText()
        {
            return 
@"【一、PDF 提取 Excel 核心特性】
1. 专为财务报表深度定制：智能识别中国小企业会计准则、企业会计准则标准《资产负债表》（8列双栏对称布局）、《利润表》（4列损益单栏）、《现金流量表》等各类电子税务局申报报表。
2. 零丢失高精度坐标还原：通过底层解析 PDF 矢量文字流物理坐标，毫秒级聚类排版，表头、行次、期末余额、年初余额绝对不串行错位。
3. 纯正数值格式与公式支持：提取的所有金额数字以纯浮点数值保存并应用千分位财务格式，导出至 Excel 后可直接使用 SUM() 等公式联动计算！

【二、导出模式与多工作表】
1. 多 Sheet 智能合规打包：同一个 PDF 文件内的多张报表（如第 1 页资产负债表、第 2 页利润表、第 3 页现金流量表）自动提取为同一个 Excel 工作簿的不同工作表。
2. 格式自由选择：支持直接导出标准现代化 Excel 工作簿 (*.xlsx)，也支持导出轻量 CSV 文本 (*.csv)。
3. 所见即所得实时表格预览：在界面右侧内置 DataGridView，载入 PDF 后即可切换查看每个 Sheet 的真实排版数据。

【三、数据安全与隐私承诺】
• 100% 纯本地单机离线运行，无需安装 Microsoft Office，不调用任何云端 API，财务核心机密绝不上网！

======================================================================
🔗 开源项目主页 (GitHub)：https://github.com/786381743syq/FinancePdfTool
欢迎 Star 支持与提交反馈！";
        }

        private string GetTab4HelpText()
        {
            return 
@"【一、Excel 转成 PDF 核心特性】
1. 支持格式广泛：支持直接载入标准 Excel 工作簿 (*.xlsx)、早期工作簿 (*.xls) 以及逗号分隔表格 (*.csv)。
2. 多工作表全量转换：支持一键将 Excel 工作簿内的全部 Sheet 合并渲染为单本 PDF，也可按需单独导出某一指定工作表。
3. 智能纸张自适应：
   • 🧭 智能感应：当工作表列数较多（>5列）时自动转为横向 A4 宽幅排版，避免挤压；列数较少时自动竖向；
   • 也可强制锁定为「横向」或「纵向」。

【二、精致财务美学样式】
1. 经典财务蓝（推荐）：柔和商务蓝表头，交替条纹斑马线，清晰浅灰网格，适合公司正式对账与内部汇报。
2. 极简黑白网格：纯正黑白分明线条，适合正式红头发文、公章加盖与黑白激光打印。
3. 现代商务灰：淡雅灰阶风格，低调沉稳。

【三、智能分页与页码】
• 自动按行高进行精确分页，表头自动对齐，并在每页底部打印「工作表名称  第 X / Y 页」页脚，专业规范。

======================================================================
🔗 开源项目主页 (GitHub)：https://github.com/786381743syq/FinancePdfTool
欢迎 Star 支持与提交反馈！";
        }
    }

    // =========================================================================
    // Supporting Classes for PDF-to-Excel & Excel-to-PDF Conversion
    // =========================================================================

    public class SimpleZipWriter
    {
        class ZipEntry
        {
            public string Path;
            public uint Crc;
            public uint CompressedSize;
            public uint UncompressedSize;
            public uint HeaderOffset;
            public ushort Method;
            public byte[] CompressedBytes;
        }

        public static byte[] CreateZip(Dictionary<string, byte[]> files)
        {
            using (var ms = new MemoryStream())
            using (var bw = new BinaryWriter(ms))
            {
                var entries = new List<ZipEntry>();

                foreach (var kvp in files)
                {
                    string entryPath = kvp.Key.Replace('\\', '/');
                    byte[] raw = kvp.Value;
                    uint crc = ComputeCrc32(raw);

                    byte[] comp;
                    ushort method;
                    using (var compMs = new MemoryStream())
                    {
                        using (var deflate = new DeflateStream(compMs, CompressionMode.Compress, true))
                        {
                            deflate.Write(raw, 0, raw.Length);
                        }
                        comp = compMs.ToArray();
                    }

                    if (comp.Length < raw.Length)
                    {
                        method = 8;
                    }
                    else
                    {
                        method = 0;
                        comp = raw;
                    }

                    var entry = new ZipEntry
                    {
                        Path = entryPath,
                        Crc = crc,
                        CompressedSize = (uint)comp.Length,
                        UncompressedSize = (uint)raw.Length,
                        HeaderOffset = (uint)ms.Position,
                        Method = method,
                        CompressedBytes = comp
                    };
                    entries.Add(entry);

                    byte[] nameBytes = Encoding.UTF8.GetBytes(entryPath);

                    bw.Write(0x04034b50);
                    bw.Write((ushort)20);
                    bw.Write((ushort)0x0800);
                    bw.Write(entry.Method);
                    bw.Write((ushort)0);
                    bw.Write((ushort)0);
                    bw.Write(entry.Crc);
                    bw.Write(entry.CompressedSize);
                    bw.Write(entry.UncompressedSize);
                    bw.Write((ushort)nameBytes.Length);
                    bw.Write((ushort)0);
                    bw.Write(nameBytes);
                    bw.Write(entry.CompressedBytes);
                }

                uint cdStart = (uint)ms.Position;

                foreach (var entry in entries)
                {
                    byte[] nameBytes = Encoding.UTF8.GetBytes(entry.Path);
                    bw.Write(0x02014b50);
                    bw.Write((ushort)20);
                    bw.Write((ushort)20);
                    bw.Write((ushort)0x0800);
                    bw.Write(entry.Method);
                    bw.Write((ushort)0);
                    bw.Write((ushort)0);
                    bw.Write(entry.Crc);
                    bw.Write(entry.CompressedSize);
                    bw.Write(entry.UncompressedSize);
                    bw.Write((ushort)nameBytes.Length);
                    bw.Write((ushort)0);
                    bw.Write((ushort)0);
                    bw.Write((ushort)0);
                    bw.Write((ushort)0);
                    bw.Write((uint)0);
                    bw.Write(entry.HeaderOffset);
                    bw.Write(nameBytes);
                }

                uint cdSize = (uint)ms.Position - cdStart;

                bw.Write(0x06054b50);
                bw.Write((ushort)0);
                bw.Write((ushort)0);
                bw.Write((ushort)entries.Count);
                bw.Write((ushort)entries.Count);
                bw.Write(cdSize);
                bw.Write(cdStart);
                bw.Write((ushort)0);

                return ms.ToArray();
            }
        }

        static uint ComputeCrc32(byte[] data)
        {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < data.Length; i++)
            {
                byte b = data[i];
                crc ^= b;
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 1) != 0) crc = (crc >> 1) ^ 0xEDB88320;
                    else crc >>= 1;
                }
            }
            return ~crc;
        }
    }

    public class SimpleZipReader
    {
        public static Dictionary<string, byte[]> ReadZip(byte[] zipBytes)
        {
            var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            int eocdPos = -1;
            for (int i = zipBytes.Length - 22; i >= 0; i--)
            {
                if (zipBytes[i] == 0x50 && zipBytes[i + 1] == 0x4b && zipBytes[i + 2] == 0x05 && zipBytes[i + 3] == 0x06)
                {
                    eocdPos = i;
                    break;
                }
            }
            if (eocdPos == -1) return result;

            using (var ms = new MemoryStream(zipBytes))
            using (var br = new BinaryReader(ms))
            {
                ms.Position = eocdPos + 10;
                ushort totalEntries = br.ReadUInt16();
                uint cdSize = br.ReadUInt32();
                uint cdOffset = br.ReadUInt32();

                ms.Position = cdOffset;
                for (int e = 0; e < totalEntries; e++)
                {
                    uint sig = br.ReadUInt32();
                    if (sig != 0x02014b50) break;

                    br.ReadUInt16();
                    br.ReadUInt16();
                    ushort flags = br.ReadUInt16();
                    ushort method = br.ReadUInt16();
                    br.ReadUInt32();
                    uint crc = br.ReadUInt32();
                    uint compSize = br.ReadUInt32();
                    uint uncompSize = br.ReadUInt32();
                    ushort nameLen = br.ReadUInt16();
                    ushort extraLen = br.ReadUInt16();
                    ushort commentLen = br.ReadUInt16();
                    br.ReadUInt32();
                    br.ReadUInt32();
                    uint localHeaderOffset = br.ReadUInt32();

                    byte[] nameBytes = br.ReadBytes(nameLen);
                    string entryName = Encoding.UTF8.GetString(nameBytes);
                    if (extraLen > 0) br.ReadBytes(extraLen);
                    if (commentLen > 0) br.ReadBytes(commentLen);

                    long curCdPos = ms.Position;
                    ms.Position = localHeaderOffset;
                    uint localSig = br.ReadUInt32();
                    if (localSig == 0x04034b50)
                    {
                        ms.Position += 22;
                        ushort locNameLen = br.ReadUInt16();
                        ushort locExtraLen = br.ReadUInt16();
                        ms.Position += locNameLen + locExtraLen;

                        byte[] compData = br.ReadBytes((int)compSize);
                        byte[] uncompData;

                        if (method == 0) uncompData = compData;
                        else if (method == 8)
                        {
                            using (var compMs = new MemoryStream(compData))
                            using (var def = new DeflateStream(compMs, CompressionMode.Decompress))
                            using (var outMs = new MemoryStream())
                            {
                                def.CopyTo(outMs);
                                uncompData = outMs.ToArray();
                            }
                        }
                        else uncompData = new byte[0];

                        result[entryName] = uncompData;
                    }
                    ms.Position = curCdPos;
                }
            }
            return result;
        }
    }

    public class TableCell
    {
        public string Text = "";
        public bool IsNumeric = false;
        public double NumericValue = 0;
    }

    public class ReportHeaderInfo
    {
        public string Title = "";
        public string FormCode = "";
        public string TaxId = "";
        public string TaxPeriod = "";
        public string CompanyName = "";
        public string FilingDate = "";
        public string MonetaryUnit = "单位：元";
    }

    public class TableSheet
    {
        public string Title = "工作表";
        public ReportHeaderInfo HeaderInfo = new ReportHeaderInfo();
        public List<string> MetaLines = new List<string>();
        public List<string> Headers = new List<string>();
        public List<List<TableCell>> Rows = new List<List<TableCell>>();
    }

    public class PdfTableExtractor
    {
        public class TextChunk
        {
            public double X;
            public double Y;
            public string Text;
            public int StreamIndex;
        }

        public static byte[] DecompressZlib(byte[] input)
        {
            if (input == null || input.Length < 6) return null;
            using (var msInput = new MemoryStream(input, 2, input.Length - 6))
            using (var deflate = new DeflateStream(msInput, CompressionMode.Decompress))
            using (var msOutput = new MemoryStream())
            {
                try { deflate.CopyTo(msOutput); return msOutput.ToArray(); }
                catch { return null; }
            }
        }

        public static Dictionary<int, string> ParseCMap(string cmapContent)
    {
        var map = new Dictionary<int, string>();
        if (string.IsNullOrEmpty(cmapContent)) return map;

        var bfcMatches = Regex.Matches(cmapContent, @"beginbfchar([\s\S]*?)endbfchar");
        foreach (Match m in bfcMatches)
        {
            var hexTokens = Regex.Matches(m.Groups[1].Value, @"<([0-9A-Fa-f]+)>");
            for (int i = 0; i + 1 < hexTokens.Count; i += 2)
            {
                int code = Convert.ToInt32(hexTokens[i].Groups[1].Value, 16);
                string uHex = hexTokens[i + 1].Groups[1].Value;
                StringBuilder sb = new StringBuilder();
                for (int u = 0; u + 4 <= uHex.Length; u += 4)
                {
                    int uVal = Convert.ToInt32(uHex.Substring(u, 4), 16);
                    sb.Append((char)uVal);
                }
                if (sb.Length == 0 && uHex.Length > 0)
                {
                    int uVal = Convert.ToInt32(uHex, 16);
                    sb.Append((char)uVal);
                }
                map[code] = sb.ToString();
            }
        }

        var bfrMatches = Regex.Matches(cmapContent, @"beginbfrange([\s\S]*?)endbfrange");
        foreach (Match m in bfrMatches)
        {
            var hexTokens = Regex.Matches(m.Groups[1].Value, @"<([0-9A-Fa-f]+)>");
            for (int i = 0; i + 2 < hexTokens.Count; i += 3)
            {
                int start = Convert.ToInt32(hexTokens[i].Groups[1].Value, 16);
                int end = Convert.ToInt32(hexTokens[i + 1].Groups[1].Value, 16);
                int targetStart = Convert.ToInt32(hexTokens[i + 2].Groups[1].Value, 16);
                for (int c = start; c <= end; c++)
                {
                    int uVal = targetStart + (c - start);
                    map[c] = ((char)uVal).ToString();
                }
            }
        }
        return map;
    }

        static double[] Concat(double[] m, double[] ctm)
        {
            return new double[]
            {
                m[0]*ctm[0] + m[1]*ctm[2],
                m[0]*ctm[1] + m[1]*ctm[3],
                m[2]*ctm[0] + m[3]*ctm[2],
                m[2]*ctm[1] + m[3]*ctm[3],
                m[4]*ctm[0] + m[5]*ctm[2] + ctm[4],
                m[4]*ctm[1] + m[5]*ctm[3] + ctm[5]
            };
        }

        public static List<TextChunk> ExtractChunks(string content, Dictionary<int, string> cmap)
        {
            var chunks = new List<TextChunk>();
            var stateStack = new Stack<double[]>();
            double[] ctm = new double[] { 1, 0, 0, 1, 0, 0 };
            double[] textMatrix = new double[] { 1, 0, 0, 1, 0, 0 };

            var regex = new Regex(@"(\[[\s\S]*?\])\s*TJ|(\([^\)]*\)|<[0-9A-Fa-f]+>)\s*Tj|([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+cm|([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+Tm|([-0-9.]+)\s+([-0-9.]+)\s+Td|\b(q|Q|BT|ET)\b");

            var matches = regex.Matches(content);
            foreach (Match m in matches)
            {
                string op = m.Value;
                if (m.Groups[17].Success)
                {
                    string cmd = m.Groups[17].Value;
                    if (cmd == "q") stateStack.Push((double[])ctm.Clone());
                    else if (cmd == "Q" && stateStack.Count > 0) ctm = stateStack.Pop();
                    else if (cmd == "BT") textMatrix = new double[] { 1, 0, 0, 1, 0, 0 };
                }
                else if (m.Groups[3].Success && op.EndsWith("cm"))
                {
                    double a = double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
                    double b = double.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
                    double c = double.Parse(m.Groups[5].Value, CultureInfo.InvariantCulture);
                    double d = double.Parse(m.Groups[6].Value, CultureInfo.InvariantCulture);
                    double e = double.Parse(m.Groups[7].Value, CultureInfo.InvariantCulture);
                    double f = double.Parse(m.Groups[8].Value, CultureInfo.InvariantCulture);
                    ctm = Concat(new double[] { a, b, c, d, e, f }, ctm);
                }
                else if (m.Groups[9].Success && op.EndsWith("Tm"))
                {
                    double a = double.Parse(m.Groups[9].Value, CultureInfo.InvariantCulture);
                    double b = double.Parse(m.Groups[10].Value, CultureInfo.InvariantCulture);
                    double c = double.Parse(m.Groups[11].Value, CultureInfo.InvariantCulture);
                    double d = double.Parse(m.Groups[12].Value, CultureInfo.InvariantCulture);
                    double e = double.Parse(m.Groups[13].Value, CultureInfo.InvariantCulture);
                    double f = double.Parse(m.Groups[14].Value, CultureInfo.InvariantCulture);
                    textMatrix = new double[] { a, b, c, d, e, f };
                }
                else if (m.Groups[15].Success && op.EndsWith("Td"))
                {
                    double tx = double.Parse(m.Groups[15].Value, CultureInfo.InvariantCulture);
                    double ty = double.Parse(m.Groups[16].Value, CultureInfo.InvariantCulture);
                    textMatrix = Concat(new double[] { 1, 0, 0, 1, tx, ty }, textMatrix);
                }
                else if (m.Groups[1].Success || m.Groups[2].Success)
                {
                    double[] eff = Concat(textMatrix, ctm);
                    double absX = eff[4];
                    double absY = eff[5];
                    string text = DecodeText(m.Value, cmap);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        chunks.Add(new TextChunk
                        {
                            X = Math.Round(absX, 1),
                            Y = Math.Round(absY, 1),
                            Text = text.Trim(),
                            StreamIndex = chunks.Count
                        });
                    }
                }
            }
            return chunks;
        }

        static string DecodeText(string chunk, Dictionary<int, string> cmap)
        {
            StringBuilder full = new StringBuilder();
            var tokens = Regex.Matches(chunk, @"<([0-9A-Fa-f]+)>|\(([\s\S]*?)\)");
            foreach (Match t in tokens)
            {
                if (t.Groups[1].Success)
                {
                    string hex = t.Groups[1].Value;
                    for (int h = 0; h < hex.Length; h += 4)
                    {
                        if (h + 4 <= hex.Length)
                        {
                            int cid = Convert.ToInt32(hex.Substring(h, 4), 16);
                            if (cmap != null && cmap.ContainsKey(cid)) full.Append(cmap[cid]);
                        }
                    }
                }
                else if (t.Groups[2].Success)
                {
                    string raw = t.Groups[2].Value;
                    foreach (char ch in raw)
                    {
                        if (ch == 0x93) full.Append('“');
                        else if (ch == 0x94) full.Append('”');
                        else if (ch == 0x91) full.Append('‘');
                        else if (ch == 0x92) full.Append('’');
                        else if (ch == 0x96 || ch == 0x97 || ch == 0xad) full.Append('-');
                        else full.Append(ch);
                    }
                }
            }
            return full.ToString();
        }

        public static List<TableSheet> ExtractWorkbook(string pdfPath)
        {
            var sheets = new List<TableSheet>();
            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);

            var contentRefs = Regex.Matches(raw, @"/Contents\s+(\d+)\s+0\s+R");
            var toUnicodeRefs = Regex.Matches(raw, @"/ToUnicode\s+(\d+)\s+0\s+R");

            int pageCount = contentRefs.Count;
            for (int p = 0; p < pageCount; p++)
            {
                int contentObjId = int.Parse(contentRefs[p].Groups[1].Value);
                int cmapObjId = (p < toUnicodeRefs.Count) ? int.Parse(toUnicodeRefs[p].Groups[1].Value) : (toUnicodeRefs.Count > 0 ? int.Parse(toUnicodeRefs[0].Groups[1].Value) : -1);

                Dictionary<int, string> cmap = null;
                if (cmapObjId > 0)
                {
                    string cmapStr = GetStreamString(bytes, raw, cmapObjId);
                    if (!string.IsNullOrEmpty(cmapStr)) cmap = ParseCMap(cmapStr);
                }

                string content = GetStreamString(bytes, raw, contentObjId);
                if (string.IsNullOrEmpty(content)) continue;

                var chunks = ExtractChunks(content, cmap);
                var sheet = ParsePageIntoSheet(chunks, p + 1);
                sheets.Add(sheet);
            }

            return sheets;
        }

        public static ReportHeaderInfo ExtractHeaderInfo(List<TextChunk> chunks, double headerY)
    {
        var info = new ReportHeaderInfo();
        if (headerY < 0) headerY = 740;

        var metaChunks = new List<TextChunk>();
        foreach (var c in chunks)
        {
            if (c.Y >= headerY - 2)
            {
                metaChunks.Add(c);
            }
        }

        metaChunks.Sort((a, b) => b.Y.CompareTo(a.Y));
        var metaRows = new List<List<TextChunk>>();
        List<TextChunk> curMeta = null;
        foreach (var c in metaChunks)
        {
            if (curMeta == null || Math.Abs(curMeta[0].Y - c.Y) > 5.0)
            {
                curMeta = new List<TextChunk>();
                metaRows.Add(curMeta);
            }
            curMeta.Add(c);
        }

        foreach (var mr in metaRows)
        {
            mr.Sort((a, b) => {
                if (Math.Abs(a.X - b.X) > 5.0) return a.X.CompareTo(b.X);
                return a.StreamIndex.CompareTo(b.StreamIndex);
            });
            StringBuilder sbLine = new StringBuilder();
            foreach (var ch in mr)
            {
                string text = ch.Text;
                if (sbLine.Length > 0)
                {
                    char lastCh = sbLine[sbLine.Length - 1];
                    char firstCh = text.Length > 0 ? text[0] : ' ';
                    bool isCjk = (lastCh >= 0x4e00 && lastCh <= 0x9fa5) || (firstCh >= 0x4e00 && firstCh <= 0x9fa5) ||
                                 lastCh == '（' || firstCh == '）' || lastCh == '：' || firstCh == '：' ||
                                 lastCh == '_' || firstCh == '_';
                    if (!isCjk && lastCh != ' ' && firstCh != ' ') sbLine.Append(" ");
                }
                sbLine.Append(text);
            }
            string line = sbLine.ToString().Trim();
            double y = mr[0].Y;

            if (y > 810 || line.Contains("表_年报") || line.Contains("（适用执行小企业"))
            {
                info.Title = line;
            }
            else if (line.Contains("会小企") || line.EndsWith("表"))
            {
                info.FormCode = line;
            }
            else if (line.Contains("纳税人识别号") || line.Contains("税款所属期"))
            {
                var leftParts = new StringBuilder();
                var rightParts = new StringBuilder();
                foreach (var ch in mr)
                {
                    if (ch.X < 260) leftParts.Append(ch.Text);
                    else rightParts.Append(ch.Text);
                }
                info.TaxId = leftParts.ToString().Trim();
                info.TaxPeriod = rightParts.ToString().Trim();
            }
            else if (line.Contains("编制单位") || line.Contains("报送日期") || line.Contains("单位："))
            {
                var leftParts = new StringBuilder();
                var midParts = new StringBuilder();
                var rightParts = new StringBuilder();
                foreach (var ch in mr)
                {
                    if (ch.X < 260) leftParts.Append(ch.Text);
                    else if (ch.X < 450) midParts.Append(ch.Text);
                    else rightParts.Append(ch.Text);
                }
                info.CompanyName = leftParts.ToString().Trim();
                info.FilingDate = midParts.ToString().Trim();
                info.MonetaryUnit = rightParts.ToString().Trim();
            }
        }

        if (string.IsNullOrEmpty(info.MonetaryUnit)) info.MonetaryUnit = "单位：元";
        return info;
    }

        public static TableSheet ParsePageIntoSheet(List<TextChunk> chunks, int pageIndex)
        {
            var sheet = new TableSheet();
            if (chunks.Count == 0) return sheet;

            chunks.Sort((a, b) => b.Y.CompareTo(a.Y));

            string detectedTitle = "";
            foreach (var c in chunks)
            {
                if (c.Y > 770)
                {
                    if (c.Text.Contains("资产负债表")) { detectedTitle = "资产负债表"; break; }
                    if (c.Text.Contains("利润表")) { detectedTitle = "利润表"; break; }
                    if (c.Text.Contains("现金流量表")) { detectedTitle = "现金流量表"; break; }
                }
            }

            bool isBalanceSheet = false;
            double headerY = -1;

            // 1. Check for Balance Sheet table header (Y > 600)
            foreach (var c in chunks)
            {
                if (c.Y > 600)
                {
                    if (c.Text.Contains("负债及所有者权益") || c.Text.Contains("负债和所有者权益") ||
                        (c.Text.Contains("所有者权益") && c.X > 250))
                    {
                        isBalanceSheet = true;
                        headerY = c.Y;
                        break;
                    }
                }
            }

            if (!isBalanceSheet)
            {
                // Secondary check for Balance Sheet: "期末余额" at X < 250 along with a "负债" chunk on the same line
                foreach (var c in chunks)
                {
                    if (c.Y > 600 && c.Text.Contains("期末余额") && c.X < 250)
                    {
                        foreach (var c2 in chunks)
                        {
                            if (Math.Abs(c2.Y - c.Y) < 4 && c2.Text.Contains("负债") && !c2.Text.Contains("流动"))
                            {
                                isBalanceSheet = true;
                                headerY = c.Y;
                                break;
                            }
                        }
                        if (isBalanceSheet) break;
                    }
                }
            }

            // 2. If not balance sheet or headerY still not found, look for "项目" at Y > 600
            if (headerY < 0)
            {
                foreach (var c in chunks)
                {
                    if (c.Y > 600 && c.Text == "项目" && c.X < 250)
                    {
                        headerY = c.Y;
                        break;
                    }
                }
            }

            // 3. Fallback for other standard headers
            if (headerY < 0)
            {
                foreach (var c in chunks)
                {
                    if (c.Y > 600 && (c.Text.Contains("本年累计") || c.Text.Contains("期末余额") || c.Text.Contains("行次")))
                    {
                        headerY = c.Y;
                        break;
                    }
                }
            }

            if (headerY > 0)
            {
                sheet.HeaderInfo = ExtractHeaderInfo(chunks, headerY);
            }

            if (detectedTitle == "资产负债表") sheet.Title = "资产负债表";
            else if (detectedTitle == "利润表") sheet.Title = "利润表";
            else if (detectedTitle == "现金流量表") sheet.Title = "现金流量表";
            else if (sheet.HeaderInfo != null && !string.IsNullOrEmpty(sheet.HeaderInfo.Title))
            {
                string t = sheet.HeaderInfo.Title;
                if (t.Length > 28) t = t.Substring(0, 28);
                sheet.Title = t;
            }
            else sheet.Title = "第" + pageIndex + "页报表";

            double[] colBounds;
            if (isBalanceSheet)
            {
                sheet.Headers = new List<string> { "资产", "行次", "期末余额", "年初余额", "负债及所有者权益", "行次", "期末余额", "年初余额" };
                colBounds = new double[] { 160, 185, 250, 300, 425, 448, 515 };
            }
            else
            {
                sheet.Headers = new List<string> { "项目", "行次", "本年累计金额", "上年金额" };
                colBounds = new double[] { 320, 380, 490 };
            }

            var bodyChunks = new List<TextChunk>();
            foreach (var c in chunks)
            {
                if (headerY > 0 && c.Y < headerY - 2 && c.Y > 20)
                {
                    bodyChunks.Add(c);
                }
            }

            bodyChunks.Sort((a, b) => {
                int cmp = b.Y.CompareTo(a.Y);
                if (cmp != 0) return cmp;
                return a.StreamIndex.CompareTo(b.StreamIndex);
            });

            var rowGroups = new List<List<TextChunk>>();
            List<TextChunk> curRow = null;

            foreach (var c in bodyChunks)
            {
                if (curRow == null || (curRow[curRow.Count - 1].Y - c.Y) > 5.8 || (curRow[0].Y - c.Y) > 12.0)
                {
                    curRow = new List<TextChunk>();
                    rowGroups.Add(curRow);
                }
                curRow.Add(c);
            }

            int colCount = sheet.Headers.Count;
            foreach (var rg in rowGroups)
            {
                var rowCells = new List<TableCell>();
                for (int i = 0; i < colCount; i++) rowCells.Add(new TableCell());

                rg.Sort((a, b) => {
                    if (Math.Abs(a.X - b.X) > 5.0) return a.X.CompareTo(b.X);
                    return a.StreamIndex.CompareTo(b.StreamIndex);
                });

                foreach (var ch in rg)
                {
                    int colIdx = 0;
                    while (colIdx < colBounds.Length && ch.X >= colBounds[colIdx])
                    {
                        colIdx++;
                    }
                    if (colIdx < colCount)
                    {
                        if (string.IsNullOrEmpty(rowCells[colIdx].Text))
                        {
                            rowCells[colIdx].Text = ch.Text;
                        }
                        else
                        {
                            string prev = rowCells[colIdx].Text;
                            char lastChar = prev[prev.Length - 1];
                            char firstChar = ch.Text.Length > 0 ? ch.Text[0] : ' ';
                            bool isCjk = (lastChar >= 0x4e00 && lastChar <= 0x9fa5) || (firstChar >= 0x4e00 && firstChar <= 0x9fa5) ||
                                         lastChar == '（' || firstChar == '）' || lastChar == '(' || firstChar == ')' ||
                                         lastChar == '“' || firstChar == '”' || lastChar == '"' || firstChar == '"' ||
                                         lastChar == '-' || firstChar == '-' || lastChar == '、' || firstChar == '、';
                            rowCells[colIdx].Text += (isCjk ? "" : " ") + ch.Text;
                        }
                    }
                }

                bool hasContent = false;
                for (int i = 0; i < colCount; i++)
                {
                    string val = rowCells[i].Text.Trim();
                    if (!string.IsNullOrEmpty(val)) hasContent = true;

                    double num;
                    string clean = val.Replace(",", "");
                    if (double.TryParse(clean, NumberStyles.Float, CultureInfo.InvariantCulture, out num) &&
                        (val.Contains(".") || val == "0" || (val.Length > 2 && !val.StartsWith("0"))))
                    {
                        rowCells[i].IsNumeric = true;
                        rowCells[i].NumericValue = num;
                    }
                }

                if (hasContent) sheet.Rows.Add(rowCells);
            }

            return sheet;
        }

        static string GetStreamString(byte[] bytes, string raw, int objId)
        {
            string pattern = objId + @"\s+0\s+obj[\s\S]*?stream\r?\n";
            Match m = Regex.Match(raw, pattern);
            if (!m.Success) return null;
            int streamStart = m.Index + m.Length;
            int endstreamPos = raw.IndexOf("endstream", streamStart);
            if (endstreamPos <= streamStart) return null;
            byte[] streamBytes = new byte[endstreamPos - streamStart];
            Buffer.BlockCopy(bytes, streamStart, streamBytes, 0, streamBytes.Length);
            byte[] decomp = DecompressZlib(streamBytes);
            if (decomp == null) return null;
            char[] chars = new char[decomp.Length];
            for (int i = 0; i < decomp.Length; i++) chars[i] = (char)decomp[i];
            return new string(chars);
        }
    }

    public class ExcelBuilder
    {
        static string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        static string GetColName(int col)
        {
            string res = "";
            while (col > 0)
            {
                col--;
                res = (char)('A' + (col % 26)) + res;
                col /= 26;
            }
            return res;
        }

        static double MeasureDisplayLen(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            double len = 0;
            foreach (char ch in text)
            {
                if (ch >= 0x4e00 && ch <= 0x9fa5) len += 2.1;
                else if (ch >= 0xff00 || ch == '（' || ch == '）' || ch == '：' || ch == '、') len += 2.0;
                else len += 1.1;
            }
            return len;
        }

        public static byte[] GenerateXlsx(List<TableSheet> sheets)
        {
            var files = new Dictionary<string, byte[]>();

            var sbTypes = new StringBuilder();
            sbTypes.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sbTypes.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sbTypes.AppendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sbTypes.AppendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sbTypes.AppendLine("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sbTypes.AppendLine("  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 0; i < sheets.Count; i++)
            {
                sbTypes.AppendLine(string.Format("  <Override PartName=\"/xl/worksheets/sheet{0}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i + 1));
            }
            sbTypes.AppendLine("</Types>");
            files["[Content_Types].xml"] = Encoding.UTF8.GetBytes(sbTypes.ToString());

            files["_rels/.rels"] = Encoding.UTF8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\r\n" +
                "</Relationships>");

            var sbWbRels = new StringBuilder();
            sbWbRels.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sbWbRels.AppendLine("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < sheets.Count; i++)
            {
                sbWbRels.AppendLine(string.Format("  <Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{0}.xml\"/>", i + 1));
            }
            sbWbRels.AppendLine(string.Format("  <Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>", sheets.Count + 1));
            sbWbRels.AppendLine("</Relationships>");
            files["xl/_rels/workbook.xml.rels"] = Encoding.UTF8.GetBytes(sbWbRels.ToString());

            var sbWb = new StringBuilder();
            sbWb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sbWb.AppendLine("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sbWb.AppendLine("  <sheets>");
            for (int i = 0; i < sheets.Count; i++)
            {
                string sName = sheets[i].Title;
                if (string.IsNullOrEmpty(sName)) sName = "Sheet" + (i + 1);
                sbWb.AppendLine(string.Format("    <sheet name=\"{0}\" sheetId=\"{1}\" r:id=\"rId{1}\"/>", EscapeXml(sName), i + 1));
            }
            sbWb.AppendLine("  </sheets>");
            sbWb.AppendLine("</workbook>");
            files["xl/workbook.xml"] = Encoding.UTF8.GetBytes(sbWb.ToString());

            // Professional Financial Style Sheet
            files["xl/styles.xml"] = Encoding.UTF8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n" +
                "  <numFmts count=\"1\">\r\n" +
                "    <numFmt numFmtId=\"164\" formatCode=\"#,##0.00;[Red]-#,##0.00;0.00\"/>\r\n" +
                "  </numFmts>\r\n" +
                "  <fonts count=\"6\">\r\n" +
                "    <font><name val=\"Microsoft YaHei\"/><sz val=\"10\"/><color rgb=\"FF333333\"/></font>\r\n" +
                "    <font><b/><name val=\"Microsoft YaHei\"/><sz val=\"10.5\"/><color rgb=\"FFFFFFFF\"/></font>\r\n" +
                "    <font><b/><name val=\"Microsoft YaHei\"/><sz val=\"15\"/><color rgb=\"FF1F4E78\"/></font>\r\n" +
                "    <font><name val=\"Microsoft YaHei\"/><sz val=\"9.5\"/><color rgb=\"FF595959\"/></font>\r\n" +
                "    <font><b/><name val=\"Microsoft YaHei\"/><sz val=\"10\"/><color rgb=\"FF1F4E78\"/></font>\r\n" +
                "    <font><b/><name val=\"Microsoft YaHei\"/><sz val=\"9.5\"/><color rgb=\"FF333333\"/></font>\r\n" +
                "  </fonts>\r\n" +
                "  <fills count=\"5\">\r\n" +
                "    <fill><patternFill patternType=\"none\"/></fill>\r\n" +
                "    <fill><patternFill patternType=\"gray125\"/></fill>\r\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF2E5B82\"/></patternFill></fill>\r\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF7FAFC\"/></patternFill></fill>\r\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE8EEF5\"/></patternFill></fill>\r\n" +
                "  </fills>\r\n" +
                "  <borders count=\"4\">\r\n" +
                "    <border><left/><right/><top/><bottom/></border>\r\n" +
                "    <border><left style=\"thin\"><color rgb=\"FFD9D9D9\"/></left><right style=\"thin\"><color rgb=\"FFD9D9D9\"/></right><top style=\"thin\"><color rgb=\"FFD9D9D9\"/></top><bottom style=\"thin\"><color rgb=\"FFD9D9D9\"/></bottom></border>\r\n" +
                "    <border><left style=\"thin\"><color rgb=\"FF203764\"/></left><right style=\"thin\"><color rgb=\"FF203764\"/></right><top style=\"thin\"><color rgb=\"FF203764\"/></top><bottom style=\"medium\"><color rgb=\"FF203764\"/></bottom></border>\r\n" +
                "    <border><left style=\"thin\"><color rgb=\"FFD9D9D9\"/></left><right style=\"thin\"><color rgb=\"FFD9D9D9\"/></right><top style=\"thin\"><color rgb=\"FFB0C4DE\"/></top><bottom style=\"double\"><color rgb=\"FF1F4E78\"/></bottom></border>\r\n" +
                "  </borders>\r\n" +
                "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\r\n" +
                "  <cellXfs count=\"24\">\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"2\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"4\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"4\" borderId=\"3\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"4\" fillId=\"4\" borderId=\"3\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"4\" borderId=\"3\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"4\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"164\" fontId=\"4\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf>\r\n" +
                "    <xf numFmtId=\"0\" fontId=\"4\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\r\n" +
                "  </cellXfs>\r\n" +
                "</styleSheet>");

            for (int sIdx = 0; sIdx < sheets.Count; sIdx++)
            {
                var sheet = sheets[sIdx];
                var sbWs = new StringBuilder();
                sbWs.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sbWs.AppendLine("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                sbWs.AppendLine("  <sheetViews>");
                sbWs.AppendLine("    <sheetView workbookViewId=\"0\"/>");
                sbWs.AppendLine("  </sheetViews>");
                sbWs.AppendLine("  <sheetFormatPr defaultRowHeight=\"21\"/>");

                int maxCols = sheet.Headers.Count;
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    if (sheet.Rows[r].Count > maxCols) maxCols = sheet.Rows[r].Count;
                }

                if (maxCols > 0)
                {
                    sbWs.AppendLine("  <cols>");
                    for (int c = 0; c < maxCols; c++)
                    {
                        double maxDisplayLen = 0;
                        if (c < sheet.Headers.Count)
                        {
                            maxDisplayLen = Math.Max(maxDisplayLen, MeasureDisplayLen(sheet.Headers[c]));
                        }
                        for (int r = 0; r < sheet.Rows.Count; r++)
                        {
                            if (c < sheet.Rows[r].Count)
                            {
                                var cell = sheet.Rows[r][c];
                                if (cell.IsNumeric)
                                {
                                    string numStr = cell.NumericValue.ToString("#,##0.00", CultureInfo.InvariantCulture);
                                    maxDisplayLen = Math.Max(maxDisplayLen, numStr.Length * 1.05);
                                }
                                else if (!string.IsNullOrEmpty(cell.Text))
                                {
                                    maxDisplayLen = Math.Max(maxDisplayLen, MeasureDisplayLen(cell.Text));
                                }
                            }
                        }

                        double finalWidth;
                        string headerText = (c < sheet.Headers.Count) ? sheet.Headers[c] : "";
                        if (headerText.Contains("余额") || headerText.Contains("金额"))
                        {
                            finalWidth = Math.Max(16.5, maxDisplayLen + 3.0);
                        }
                        else if (headerText == "行次")
                        {
                            finalWidth = Math.Max(8.0, maxDisplayLen + 2.0);
                        }
                        else
                        {
                            finalWidth = Math.Max(18.0, maxDisplayLen + 3.0);
                        }
                        finalWidth = Math.Round(Math.Min(finalWidth, 50.0), 1);
                        sbWs.AppendLine(string.Format("    <col min=\"{0}\" max=\"{0}\" width=\"{1}\" customWidth=\"1\"/>", c + 1, finalWidth.ToString("F1", CultureInfo.InvariantCulture)));
                    }
                    sbWs.AppendLine("  </cols>");
                }

                sbWs.AppendLine("  <sheetData>");
                int rowNum = 1;
                var merges = new List<string>();

                var hInfo = sheet.HeaderInfo;
                bool hasHeader = hInfo != null && !string.IsNullOrEmpty(hInfo.Title);

                if (hasHeader)
                {
                    // Row 1: Title
                    string lastCol = GetColName(maxCols);
                    sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"34\" customHeight=\"1\">", rowNum));
                    sbWs.AppendLine(string.Format("      <c r=\"A{0}\" s=\"13\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.Title)));
                    sbWs.AppendLine("    </row>");
                    merges.Add(string.Format("A{0}:{1}{0}", rowNum, lastCol));
                    rowNum++;

                    // Row 2: FormCode
                    if (!string.IsNullOrEmpty(hInfo.FormCode))
                    {
                        sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"18\" customHeight=\"1\">", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"A{0}\" s=\"14\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.FormCode)));
                        sbWs.AppendLine("    </row>");
                        rowNum++;
                    }

                    // Row 3: TaxId and TaxPeriod
                    sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"20\" customHeight=\"1\">", rowNum));
                    sbWs.AppendLine(string.Format("      <c r=\"A{0}\" s=\"14\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.TaxId)));
                    if (maxCols >= 8)
                    {
                        merges.Add(string.Format("A{0}:D{0}", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"E{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.TaxPeriod)));
                        merges.Add(string.Format("E{0}:H{0}", rowNum));
                    }
                    else if (maxCols >= 4)
                    {
                        merges.Add(string.Format("A{0}:B{0}", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"C{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.TaxPeriod)));
                        merges.Add(string.Format("C{0}:D{0}", rowNum));
                    }
                    else
                    {
                        sbWs.AppendLine(string.Format("      <c r=\"B{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.TaxPeriod)));
                    }
                    sbWs.AppendLine("    </row>");
                    rowNum++;

                    // Row 4: CompanyName, FilingDate, MonetaryUnit
                    sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"20\" customHeight=\"1\">", rowNum));
                    sbWs.AppendLine(string.Format("      <c r=\"A{0}\" s=\"14\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.CompanyName)));
                    if (maxCols >= 8)
                    {
                        merges.Add(string.Format("A{0}:D{0}", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"E{0}\" s=\"16\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.FilingDate)));
                        merges.Add(string.Format("E{0}:G{0}", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"H{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.MonetaryUnit)));
                    }
                    else if (maxCols >= 4)
                    {
                        merges.Add(string.Format("A{0}:B{0}", rowNum));
                        sbWs.AppendLine(string.Format("      <c r=\"C{0}\" s=\"16\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.FilingDate)));
                        sbWs.AppendLine(string.Format("      <c r=\"D{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.MonetaryUnit)));
                    }
                    else
                    {
                        sbWs.AppendLine(string.Format("      <c r=\"B{0}\" s=\"15\" t=\"inlineStr\"><is><t>{1}</t></is></c>", rowNum, EscapeXml(hInfo.MonetaryUnit)));
                    }
                    sbWs.AppendLine("    </row>");
                    rowNum++;
                }

                // Table Headers (Row 5 or rowNum)
                if (sheet.Headers.Count > 0)
                {
                    sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"26\" customHeight=\"1\">", rowNum));
                    for (int c = 0; c < sheet.Headers.Count; c++)
                    {
                        string colLetter = GetColName(c + 1);
                        sbWs.AppendLine(string.Format("      <c r=\"{0}{1}\" s=\"1\" t=\"inlineStr\"><is><t>{2}</t></is></c>", colLetter, rowNum, EscapeXml(sheet.Headers[c])));
                    }
                    sbWs.AppendLine("    </row>");
                    rowNum++;
                }

                // Table Data Rows
                bool is8Col = (sheet.Headers.Count >= 8 || maxCols >= 8);

                for (int rIdx = 0; rIdx < sheet.Rows.Count; rIdx++)
                {
                    var row = sheet.Rows[rIdx];
                    bool isEvenRow = (rIdx % 2 == 1);

                    // Check if summary row
                    bool isGrandTotal = false;
                    bool leftIsSubtotal = false;
                    bool rightIsSubtotal = false;

                    string c0 = (row.Count > 0) ? row[0].Text.Trim() : "";
                    string c4 = (row.Count > 4) ? row[4].Text.Trim() : "";

                    if (is8Col)
                    {
                        if (c0.Contains("资产总计") || c4.Contains("负债和所有者权益") || c4.Contains("负债及所有者权益") || c4.Contains("总计"))
                        {
                            isGrandTotal = true;
                        }
                        else
                        {
                            leftIsSubtotal = c0.Contains("合计") || c0.Contains("小计") || c0.EndsWith("：") || c0.EndsWith(":");
                            rightIsSubtotal = c4.Contains("合计") || c4.Contains("小计") || c4.EndsWith("：") || c4.EndsWith(":");
                        }
                    }
                    else
                    {
                        if (c0.Contains("净利润") || c0.Contains("期末现金余额"))
                        {
                            isGrandTotal = true;
                        }
                        else
                        {
                            leftIsSubtotal = c0.Contains("合计") || c0.Contains("小计") ||
                                             c0.Contains("营业利润") || c0.Contains("利润总额") ||
                                             c0.Contains("现金净增加额") || c0.EndsWith("：") || c0.EndsWith(":");
                        }
                    }

                    int ht = isGrandTotal ? 24 : ((leftIsSubtotal || rightIsSubtotal) ? 22 : 21);
                    sbWs.AppendLine(string.Format("    <row r=\"{0}\" ht=\"{1}\" customHeight=\"1\">", rowNum, ht));

                    for (int cIdx = 0; cIdx < row.Count; cIdx++)
                    {
                        string colLetter = GetColName(cIdx + 1);
                        var cell = row[cIdx];
                        string hName = (cIdx < sheet.Headers.Count) ? sheet.Headers[cIdx] : "";
                        bool isLineCol = (hName == "行次");

                        bool cellIsSubtotal = is8Col ? ((cIdx < 4) ? leftIsSubtotal : rightIsSubtotal) : leftIsSubtotal;

                        int style;
                        if (isGrandTotal)
                        {
                            style = cell.IsNumeric ? 11 : (isLineCol ? 12 : 10);
                        }
                        else if (cellIsSubtotal)
                        {
                            if (isEvenRow)
                            {
                                style = cell.IsNumeric ? 22 : (isLineCol ? 23 : 21);
                            }
                            else
                            {
                                style = cell.IsNumeric ? 19 : (isLineCol ? 20 : 18);
                            }
                        }
                        else if (isEvenRow)
                        {
                            style = cell.IsNumeric ? 5 : (isLineCol ? 6 : 4);
                        }
                        else
                        {
                            style = cell.IsNumeric ? 2 : (isLineCol ? 3 : 0);
                        }

                        if (cell.IsNumeric)
                        {
                            sbWs.AppendLine(string.Format("      <c r=\"{0}{1}\" s=\"{2}\"><v>{3}</v></c>", colLetter, rowNum, style, cell.NumericValue.ToString(CultureInfo.InvariantCulture)));
                        }
                        else
                        {
                            sbWs.AppendLine(string.Format("      <c r=\"{0}{1}\" s=\"{2}\" t=\"inlineStr\"><is><t>{3}</t></is></c>", colLetter, rowNum, style, EscapeXml(cell.Text)));
                        }
                    }
                    sbWs.AppendLine("    </row>");
                    rowNum++;
                }

                sbWs.AppendLine("  </sheetData>");

                if (merges.Count > 0)
                {
                    sbWs.AppendLine(string.Format("  <mergeCells count=\"{0}\">", merges.Count));
                    foreach (var m in merges)
                    {
                        sbWs.AppendLine(string.Format("    <mergeCell ref=\"{0}\"/>", m));
                    }
                    sbWs.AppendLine("  </mergeCells>");
                }

                sbWs.AppendLine("</worksheet>");
                files[string.Format("xl/worksheets/sheet{0}.xml", sIdx + 1)] = Encoding.UTF8.GetBytes(sbWs.ToString());
            }

            return SimpleZipWriter.CreateZip(files);
        }
    }

    public class ExcelSheetData
    {
        public string Name;
        public List<List<string>> Rows = new List<List<string>>();
    }

    public class ExcelReader
    {
        public static List<ExcelSheetData> ReadCsv(string filePath)
        {
            var list = new List<ExcelSheetData>();
            var sheet = new ExcelSheetData { Name = Path.GetFileNameWithoutExtension(filePath) };
            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var cells = new List<string>();
                bool inQuotes = false;
                StringBuilder cur = new StringBuilder();
                for (int i = 0; i < line.Length; i++)
                {
                    char ch = line[i];
                    if (ch == '\"')
                    {
                        inQuotes = !inQuotes;
                    }
                    else if (ch == ',' && !inQuotes)
                    {
                        cells.Add(cur.ToString().Trim());
                        cur.Length = 0;
                    }
                    else
                    {
                        cur.Append(ch);
                    }
                }
                cells.Add(cur.ToString().Trim());
                sheet.Rows.Add(cells);
            }
            list.Add(sheet);
            return list;
        }

        public static List<ExcelSheetData> ReadXlsx(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            var entries = SimpleZipReader.ReadZip(bytes);
            var sheets = new List<ExcelSheetData>();

            var sharedStrings = new List<string>();
            if (entries.ContainsKey("xl/sharedStrings.xml"))
            {
                string sstXml = Encoding.UTF8.GetString(entries["xl/sharedStrings.xml"]);
                var siMatches = Regex.Matches(sstXml, @"<si>([\s\S]*?)</si>");
                foreach (Match si in siMatches)
                {
                    var tMatches = Regex.Matches(si.Value, @"<t[^>]*>([\s\S]*?)</t>");
                    StringBuilder sb = new StringBuilder();
                    foreach (Match t in tMatches) sb.Append(t.Groups[1].Value);
                    sharedStrings.Add(System.Net.WebUtility.HtmlDecode(sb.ToString()));
                }
            }

            if (!entries.ContainsKey("xl/workbook.xml")) return sheets;
            string wbXml = Encoding.UTF8.GetString(entries["xl/workbook.xml"]);

            var sheetMatches = Regex.Matches(wbXml, @"<sheet[^>]+name=""([^""]+)""[^>]+r:id=""([^""]+)""");
            if (sheetMatches.Count == 0)
            {
                sheetMatches = Regex.Matches(wbXml, @"<sheet[^>]+r:id=""([^""]+)""[^>]+name=""([^""]+)""");
            }

            var rels = new Dictionary<string, string>();
            if (entries.ContainsKey("xl/_rels/workbook.xml.rels"))
            {
                string relsXml = Encoding.UTF8.GetString(entries["xl/_rels/workbook.xml.rels"]);
                var rMatches = Regex.Matches(relsXml, @"<Relationship[^>]+Id=""([^""]+)""[^>]+Target=""([^""]+)""");
                foreach (Match rm in rMatches)
                {
                    rels[rm.Groups[1].Value] = rm.Groups[2].Value.Replace('\\', '/');
                }
            }

            for (int i = 0; i < sheetMatches.Count; i++)
            {
                Match m = sheetMatches[i];
                string sheetName = m.Groups[1].Value;
                string rId = m.Groups[2].Value;
                if (rId.StartsWith("rId") == false && sheetName.StartsWith("rId"))
                {
                    string tmp = sheetName; sheetName = rId; rId = tmp;
                }

                string target = "worksheets/sheet" + (i + 1) + ".xml";
                if (rels.ContainsKey(rId)) target = rels[rId];
                if (!target.StartsWith("xl/")) target = "xl/" + target.TrimStart('/');

                if (!entries.ContainsKey(target)) continue;

                string wsXml = Encoding.UTF8.GetString(entries[target]);
                var sheetData = new ExcelSheetData { Name = sheetName };

                var rowMatches = Regex.Matches(wsXml, @"<row[^>]*>([\s\S]*?)</row>");
                foreach (Match rowM in rowMatches)
                {
                    var rowList = new List<string>();
                    var cellMatches = Regex.Matches(rowM.Value, @"<c\s+r=""([A-Z]+)(\d+)""([^>]*)>([\s\S]*?)</c>");
                    int curCol = 0;
                    foreach (Match cm in cellMatches)
                    {
                        string colLetters = cm.Groups[1].Value;
                        int targetCol = ColNameToIndex(colLetters);
                        while (curCol < targetCol)
                        {
                            rowList.Add("");
                            curCol++;
                        }

                        string attrs = cm.Groups[3].Value;
                        string inner = cm.Groups[4].Value;

                        string val = "";
                        if (attrs.Contains("t=\"s\""))
                        {
                            var vm = Regex.Match(inner, @"<v>(\d+)</v>");
                            if (vm.Success)
                            {
                                int sIdx = int.Parse(vm.Groups[1].Value);
                                if (sIdx >= 0 && sIdx < sharedStrings.Count) val = sharedStrings[sIdx];
                            }
                        }
                        else if (attrs.Contains("t=\"inlineStr\""))
                        {
                            var tm = Regex.Match(inner, @"<t[^>]*>([\s\S]*?)</t>");
                            if (tm.Success) val = System.Net.WebUtility.HtmlDecode(tm.Groups[1].Value);
                        }
                        else
                        {
                            var vm = Regex.Match(inner, @"<v>([\s\S]*?)</v>");
                            if (vm.Success) val = vm.Groups[1].Value;
                        }

                        rowList.Add(val);
                        curCol++;
                    }

                    if (rowList.Count > 0)
                    {
                        sheetData.Rows.Add(rowList);
                    }
                }

                sheets.Add(sheetData);
            }

            return sheets;
        }

        static int ColNameToIndex(string col)
        {
            int res = 0;
            for (int i = 0; i < col.Length; i++)
            {
                res = res * 26 + (col[i] - 'A' + 1);
            }
            return res - 1;
        }
    }

    public class ExcelToPdfRenderer
    {
        public static List<Bitmap> RenderSheetToBitmaps(ExcelSheetData sheet, int orientationMode = 0, int themeMode = 0, bool printFooter = true)
        {
            var bitmaps = new List<Bitmap>();
            if (sheet.Rows.Count == 0) return bitmaps;

            int tableHeaderRowIdx = -1;
            for (int r = 0; r < Math.Min(6, sheet.Rows.Count); r++)
            {
                var row = sheet.Rows[r];
                for (int c = 0; c < row.Count; c++)
                {
                    if (row[c] == "项目" || row[c] == "资产" || row[c] == "负债及所有者权益" || row[c] == "行次")
                    {
                        tableHeaderRowIdx = r;
                        break;
                    }
                }
                if (tableHeaderRowIdx >= 0) break;
            }

            int dataStartRow = (tableHeaderRowIdx >= 0) ? tableHeaderRowIdx : 0;
            int maxCols = 0;
            for (int r = dataStartRow; r < sheet.Rows.Count; r++) maxCols = Math.Max(maxCols, sheet.Rows[r].Count);
            if (maxCols == 0) return bitmaps;

            bool isLandscape;
            if (orientationMode == 1) isLandscape = true;
            else if (orientationMode == 2) isLandscape = false;
            else isLandscape = maxCols > 5;

            int pageWidth = isLandscape ? 2338 : 1654;
            int pageHeight = isLandscape ? 1654 : 2338;
            int marginX = 80;
            int marginTop = (tableHeaderRowIdx > 0) ? 140 : 100;
            int marginBottom = 90;
            int printableWidth = pageWidth - marginX * 2;

            float[] colWeights = new float[maxCols];
            for (int c = 0; c < maxCols; c++) colWeights[c] = 5f;

            for (int r = dataStartRow; r < sheet.Rows.Count; r++)
            {
                var row = sheet.Rows[r];
                for (int c = 0; c < row.Count; c++)
                {
                    string text = row[c];
                    float len = 0;
                    foreach (char ch in text)
                    {
                        if (ch >= 0x4e00 && ch <= 0x9fa5) len += 2.0f;
                        else if (ch >= 0xff00 || ch == '（' || ch == '）' || ch == '：' || ch == '、') len += 2.0f;
                        else len += 1.0f;
                    }
                    if (len > colWeights[c]) colWeights[c] = Math.Min(50f, len);
                }
            }

            float totalWeight = 0;
            for (int c = 0; c < maxCols; c++) totalWeight += colWeights[c];

            float[] colWidths = new float[maxCols];
            for (int c = 0; c < maxCols; c++)
            {
                colWidths[c] = (colWeights[c] / totalWeight) * printableWidth;
            }

            int rowHeight = 44;
            int headerHeight = 52;
            int printableTableHeight = pageHeight - marginTop - marginBottom - 120;
            int tableDataRowCount = sheet.Rows.Count - dataStartRow;
            int rowsPerPage = printableTableHeight / rowHeight;
            if (rowsPerPage < 10) rowsPerPage = 10;

            int totalPages = (int)Math.Ceiling((double)tableDataRowCount / rowsPerPage);
            if (totalPages == 0) totalPages = 1;

            Color hdrBgColor = Color.FromArgb(46, 91, 130);
            Color zebraColor = Color.FromArgb(247, 250, 252);
            Color gridColor = Color.FromArgb(216, 224, 232);
            Color thickColor = Color.FromArgb(31, 78, 120);
            Color titleColor = Color.FromArgb(31, 78, 120);

            if (themeMode == 1)
            {
                hdrBgColor = Color.FromArgb(235, 235, 235);
                zebraColor = Color.White;
                gridColor = Color.FromArgb(180, 180, 180);
                thickColor = Color.Black;
                titleColor = Color.Black;
            }
            else if (themeMode == 2)
            {
                hdrBgColor = Color.FromArgb(243, 244, 246);
                zebraColor = Color.FromArgb(249, 250, 251);
                gridColor = Color.FromArgb(229, 231, 235);
                thickColor = Color.FromArgb(107, 114, 128);
                titleColor = Color.FromArgb(31, 41, 55);
            }

            for (int p = 0; p < totalPages; p++)
            {
                Bitmap bmp = new Bitmap(pageWidth, pageHeight, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                    using (Font fontTitle = new Font("Microsoft YaHei", 18, FontStyle.Bold))
                    using (Font fontHeader = new Font("Microsoft YaHei", 10.5f, FontStyle.Bold))
                    using (Font fontBody = new Font("Microsoft YaHei", 9.5f, FontStyle.Regular))
                    using (Font fontFooter = new Font("Microsoft YaHei", 8.5f, FontStyle.Regular))
                    using (Brush brushTitle = new SolidBrush(titleColor))
                    using (Brush brushText = new SolidBrush(Color.FromArgb(40, 40, 40)))
                    using (Brush brushHdrText = new SolidBrush(themeMode == 0 ? Color.White : Color.FromArgb(30, 30, 30)))
                    using (Brush brushFooter = new SolidBrush(Color.FromArgb(140, 140, 140)))
                    using (Pen penGrid = new Pen(gridColor, 1.2f))
                    using (Pen penThick = new Pen(thickColor, 2f))
                    using (Brush brushHdrBg = new SolidBrush(hdrBgColor))
                    using (Brush brushZebra = new SolidBrush(zebraColor))
                    {
                        string mainTitle = sheet.Name;
                        if (tableHeaderRowIdx > 0 && sheet.Rows[0].Count > 0 && !string.IsNullOrEmpty(sheet.Rows[0][0]))
                        {
                            mainTitle = sheet.Rows[0][0];
                        }

                        var sfTitle = new StringFormat { Alignment = StringAlignment.Center };
                        g.DrawString(mainTitle, fontTitle, brushTitle, pageWidth / 2, 35, sfTitle);

                        if (tableHeaderRowIdx > 0)
                        {
                            float metaY = 75;
                            for (int m = 1; m < tableHeaderRowIdx; m++)
                            {
                                var mRow = sheet.Rows[m];
                                string leftText = (mRow.Count > 0) ? mRow[0] : "";
                                string rightText = (mRow.Count > 1) ? mRow[mRow.Count - 1] : "";
                                string midText = (mRow.Count > 2) ? mRow[mRow.Count / 2] : "";

                                if (!string.IsNullOrEmpty(leftText))
                                    g.DrawString(leftText, fontBody, brushFooter, marginX, metaY);
                                if (!string.IsNullOrEmpty(midText) && midText != leftText && midText != rightText)
                                    g.DrawString(midText, fontBody, brushFooter, pageWidth / 2, metaY, sfTitle);
                                if (!string.IsNullOrEmpty(rightText) && rightText != leftText)
                                {
                                    var sfRight = new StringFormat { Alignment = StringAlignment.Far };
                                    g.DrawString(rightText, fontBody, brushFooter, pageWidth - marginX, metaY, sfRight);
                                }
                                metaY += 24;
                            }
                        }

                        float curY = marginTop;
                        int startDataIdx = p * rowsPerPage;
                        int endDataIdx = Math.Min(tableDataRowCount, (p + 1) * rowsPerPage);

                        for (int i = startDataIdx; i < endDataIdx; i++)
                        {
                            int actualRowIdx = dataStartRow + i;
                            var row = sheet.Rows[actualRowIdx];
                            bool isHdr = (actualRowIdx == dataStartRow && tableHeaderRowIdx >= 0);
                            float rH = isHdr ? headerHeight : rowHeight;

                            if (isHdr)
                            {
                                g.FillRectangle(brushHdrBg, marginX, curY, printableWidth, rH);
                            }
                            else if (i % 2 == 1)
                            {
                                g.FillRectangle(brushZebra, marginX, curY, printableWidth, rH);
                            }

                            float curX = marginX;
                            for (int c = 0; c < maxCols; c++)
                            {
                                string text = (c < row.Count) ? row[c] : "";
                                Font f = isHdr ? fontHeader : fontBody;
                                Brush bText = isHdr ? brushHdrText : brushText;

                                double dummy;
                                bool isNum = double.TryParse(text.Replace(",", ""), out dummy);

                                var sf = new StringFormat
                                {
                                    Alignment = isHdr ? StringAlignment.Center : (isNum ? StringAlignment.Far : StringAlignment.Near),
                                    LineAlignment = StringAlignment.Center,
                                    Trimming = StringTrimming.EllipsisCharacter
                                };

                                var cellRect = new RectangleF(curX + 6, curY + 2, colWidths[c] - 12, rH - 4);
                                g.DrawString(text, f, bText, cellRect, sf);

                                g.DrawLine(penGrid, curX, curY, curX, curY + rH);
                                curX += colWidths[c];
                            }
                            g.DrawLine(penGrid, curX, curY, curX, curY + rH);

                            g.DrawLine(isHdr ? penThick : penGrid, marginX, curY + rH, marginX + printableWidth, curY + rH);
                            curY += rH;
                        }

                        g.DrawLine(penThick, marginX, marginTop, marginX + printableWidth, marginTop);

                        if (printFooter)
                        {
                            string footer = string.Format("工作表：{0}    第 {1} / {2} 页", sheet.Name, p + 1, totalPages);
                            var sfFooter = new StringFormat { Alignment = StringAlignment.Center };
                            g.DrawString(footer, fontFooter, brushFooter, pageWidth / 2, pageHeight - 55, sfFooter);
                        }
                    }
                }
                bitmaps.Add(bmp);
            }

            return bitmaps;
        }
    }

    public static class PdfBuilder
    {
        public static void BuildPdfFromBitmaps(List<Bitmap> images, string outputPath)
        {
            if (images == null || images.Count == 0) return;

            using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            using (var sw = new StreamWriter(fs, Encoding.ASCII))
            {
                var offsets = new List<long>();
                sw.Write("%PDF-1.4\r\n");
                sw.Flush();

                int count = images.Count;
                offsets.Add(fs.Position);
                sw.Write("1 0 obj\r\n<< /Type /Catalog /Pages 2 0 R >>\r\nendobj\r\n");
                sw.Flush();

                offsets.Add(fs.Position);
                StringBuilder kids = new StringBuilder();
                for (int i = 0; i < count; i++)
                {
                    kids.Append(string.Format("{0} 0 R ", 3 + i * 3));
                }
                sw.Write(string.Format("2 0 obj\r\n<< /Type /Pages /Kids [{0}] /Count {1} >>\r\nendobj\r\n", kids.ToString(), count));
                sw.Flush();

                for (int i = 0; i < count; i++)
                {
                    Bitmap bmp = images[i];
                    int pageObj = 3 + i * 3;
                    int contentObj = pageObj + 1;
                    int imageObj = pageObj + 2;

                    float ptWidth = (float)(bmp.Width * 72.0 / 200.0);
                    float ptHeight = (float)(bmp.Height * 72.0 / 200.0);

                    offsets.Add(fs.Position);
                    sw.Write(string.Format("{0} 0 obj\r\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {1:F2} {2:F2}] /Contents {3} 0 R /Resources << /XObject << /Im{4} {5} 0 R >> >> >>\r\nendobj\r\n",
                        pageObj, ptWidth, ptHeight, contentObj, i, imageObj));
                    sw.Flush();

                    string contentStream = string.Format("q\r\n{0:F2} 0 0 {1:F2} 0 0 cm\r\n/Im{2} Do\r\nQ\r\n", ptWidth, ptHeight, i);
                    byte[] contentBytes = Encoding.ASCII.GetBytes(contentStream);

                    offsets.Add(fs.Position);
                    sw.Write(string.Format("{0} 0 obj\r\n<< /Length {1} >>\r\nstream\r\n", contentObj, contentBytes.Length));
                    sw.Flush();
                    fs.Write(contentBytes, 0, contentBytes.Length);
                    sw.Write("\r\nendstream\r\nendobj\r\n");
                    sw.Flush();

                    byte[] jpegBytes;
                    using (var ms = new MemoryStream())
                    {
                        ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                        var myEncoderParameters = new EncoderParameters(1);
                        myEncoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 92L);
                        bmp.Save(ms, jpgEncoder, myEncoderParameters);
                        jpegBytes = ms.ToArray();
                    }

                    offsets.Add(fs.Position);
                    sw.Write(string.Format("{0} 0 obj\r\n<< /Type /XObject /Subtype /Image /Width {1} /Height {2} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {3} >>\r\nstream\r\n",
                        imageObj, bmp.Width, bmp.Height, jpegBytes.Length));
                    sw.Flush();
                    fs.Write(jpegBytes, 0, jpegBytes.Length);
                    sw.Write("\r\nendstream\r\nendobj\r\n");
                    sw.Flush();
                }

                long xrefPos = fs.Position;
                sw.Write(string.Format("xref\r\n0 {0}\r\n0000000000 65535 f \r\n", offsets.Count + 1));
                for (int i = 0; i < offsets.Count; i++)
                {
                    sw.Write(string.Format("{0:D10} 00000 n \r\n", offsets[i]));
                }
                sw.Write(string.Format("trailer\r\n<< /Size {0} /Root 1 0 R >>\r\nstartxref\r\n{1}\r\n%%EOF\r\n", offsets.Count + 1, xrefPos));
                sw.Flush();
            }
        }

        static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
        }
    }

}
