namespace ExcelFormulaRemover.App;

using System.Security;
using ExcelFormulaRemover.Core;

internal sealed class MainForm : Form
{
    private const int WindowWidth = 860;
    private const int WindowHeight = 700;
    private const int MaximumSummaryDetails = 6;

    private readonly InputWorkbookSet inputWorkbooks = new();
    private readonly DesktopConversionController conversionController =
        new(FormulaRemovalEngine.ConvertAsync);
    private readonly Font titleFont;
    private readonly ToolTip toolTip = new();
    private readonly ListView fileListView = new();
    private readonly Label fileCountLabel = new();
    private readonly Button addFilesButton = new();
    private readonly Button removeFilesButton = new();
    private readonly Button clearFilesButton = new();
    private readonly TextBox outputDirectoryTextBox = new();
    private readonly Button browseOutputButton = new();
    private readonly NumericUpDown startRowInput = new();
    private readonly TextBox excludedSheetsTextBox = new();
    private readonly ProgressBar progressBar = new();
    private readonly Label statusLabel = new();
    private readonly Button convertButton = new();
    private readonly Button cancelButton = new();

    private bool isProcessing;
    private bool closeAfterCancellation;
    private bool allowClose;

    public MainForm()
    {
        Font systemFont = SystemFonts.MessageBoxFont ?? Control.DefaultFont;
        titleFont = new Font(
            systemFont.FontFamily,
            systemFont.SizeInPoints + 8,
            FontStyle.Bold,
            GraphicsUnit.Point);

        InitializeWindow();
        InitializeLayout();
        WireEvents();
        UpdateFileList();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            conversionController.Dispose();
            toolTip.Dispose();
            titleFont.Dispose();
        }

        base.Dispose(disposing);
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        addFilesButton.Select();
    }

    private void InitializeWindow()
    {
        Text = "Excel 수식 값 변환기";
        AccessibleName = "Excel 수식 값 변환기";
        AccessibleDescription = "Excel 파일의 수식을 마지막으로 저장된 계산값으로 바꿉니다.";
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(WindowWidth, WindowHeight);
        MinimumSize = new Size(740, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Font = SystemFonts.MessageBoxFont;
        KeyPreview = true;
        AcceptButton = convertButton;
        CancelButton = cancelButton;
    }

    private void InitializeLayout()
    {
        SuspendLayout();

        TableLayoutPanel root = new()
        {
            ColumnCount = 1,
            Dock = DockStyle.Fill,
            Padding = new Padding(24),
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        Label titleLabel = new()
        {
            AutoSize = true,
            Font = titleFont,
            Margin = new Padding(0, 0, 0, 6),
            Text = "수식을 마지막 저장 값으로 바꾸기",
            UseMnemonic = false,
        };

        Label explanationLabel = new()
        {
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 18),
            MaximumSize = new Size(680, 0),
            Text = "원본 파일은 변경하지 않습니다. Excel이 마지막으로 저장한 계산값만 사용하며, 이 앱에서 수식을 다시 계산하지 않습니다.",
            UseMnemonic = false,
        };

        root.Controls.Add(titleLabel, 0, 0);
        root.Controls.Add(explanationLabel, 0, 1);
        root.Controls.Add(CreateFilesSection(), 0, 2);
        root.Controls.Add(CreateSettingsSection(), 0, 3);
        root.Controls.Add(CreateProgressSection(), 0, 4);
        root.Controls.Add(CreateActionSection(), 0, 5);

        Controls.Add(root);
        ResumeLayout(performLayout: true);
    }

    private GroupBox CreateFilesSection()
    {
        GroupBox group = new()
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 14),
            Padding = new Padding(12),
            TabIndex = 0,
            Text = "1. 원본 파일",
        };

        TableLayoutPanel layout = new()
        {
            ColumnCount = 1,
            Dock = DockStyle.Fill,
            RowCount = 4,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        FlowLayoutPanel actions = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 6),
            TabIndex = 0,
            WrapContents = true,
        };

        ConfigureButton(addFilesButton, "파일 추가(&A)...", "변환할 Excel 파일을 선택합니다.");
        ConfigureButton(removeFilesButton, "선택 제거(&R)", "목록에서 선택한 파일을 제거합니다.");
        ConfigureButton(clearFilesButton, "모두 지우기(&L)", "목록의 모든 파일을 제거합니다.");
        addFilesButton.TabIndex = 0;
        removeFilesButton.TabIndex = 1;
        clearFilesButton.TabIndex = 2;
        actions.Controls.AddRange([addFilesButton, removeFilesButton, clearFilesButton]);

        Label hintLabel = new()
        {
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = new Padding(3, 0, 0, 8),
            Text = ".xlsx, .xlsm 또는 .xlsb 파일을 선택하거나 아래 목록으로 끌어 놓으세요.",
            UseMnemonic = false,
        };

        fileListView.AccessibleName = "변환할 원본 파일";
        fileListView.AccessibleDescription = "추가한 Excel 파일의 이름과 폴더를 보여 줍니다.";
        fileListView.AllowDrop = true;
        fileListView.Dock = DockStyle.Fill;
        fileListView.FullRowSelect = true;
        fileListView.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        fileListView.HideSelection = false;
        fileListView.Margin = new Padding(0);
        fileListView.MultiSelect = true;
        fileListView.ShowItemToolTips = true;
        fileListView.TabIndex = 1;
        fileListView.TabStop = true;
        fileListView.View = View.Details;
        fileListView.Columns.Add("파일 이름");
        fileListView.Columns.Add("폴더");

        fileCountLabel.AutoSize = true;
        fileCountLabel.Dock = DockStyle.Fill;
        fileCountLabel.Margin = new Padding(3, 7, 0, 0);
        fileCountLabel.TextAlign = ContentAlignment.MiddleRight;
        fileCountLabel.UseMnemonic = false;

        layout.Controls.Add(actions, 0, 0);
        layout.Controls.Add(hintLabel, 0, 1);
        layout.Controls.Add(fileListView, 0, 2);
        layout.Controls.Add(fileCountLabel, 0, 3);
        group.Controls.Add(layout);
        return group;
    }

    private GroupBox CreateSettingsSection()
    {
        GroupBox group = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 14),
            Padding = new Padding(12),
            TabIndex = 1,
            Text = "2. 저장 위치와 범위",
        };

        TableLayoutPanel layout = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 3,
            Dock = DockStyle.Fill,
            RowCount = 4,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        Label outputLabel = CreateFieldLabel("출력 폴더(&O)");
        outputLabel.TabIndex = 0;
        outputDirectoryTextBox.AccessibleName = "출력 폴더";
        outputDirectoryTextBox.AccessibleDescription = "변환한 새 파일을 저장할 폴더입니다.";
        outputDirectoryTextBox.Dock = DockStyle.Fill;
        outputDirectoryTextBox.Margin = new Padding(8, 3, 8, 3);
        outputDirectoryTextBox.TabIndex = 1;
        toolTip.SetToolTip(outputDirectoryTextBox, "기존 파일은 덮어쓰지 않고 번호를 붙인 새 이름을 사용합니다.");

        ConfigureButton(browseOutputButton, "찾아보기(&B)...", "출력 폴더를 선택합니다.");
        browseOutputButton.Margin = new Padding(0, 0, 0, 4);
        browseOutputButton.TabIndex = 2;

        Label startRowLabel = CreateFieldLabel("시작 행(&S)");
        startRowLabel.TabIndex = 3;
        startRowInput.AccessibleName = "수식 변환 시작 행";
        startRowInput.AccessibleDescription = "이 행부터 마지막 행까지 수식을 값으로 바꿉니다.";
        startRowInput.Minimum = 1;
        startRowInput.Maximum = FormulaRemovalRequest.MaximumRowNumber;
        startRowInput.Value = 1;
        startRowInput.ThousandsSeparator = true;
        startRowInput.Width = 140;
        startRowInput.Margin = new Padding(8, 3, 8, 3);
        startRowInput.TabIndex = 4;

        Label startRowHint = new()
        {
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            TabIndex = 5,
            Text = "이 행부터 변환",
            UseMnemonic = false,
        };

        Label excludedSheetsLabel = CreateFieldLabel("제외할 시트(&E)");
        excludedSheetsLabel.TabIndex = 6;
        excludedSheetsTextBox.AccessibleName = "제외할 시트 이름 패턴";
        excludedSheetsTextBox.AccessibleDescription = "수식을 유지할 시트 이름의 일부를 쉼표나 세미콜론으로 구분해 입력합니다.";
        excludedSheetsTextBox.Dock = DockStyle.Fill;
        excludedSheetsTextBox.Margin = new Padding(8, 3, 0, 3);
        excludedSheetsTextBox.PlaceholderText = "예: TF, 변동성";
        excludedSheetsTextBox.TabIndex = 7;
        layout.SetColumnSpan(excludedSheetsTextBox, 2);

        Label excludedSheetsHint = new()
        {
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = new Padding(8, 2, 0, 0),
            TabIndex = 8,
            Text = "선택 사항 · 쉼표(,) 또는 세미콜론(;)으로 구분하면, 이름에 해당 문구가 들어간 시트의 수식을 유지합니다.",
            UseMnemonic = false,
        };
        layout.SetColumnSpan(excludedSheetsHint, 2);

        layout.Controls.Add(outputLabel, 0, 0);
        layout.Controls.Add(outputDirectoryTextBox, 1, 0);
        layout.Controls.Add(browseOutputButton, 2, 0);
        layout.Controls.Add(startRowLabel, 0, 1);
        layout.Controls.Add(startRowInput, 1, 1);
        layout.Controls.Add(startRowHint, 2, 1);
        layout.Controls.Add(excludedSheetsLabel, 0, 2);
        layout.Controls.Add(excludedSheetsTextBox, 1, 2);
        layout.Controls.Add(excludedSheetsHint, 1, 3);
        group.Controls.Add(layout);
        return group;
    }

    private TableLayoutPanel CreateProgressSection()
    {
        TableLayoutPanel layout = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 12),
            RowCount = 2,
            TabIndex = 2,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        statusLabel.AccessibleName = "변환 상태";
        statusLabel.AccessibleRole = AccessibleRole.StatusBar;
        statusLabel.AutoEllipsis = true;
        statusLabel.Dock = DockStyle.Fill;
        statusLabel.Margin = new Padding(0, 0, 0, 6);
        statusLabel.Text = "파일을 추가하면 시작할 수 있습니다.";
        statusLabel.UseMnemonic = false;

        progressBar.AccessibleName = "파일 변환 진행률";
        progressBar.Dock = DockStyle.Fill;
        progressBar.Maximum = 1;
        progressBar.Minimum = 0;
        progressBar.Value = 0;

        layout.Controls.Add(statusLabel, 0, 0);
        layout.Controls.Add(progressBar, 0, 1);
        return layout;
    }

    private FlowLayoutPanel CreateActionSection()
    {
        FlowLayoutPanel actions = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            Margin = new Padding(0),
            TabIndex = 3,
            WrapContents = false,
        };

        ConfigureButton(convertButton, "값으로 변환(&C)", "선택한 파일의 수식을 마지막 저장 값으로 바꿉니다.");
        convertButton.AutoSize = false;
        convertButton.Size = new Size(130, 36);
        convertButton.TabIndex = 0;

        ConfigureButton(cancelButton, "취소(&X)", "진행 중인 변환을 안전하게 취소합니다.");
        cancelButton.AutoSize = false;
        cancelButton.Size = new Size(90, 36);
        cancelButton.Enabled = false;
        cancelButton.TabIndex = 1;

        actions.Controls.Add(convertButton);
        actions.Controls.Add(cancelButton);
        return actions;
    }

    private void WireEvents()
    {
        addFilesButton.Click += AddFilesButton_Click;
        removeFilesButton.Click += RemoveFilesButton_Click;
        clearFilesButton.Click += ClearFilesButton_Click;
        browseOutputButton.Click += BrowseOutputButton_Click;
        convertButton.Click += ConvertButton_Click;
        cancelButton.Click += CancelButton_Click;
        outputDirectoryTextBox.TextChanged += (_, _) => UpdateActionState();
        fileListView.SelectedIndexChanged += (_, _) => UpdateActionState();
        fileListView.DragEnter += FileListView_DragEnter;
        fileListView.DragDrop += FileListView_DragDrop;
        fileListView.KeyDown += FileListView_KeyDown;
        fileListView.Resize += (_, _) => ResizeFileColumns();
        FormClosing += MainForm_FormClosing;
    }

    private static Label CreateFieldLabel(string text)
    {
        return new Label
        {
            Anchor = AnchorStyles.Left,
            AutoSize = true,
            AccessibleDescription = $"{text.Replace("&", string.Empty, StringComparison.Ordinal)} 입력란의 레이블",
            Margin = new Padding(0, 6, 0, 6),
            Text = text,
            TextAlign = ContentAlignment.MiddleLeft,
            UseMnemonic = true,
        };
    }

    private static void ConfigureButton(Button button, string text, string accessibleDescription)
    {
        button.AutoSize = true;
        button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        button.Margin = new Padding(0, 0, 8, 0);
        button.Padding = new Padding(8, 3, 8, 3);
        button.Text = text;
        button.AccessibleName = text.Replace("&", string.Empty, StringComparison.Ordinal);
        button.AccessibleDescription = accessibleDescription;
        button.UseVisualStyleBackColor = true;
    }

    private void AddFilesButton_Click(object? sender, EventArgs e)
    {
        using OpenFileDialog dialog = new()
        {
            CheckFileExists = true,
            Filter = "Excel 통합 문서 (*.xlsx;*.xlsm;*.xlsb)|*.xlsx;*.xlsm;*.xlsb",
            Multiselect = true,
            RestoreDirectory = true,
            Title = "변환할 Excel 파일 선택",
        };

        string? initialDirectory = GetInitialDirectory();
        if (initialDirectory is not null)
        {
            dialog.InitialDirectory = initialDirectory;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            AddWorkbooks(dialog.FileNames);
        }
    }

    private void RemoveFilesButton_Click(object? sender, EventArgs e)
    {
        string[] selectedPaths = fileListView.SelectedItems
            .Cast<ListViewItem>()
            .Select(item => (string)item.Tag!)
            .ToArray();

        if (selectedPaths.Length == 0)
        {
            return;
        }

        inputWorkbooks.Remove(selectedPaths);
        UpdateFileList();
        SetStatus($"{selectedPaths.Length:N0}개 파일을 목록에서 제거했습니다.");
    }

    private void ClearFilesButton_Click(object? sender, EventArgs e)
    {
        int removedCount = inputWorkbooks.Items.Count;
        if (removedCount == 0)
        {
            return;
        }

        inputWorkbooks.Clear();
        UpdateFileList();
        SetStatus("파일 목록을 비웠습니다.");
    }

    private void BrowseOutputButton_Click(object? sender, EventArgs e)
    {
        using FolderBrowserDialog dialog = new()
        {
            Description = "변환한 새 파일을 저장할 폴더를 선택하세요.",
            ShowNewFolderButton = true,
            UseDescriptionForTitle = true,
        };

        string? initialDirectory = GetInitialDirectory();
        if (initialDirectory is not null)
        {
            dialog.InitialDirectory = initialDirectory;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            outputDirectoryTextBox.Text = dialog.SelectedPath;
        }
    }

    private async void ConvertButton_Click(object? sender, EventArgs e)
    {
        FormulaRemovalRequest request;
        try
        {
            request = CreateRequest();
        }
        catch (Exception exception) when (
            exception is ArgumentException or NotSupportedException or PathTooLongException or
                SecurityException)
        {
            ShowRequestError(exception.Message);
            return;
        }

        isProcessing = true;
        progressBar.Maximum = Math.Max(1, request.InputPaths.Count);
        progressBar.Value = 0;
        SetProcessingState(processing: true);
        SetStatus($"변환 준비 중 · {request.InputPaths.Count:N0}개 파일");

        DesktopConversionOutcome outcome;
        try
        {
            Progress<FormulaRemovalProgress> progress = new(UpdateProgress);
            outcome = await conversionController.RunAsync(request, progress);
        }
        finally
        {
            isProcessing = false;
            SetProcessingState(processing: false);
        }

        if (closeAfterCancellation)
        {
            allowClose = true;
            BeginInvoke(new Action(Close));
            return;
        }

        if (outcome.Failure is not null)
        {
            SetStatus("변환을 시작하지 못했습니다. 출력 폴더와 파일 접근 권한을 확인하세요.");
            MessageBox.Show(
                this,
                "변환을 시작하지 못했습니다.\n\n출력 폴더 경로와 파일 접근 권한을 확인한 뒤 다시 시도하세요.\n\n" +
                    LimitText(outcome.Failure.Message, 500),
                "변환 시작 실패",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        ShowBatchResult(outcome.Result, outcome.WasCanceled, request.OutputDirectory);
    }

    private void CancelButton_Click(object? sender, EventArgs e)
    {
        RequestCancellation();
    }

    private void FileListView_DragEnter(object? sender, DragEventArgs e)
    {
        if (!isProcessing && TryGetDroppedFiles(e, out string[] files) &&
            files.Any(IsSupportedWorkbookPath))
        {
            e.Effect = DragDropEffects.Copy;
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    private void FileListView_DragDrop(object? sender, DragEventArgs e)
    {
        if (!isProcessing && TryGetDroppedFiles(e, out string[] files))
        {
            AddWorkbooks(files);
        }
    }

    private void FileListView_KeyDown(object? sender, KeyEventArgs e)
    {
        if (!isProcessing && e.KeyCode == Keys.Delete)
        {
            RemoveFilesButton_Click(sender, EventArgs.Empty);
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (allowClose || !isProcessing)
        {
            return;
        }

        if (IsNonInteractiveCloseReason(e.CloseReason))
        {
            allowClose = true;
            RequestCancellation();
            return;
        }

        e.Cancel = true;
        if (closeAfterCancellation)
        {
            return;
        }

        DialogResult choice = MessageBox.Show(
            this,
            "변환을 안전하게 취소한 뒤 창을 닫을까요?\n\n이미 완료된 파일은 그대로 유지되고, 처리 중인 임시 파일은 정리됩니다.",
            "변환 중",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2);

        if (choice == DialogResult.Yes)
        {
            closeAfterCancellation = true;
            RequestCancellation();
        }
    }

    internal static bool IsNonInteractiveCloseReason(CloseReason closeReason) =>
        closeReason is CloseReason.WindowsShutDown or CloseReason.TaskManagerClosing;

    private FormulaRemovalRequest CreateRequest()
    {
        string[] patterns = excludedSheetsTextBox.Text.Split(
            [',', ';'],
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return new FormulaRemovalRequest(
            inputWorkbooks.Items,
            outputDirectoryTextBox.Text.Trim(),
            decimal.ToInt32(startRowInput.Value),
            patterns);
    }

    private void AddWorkbooks(IEnumerable<string> paths)
    {
        InputAdditionResult addition = inputWorkbooks.Add(paths);
        if (addition.AddedCount > 0 && string.IsNullOrWhiteSpace(outputDirectoryTextBox.Text))
        {
            outputDirectoryTextBox.Text = Path.GetDirectoryName(inputWorkbooks.Items[0]) ?? string.Empty;
        }

        UpdateFileList();
        List<string> statusParts = [$"{addition.AddedCount:N0}개 추가"];
        if (addition.DuplicateCount > 0)
        {
            statusParts.Add($"중복 {addition.DuplicateCount:N0}개 제외");
        }

        if (addition.RejectedCount > 0)
        {
            statusParts.Add($"지원하지 않거나 찾을 수 없는 파일 {addition.RejectedCount:N0}개 제외");
        }

        SetStatus(string.Join(" · ", statusParts));
    }

    private void UpdateFileList()
    {
        fileListView.BeginUpdate();
        try
        {
            fileListView.Items.Clear();
            foreach (string path in inputWorkbooks.Items)
            {
                ListViewItem item = new(Path.GetFileName(path))
                {
                    Tag = path,
                    ToolTipText = path,
                };
                item.SubItems.Add(Path.GetDirectoryName(path) ?? string.Empty);
                fileListView.Items.Add(item);
            }
        }
        finally
        {
            fileListView.EndUpdate();
        }

        fileCountLabel.Text = $"{inputWorkbooks.Items.Count:N0}개 파일";
        ResizeFileColumns();
        UpdateActionState();
    }

    private void ResizeFileColumns()
    {
        if (fileListView.Columns.Count != 2 || fileListView.ClientSize.Width <= 0)
        {
            return;
        }

        int availableWidth = Math.Max(
            320,
            fileListView.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 24);
        int nameWidth = Math.Clamp((int)(availableWidth * 0.38), 180, 320);
        fileListView.Columns[0].Width = nameWidth;
        fileListView.Columns[1].Width = Math.Max(140, availableWidth - nameWidth);
    }

    private void SetProcessingState(bool processing)
    {
        addFilesButton.Enabled = !processing;
        fileListView.Enabled = !processing;
        outputDirectoryTextBox.Enabled = !processing;
        browseOutputButton.Enabled = !processing;
        startRowInput.Enabled = !processing;
        excludedSheetsTextBox.Enabled = !processing;
        cancelButton.Enabled = processing;
        UpdateActionState();
    }

    private void UpdateActionState()
    {
        bool canMutateFiles = !isProcessing;
        removeFilesButton.Enabled = canMutateFiles && fileListView.SelectedItems.Count > 0;
        clearFilesButton.Enabled = canMutateFiles && inputWorkbooks.Items.Count > 0;
        convertButton.Enabled = canMutateFiles && inputWorkbooks.Items.Count > 0 &&
            !string.IsNullOrWhiteSpace(outputDirectoryTextBox.Text);
    }

    private void UpdateProgress(FormulaRemovalProgress progress)
    {
        progressBar.Maximum = Math.Max(1, progress.TotalFiles);
        progressBar.Value = Math.Clamp(progress.CompletedFiles, 0, progressBar.Maximum);

        string fileName = Path.GetFileName(progress.InputPath);
        switch (progress.Stage)
        {
            case FormulaRemovalStage.StartingFile:
                SetStatus($"변환 중 · {progress.CompletedFiles + 1:N0}/{progress.TotalFiles:N0} · {fileName}");
                break;
            case FormulaRemovalStage.ProcessingWorksheet:
                SetStatus($"변환 중 · {progress.CompletedFiles + 1:N0}/{progress.TotalFiles:N0} · {fileName} · {progress.WorksheetName}");
                break;
            case FormulaRemovalStage.CompletedFile:
                SetStatus($"{progress.CompletedFiles:N0}/{progress.TotalFiles:N0}개 파일 처리 완료");
                break;
        }
    }

    private void ShowBatchResult(
        FormulaRemovalBatchResult result,
        bool canceled,
        string outputDirectory)
    {
        progressBar.Value = Math.Clamp(result.Files.Count, 0, progressBar.Maximum);

        bool hasWarnings = result.Files.Any(file => file.Warnings.Count > 0);
        string caption;
        MessageBoxIcon icon;
        if (canceled)
        {
            caption = "변환 취소됨";
            icon = MessageBoxIcon.Information;
            SetStatus($"취소됨 · 완료된 출력 {result.SucceededCount:N0}개");
        }
        else if (result.FailedCount > 0 || hasWarnings)
        {
            caption = "변환 결과 확인";
            icon = MessageBoxIcon.Warning;
            SetStatus($"완료 · 성공 {result.SucceededCount:N0}개 · 실패 {result.FailedCount:N0}개");
        }
        else
        {
            caption = "변환 완료";
            icon = MessageBoxIcon.Information;
            SetStatus($"완료 · {result.SucceededCount:N0}개 새 파일 저장");
        }

        MessageBox.Show(
            this,
            BatchSummaryFormatter.Format(
                result,
                canceled,
                outputDirectory,
                MaximumSummaryDetails),
            caption,
            MessageBoxButtons.OK,
            icon);
    }

    private void RequestCancellation()
    {
        if (!isProcessing || !conversionController.RequestCancellation())
        {
            return;
        }

        cancelButton.Enabled = false;
        SetStatus("취소 요청됨 · 처리 중인 임시 파일을 안전하게 정리하고 있습니다...");
    }

    private void ShowRequestError(string technicalMessage)
    {
        SetStatus("파일 목록, 출력 폴더 또는 시작 행을 확인하세요.");
        MessageBox.Show(
            this,
            "변환 설정을 확인해 주세요.\n\n원본 파일을 한 개 이상 추가하고 유효한 출력 폴더를 입력하세요.\n\n" +
                LimitText(technicalMessage, 500),
            "변환 설정 확인",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
    }

    private void SetStatus(string text)
    {
        statusLabel.Text = text;
        statusLabel.AccessibleDescription = text;
    }

    private string? GetInitialDirectory()
    {
        string outputDirectory = outputDirectoryTextBox.Text.Trim();
        if (Directory.Exists(outputDirectory))
        {
            return outputDirectory;
        }

        if (inputWorkbooks.Items.Count > 0)
        {
            string? inputDirectory = Path.GetDirectoryName(inputWorkbooks.Items[0]);
            if (Directory.Exists(inputDirectory))
            {
                return inputDirectory;
            }
        }

        return null;
    }

    private static bool TryGetDroppedFiles(DragEventArgs e, out string[] files)
    {
        files = [];
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) != true ||
            e.Data.GetData(DataFormats.FileDrop) is not string[] droppedFiles)
        {
            return false;
        }

        files = droppedFiles;
        return true;
    }

    private static bool IsSupportedWorkbookPath(string path)
    {
        string extension = Path.GetExtension(path);
        return File.Exists(path) &&
            (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xlsb", StringComparison.OrdinalIgnoreCase));
    }

    private static string LimitText(string text, int maximumLength)
    {
        return text.Length <= maximumLength ? text : text[..maximumLength] + "…";
    }
}
