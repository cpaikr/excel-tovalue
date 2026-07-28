// <copyright file="MainForm.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace ExcelFormulaRemover
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using WinFormsControl = System.Windows.Forms.Control;

    public partial class MainForm : Form
    {
        private TextBox outputDirectoryTextBox = null!;
        private Button browseButton = null!;
        private ListBox fileListBox = null!;
        private Button processButton = null!;
        private Button stopButton = null!;
        private ProgressBar progressBar = null!;
        private Label statusLabel = null!;
        private Panel dropPanel = null!;
        private Label startRowLabel = null!;
        private TextBox startRowTextBox = null!;
        private Label excludePatternLabel = null!;
        private TextBox excludePatternTextBox = null!;
        private ToolTip excludePatternToolTip = null!;
        private CancellationTokenSource? cancellationTokenSource;
        private bool isProcessing;

        public MainForm()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "엑셀 수식 제거기";
            this.Size = new System.Drawing.Size(600, 560);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AllowDrop = true;

            // Drop Panel
            this.dropPanel = new Panel
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 120),
                BorderStyle = BorderStyle.FixedSingle,
                AllowDrop = true,
                BackColor = System.Drawing.Color.LightBlue,
            };

            var dropLabel = new Label
            {
                Text = "엑셀 파일을 여기에 드래그하세요\n(.xlsx, .xlsm 파일)",
                Location = new System.Drawing.Point(150, 40),
                Size = new System.Drawing.Size(240, 40),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
            };

            this.dropPanel.Controls.Add(dropLabel);
            this.dropPanel.DragEnter += this.DropPanel_DragEnter;
            this.dropPanel.DragDrop += this.DropPanel_DragDrop;

            // Output Directory Selection
            var outputLabel = new Label
            {
                Text = "출력 폴더:",
                Location = new System.Drawing.Point(20, 160),
                Size = new System.Drawing.Size(100, 20),
            };

            this.outputDirectoryTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 158),
                Size = new System.Drawing.Size(350, 20),
                ReadOnly = true,
            };

            this.browseButton = new Button
            {
                Text = "찾아보기...",
                Location = new System.Drawing.Point(480, 156),
                Size = new System.Drawing.Size(80, 25),
            };
            this.browseButton.Click += this.BrowseButton_Click;

            // File List
            var filesLabel = new Label
            {
                Text = "처리할 파일:",
                Location = new System.Drawing.Point(20, 200),
                Size = new System.Drawing.Size(100, 20),
            };

            this.fileListBox = new ListBox
            {
                Location = new System.Drawing.Point(20, 220),
                Size = new System.Drawing.Size(440, 120),
                AllowDrop = true,
            };
            this.fileListBox.DragEnter += this.DropPanel_DragEnter;
            this.fileListBox.DragDrop += this.DropPanel_DragDrop;

            var clearButton = new Button
            {
                Text = "지우기",
                Location = new System.Drawing.Point(480, 220),
                Size = new System.Drawing.Size(80, 25),
            };
            clearButton.Click += (s, e) => this.fileListBox.Items.Clear();

            // Start Row Input
            this.startRowLabel = new Label
            {
                Text = "수식제거 시작행(포함):",
                Location = new System.Drawing.Point(20, 347),
                Size = new System.Drawing.Size(150, 20),
            };

            this.startRowTextBox = new TextBox
            {
                Location = new System.Drawing.Point(175, 345),
                Size = new System.Drawing.Size(50, 20),
                Text = "1",
                MaxLength = 3,
            };

            // Exclude Pattern Input
            this.excludePatternLabel = new Label
            {
                Text = "다음 텍스트를 포함하는 시트는 수식 유지:",
                Location = new System.Drawing.Point(20, 375),
                Size = new System.Drawing.Size(270, 20),
            };

            this.excludePatternTextBox = new TextBox
            {
                Location = new System.Drawing.Point(295, 373),
                Size = new System.Drawing.Size(100, 20),
                PlaceholderText = "TF, 변동성",
                MaxLength = 100,
            };

            // Tooltip for exclude pattern
            this.excludePatternToolTip = new ToolTip();
            this.excludePatternToolTip.SetToolTip(
                this.excludePatternTextBox,
                "입력한 텍스트를 포함하는 시트 이름은 수식을 유지합니다\n(여러 패턴은 콤마로 구분: TF, 변동성)\n(예: TF 입력시 TF_Summary, Data_TF 등의 시트는 수식 유지)");

            // Process Button
            this.processButton = new Button
            {
                Text = "수식 제거",
                Location = new System.Drawing.Point(20, 405),
                Size = new System.Drawing.Size(120, 30),
                Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Green,
                ForeColor = System.Drawing.Color.White,
                Enabled = false,
            };
            this.processButton.Click += this.ProcessButton_Click;

            // Progress Bar
            this.progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(160, 410),
                Size = new System.Drawing.Size(280, 20),
                Visible = false,
            };

            // Stop Button
            this.stopButton = new Button
            {
                Text = "중지",
                Location = new System.Drawing.Point(450, 410),
                Size = new System.Drawing.Size(90, 20),
                BackColor = System.Drawing.Color.Red,
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold),
                Visible = false,
            };
            this.stopButton.Click += this.StopButton_Click;

            // Status Label
            this.statusLabel = new Label
            {
                Text = "준비 완료 - 파일을 드롭하고 출력 폴더를 선택하세요",
                Location = new System.Drawing.Point(20, 455),
                Size = new System.Drawing.Size(540, 20),
                ForeColor = System.Drawing.Color.Blue,
            };

            // Add all controls
            this.Controls.AddRange(
                this.dropPanel, outputLabel, this.outputDirectoryTextBox, this.browseButton,
                filesLabel, this.fileListBox, clearButton, this.startRowLabel, this.startRowTextBox,
                this.excludePatternLabel, this.excludePatternTextBox,
                this.processButton, this.progressBar, this.stopButton, this.statusLabel);

            this.UpdateUI();
        }

        private void DropPanel_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (files != null && Array.Exists(files, f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                                   f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                    this.dropPanel.BackColor = System.Drawing.Color.LightGreen;
                }
            }
        }

        private void DropPanel_DragDrop(object? sender, DragEventArgs e)
        {
            this.dropPanel.BackColor = System.Drawing.Color.LightBlue;

            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (files != null)
                {
                    var excelFiles = files.Where(f =>
                        f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

                    foreach (var file in excelFiles)
                    {
                        if (!this.fileListBox.Items.Contains(file))
                        {
                            this.fileListBox.Items.Add(file);
                        }
                    }

                    this.statusLabel.Text = $"{excelFiles.Count()}개 엑셀 파일 추가됨";
                    this.UpdateUI();
                }
            }
        }

        private void BrowseButton_Click(object? sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "처리된 파일의 출력 폴더를 선택하세요";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.outputDirectoryTextBox.Text = dialog.SelectedPath;
                    this.UpdateUI();
                }
            }
        }

        private void UpdateUI()
        {
            this.processButton.Enabled = true;
        }

        private async void ProcessButton_Click(object? sender, EventArgs e)
        {
            // Validate prerequisites before processing
            if (string.IsNullOrEmpty(this.outputDirectoryTextBox.Text))
            {
                MessageBox.Show(
                    "출력 폴더를 선택해주세요!\n\n처리된 파일이 저장될 폴더를 지정해야 합니다.",
                    "출력 폴더 미선택",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                this.statusLabel.Text = "오류: 출력 폴더를 선택해주세요";
                this.statusLabel.ForeColor = System.Drawing.Color.Red;
                return;
            }

            if (this.fileListBox.Items.Count == 0)
            {
                MessageBox.Show(
                    "처리할 엑셀 파일을 추가해주세요!\n\n파일을 드래그하여 추가하거나 선택해주세요.",
                    "처리할 파일 없음",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                this.statusLabel.Text = "오류: 처리할 파일을 추가해주세요";
                this.statusLabel.ForeColor = System.Drawing.Color.Red;
                return;
            }

            // Validate start row input
            if (!this.ValidateStartRow(out int startRow))
            {
                this.statusLabel.Text = "오류: 시작행 입력을 확인해주세요";
                this.statusLabel.ForeColor = System.Drawing.Color.Red;
                return;
            }

            // Reset status label color to normal
            this.statusLabel.ForeColor = System.Drawing.Color.Blue;

            // Declare variables outside try block for scope access
            var outputDir = this.outputDirectoryTextBox.Text;
            var files = this.fileListBox.Items.Cast<string>().ToList();
            var completedFiles = 0;
            var currentOutputFile = string.Empty;

            try
            {
                // Initialize cancellation token and UI state
                this.cancellationTokenSource = new CancellationTokenSource();
                this.isProcessing = true;

                this.processButton.Enabled = false;
                this.stopButton.Visible = true;
                this.progressBar.Visible = true;
                this.progressBar.Style = ProgressBarStyle.Marquee;
                this.progressBar.MarqueeAnimationSpeed = 30;

                this.statusLabel.Text = "파일 처리 중...";

                await Task.Run(
                    () =>
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        // Check for cancellation before processing each file
                        this.cancellationTokenSource.Token.ThrowIfCancellationRequested();

                        var inputFile = files[i];
                        var fileName = Path.GetFileNameWithoutExtension(inputFile);
                        var extension = Path.GetExtension(inputFile);
                        currentOutputFile = Path.Combine(outputDir, $"{fileName}_수식제거{extension}");

                        try
                        {
                            var excludePattern = this.excludePatternTextBox.Text.Trim();
                            this.ConvertFormulasToValues(inputFile, currentOutputFile, startRow, excludePattern, this.cancellationTokenSource.Token);

                            completedFiles++;
                        }
                        catch (OperationCanceledException)
                        {
                            // Clean up partial file if it exists
                            try
                            {
                                if (File.Exists(currentOutputFile))
                                {
                                    File.Delete(currentOutputFile);
                                }
                            }
                            catch
                            { /* Ignore cleanup errors */
                            }

                            throw; // Re-throw to be caught by outer handler
                        }
                        catch (Exception ex)
                        {
                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show(
                                    $"{Path.GetFileName(inputFile)} 처리 중 오류:\n{ex.Message}",
                                    "처리 오류",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                            }));
                        }
                    }
                }, this.cancellationTokenSource.Token);

                this.statusLabel.Text = $"완료! {completedFiles}개 파일 처리됨.";
                MessageBox.Show(
                    $"{completedFiles}개 파일 처리 완료!\n출력 폴더: {outputDir}\n\n" +
                              "★★★ 중요 ★★★\n" +
                              "반드시 산출 파일을 열어서 정상적으로 변환되었는지 확인하세요!\n" +
                              "수식이 값으로 올바르게 대체되었는지 검증이 필요합니다.",
                    "처리 완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                this.statusLabel.Text = $"처리 중지됨. {completedFiles}개 파일 완료.";
                MessageBox.Show(
                    $"사용자에 의해 처리가 중지되었습니다.\n완료된 파일: {completedFiles}개\n출력 폴더: {outputDir}",
                    "처리 중지",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"오류가 발생했습니다: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // Reset UI state
                this.isProcessing = false;
                this.processButton.Enabled = true;
                this.stopButton.Visible = false;
                this.progressBar.Visible = false;
                this.progressBar.Style = ProgressBarStyle.Continuous;

                // Dispose cancellation token source
                this.cancellationTokenSource?.Dispose();
                this.cancellationTokenSource = null;

                this.UpdateUI();
            }
        }

        private void StopButton_Click(object? sender, EventArgs e)
        {
            if (this.isProcessing && this.cancellationTokenSource != null)
            {
                this.statusLabel.Text = "중지 요청됨 - 처리를 안전하게 중단하는 중...";
                this.stopButton.Enabled = false;
                this.cancellationTokenSource.Cancel();
            }
        }

        private void ConvertFormulasToValues(string inputFile, string outputFile, int startingRow, string excludePattern, CancellationToken cancellationToken)
        {
            // Copy file first to preserve original
            File.Copy(inputFile, outputFile, true);

            using var document = SpreadsheetDocument.Open(outputFile, true);

            int totalFormulas = 0;
            int processedFormulas = 0;

            // Filter worksheets based on exclusion pattern
            var workbook = document.WorkbookPart?.Workbook ??
                throw new InvalidDataException("Workbook is missing its root element.");
            var sheets = workbook.Sheets?.Cast<Sheet>().ToList() ??
                throw new InvalidDataException("Workbook has no sheets collection.");

            // Get sheets to process (exclude those matching the pattern)
            var sheetsToProcess = sheets;
            var excludedSheets = new List<Sheet>();

            if (!string.IsNullOrWhiteSpace(excludePattern))
            {
                // Split by comma and trim each pattern
                var patterns = excludePattern.Split(',')
                    .Select(p => p.Trim())
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .ToList();

                excludedSheets = sheets.Where(s =>
                    patterns.Any(pattern => s.Name!.Value!.Contains(pattern, StringComparison.OrdinalIgnoreCase))).ToList();

                sheetsToProcess = sheets.Where(s =>
                    !patterns.Any(pattern => s.Name!.Value!.Contains(pattern, StringComparison.OrdinalIgnoreCase))).ToList();
            }

            // Get corresponding worksheet parts
            var worksheetPartsToProcess = sheetsToProcess.Select(sheet =>
                document.WorkbookPart.GetPartById(sheet.Id!.Value!) as WorksheetPart)
                .Where(part => part != null)
                .ToList();

            // First, count total formulas across filtered worksheets from starting row
            totalFormulas = worksheetPartsToProcess
                .Select(worksheetPart => worksheetPart!.Worksheet)
                .Sum(worksheet => worksheet!.Descendants<Cell>().Count(c =>
                    c.CellFormula != null &&
                    GetRowNumberFromCellReference(c.CellReference?.Value ?? string.Empty) >= startingRow));

            // Update status with sheet processing information
            var fileName = Path.GetFileName(inputFile);
            var statusMessage = excludedSheets.Count > 0
                ? $"처리 중: {fileName} - {sheetsToProcess.Count}개 시트 처리 중 ({excludedSheets.Count}개 시트 제외)"
                : $"처리 중: {fileName} - {sheetsToProcess.Count}개 시트 처리 중";

            this.Invoke(new Action(() =>
            {
                this.statusLabel.Text = statusMessage;
            }));

            // Now process the formulas
            var worksheets = worksheetPartsToProcess.Select(worksheetPart => worksheetPart!.Worksheet);
            foreach (var worksheet in worksheets)
            {
                // Check for cancellation before processing each worksheet
                cancellationToken.ThrowIfCancellationRequested();

                var cells = worksheet!.Descendants<Cell>().Where(c =>
                    c.CellFormula != null &&
                    GetRowNumberFromCellReference(c.CellReference?.Value ?? string.Empty) >= startingRow).ToList();

                foreach (var cell in cells)
                {
                    try
                    {
                        processedFormulas++;

                        // Check for cancellation every 100 formulas for faster response
                        if (processedFormulas % 100 == 0 || processedFormulas == totalFormulas)
                        {
                            // Check for cancellation
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // Update status every 100K formulas for performance
                        if (processedFormulas % 100000 == 0 || processedFormulas == totalFormulas)
                        {
                            this.Invoke(new Action(() =>
                            {
                                // Display counts in millions with one decimal place
                                double processedMillions = processedFormulas / 1000000.0;
                                double totalMillions = totalFormulas / 1000000.0;
                                this.statusLabel.Text = $"처리 중: {Path.GetFileName(inputFile)} - 수식 {processedMillions:F1}백만개 / {totalMillions:F1}백만개";
                            }));
                        }

                        // Keep the calculated value, remove the formula
                        var calculatedValue = cell.CellValue?.Text;
                        if (calculatedValue != null)
                        {
                            cell.CellFormula = null;
                        }
                    }
                    catch
                    {
                        // Skip problematic cells
                        continue;
                    }
                }
            }

            // Remove calculation chain to prevent corruption warnings
            // Since we've removed all formulas, the calculation chain is no longer needed
            try
            {
                var calculationChainPart = document.WorkbookPart!.CalculationChainPart;
                if (calculationChainPart != null)
                {
                    document.WorkbookPart.DeletePart(calculationChainPart);
                }
            }
            catch
            {
                // Ignore errors if calculation chain doesn't exist or can't be removed
            }

            document.Save();
        }

        private static int GetRowNumberFromCellReference(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return 0;
            }

            var match = Regex.Match(cellReference, @"\d+$");
            return match.Success ? int.Parse(match.Value, CultureInfo.InvariantCulture) : 0;
        }

        private bool ValidateStartRow(out int startRow)
        {
            startRow = 1;
            var input = this.startRowTextBox.Text.Trim();

            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show(
                    "수식제거 시작행을 입력해주세요!\n\n1부터 100 사이의 숫자를 입력하세요.",
                    "시작행 입력 필요",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (!int.TryParse(input, out startRow))
            {
                MessageBox.Show(
                    "올바른 숫자를 입력해주세요!\n\n1부터 100 사이의 정수를 입력하세요.",
                    "잘못된 입력",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (startRow < 1 || startRow > 100)
            {
                MessageBox.Show(
                    "시작행은 1부터 100 사이의 값이어야 합니다!\n\n현재 입력값: " + startRow,
                    "범위 오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
    }
}
