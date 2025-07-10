using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WinFormsControl = System.Windows.Forms.Control;

namespace ExcelFormulaRemover
{
    public partial class MainForm : Form
    {
        private TextBox outputDirectoryTextBox = null!;
        private Button browseButton = null!;
        private ListBox fileListBox = null!;
        private Button processButton = null!;
        private ProgressBar progressBar = null!;
        private Label statusLabel = null!;
        private Panel dropPanel = null!;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Excel Formula Remover";
            this.Size = new System.Drawing.Size(600, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AllowDrop = true;

            // Drop Panel
            dropPanel = new Panel
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 120),
                BorderStyle = BorderStyle.FixedSingle,
                AllowDrop = true,
                BackColor = System.Drawing.Color.LightBlue
            };

            var dropLabel = new Label
            {
                Text = "Drag & Drop Excel Files Here\n(.xlsx, .xlsm files)",
                Location = new System.Drawing.Point(150, 40),
                Size = new System.Drawing.Size(240, 40),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold)
            };

            dropPanel.Controls.Add(dropLabel);
            dropPanel.DragEnter += DropPanel_DragEnter;
            dropPanel.DragDrop += DropPanel_DragDrop;

            // Output Directory Selection
            var outputLabel = new Label
            {
                Text = "Output Directory:",
                Location = new System.Drawing.Point(20, 160),
                Size = new System.Drawing.Size(100, 20)
            };

            outputDirectoryTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 158),
                Size = new System.Drawing.Size(350, 20),
                ReadOnly = true
            };

            browseButton = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(480, 156),
                Size = new System.Drawing.Size(80, 25)
            };
            browseButton.Click += BrowseButton_Click;

            // File List
            var filesLabel = new Label
            {
                Text = "Files to Process:",
                Location = new System.Drawing.Point(20, 200),
                Size = new System.Drawing.Size(100, 20)
            };

            fileListBox = new ListBox
            {
                Location = new System.Drawing.Point(20, 220),
                Size = new System.Drawing.Size(440, 120),
                AllowDrop = true
            };
            fileListBox.DragEnter += DropPanel_DragEnter;
            fileListBox.DragDrop += DropPanel_DragDrop;

            var clearButton = new Button
            {
                Text = "Clear",
                Location = new System.Drawing.Point(480, 220),
                Size = new System.Drawing.Size(80, 25)
            };
            clearButton.Click += (s, e) => fileListBox.Items.Clear();

            // Process Button
            processButton = new Button
            {
                Text = "Remove Formulas",
                Location = new System.Drawing.Point(20, 360),
                Size = new System.Drawing.Size(120, 30),
                Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Green,
                ForeColor = System.Drawing.Color.White,
                Enabled = false
            };
            processButton.Click += ProcessButton_Click;

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(160, 365),
                Size = new System.Drawing.Size(300, 20),
                Visible = false
            };

            // Status Label
            statusLabel = new Label
            {
                Text = "Ready - Drop files and select output directory",
                Location = new System.Drawing.Point(20, 410),
                Size = new System.Drawing.Size(540, 20),
                ForeColor = System.Drawing.Color.Blue
            };

            // Add all controls
            this.Controls.AddRange(new WinFormsControl[] {
                dropPanel, outputLabel, outputDirectoryTextBox, browseButton,
                filesLabel, fileListBox, clearButton, processButton,
                progressBar, statusLabel
            });

            UpdateUI();
        }

        private void DropPanel_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
                if (files.Any(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                                   f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                    dropPanel.BackColor = System.Drawing.Color.LightGreen;
                }
            }
        }

        private void DropPanel_DragDrop(object? sender, DragEventArgs e)
        {
            dropPanel.BackColor = System.Drawing.Color.LightBlue;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
                var excelFiles = files.Where(f => 
                    f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                    f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

                foreach (var file in excelFiles)
                {
                    if (!fileListBox.Items.Contains(file))
                    {
                        fileListBox.Items.Add(file);
                    }
                }

                statusLabel.Text = $"Added {excelFiles.Count()} Excel files";
                UpdateUI();
            }
        }

        private void BrowseButton_Click(object? sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select output directory for processed files";
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    outputDirectoryTextBox.Text = dialog.SelectedPath;
                    UpdateUI();
                }
            }
        }

        private void UpdateUI()
        {
            processButton.Enabled = fileListBox.Items.Count > 0 && 
                                   !string.IsNullOrEmpty(outputDirectoryTextBox.Text);
        }

        private async void ProcessButton_Click(object? sender, EventArgs e)
        {
            try
            {
                processButton.Enabled = false;
                progressBar.Visible = true;
                progressBar.Maximum = fileListBox.Items.Count;
                progressBar.Value = 0;

                var outputDir = outputDirectoryTextBox.Text;
                var files = fileListBox.Items.Cast<string>().ToList();

                statusLabel.Text = "Processing files...";

                await Task.Run(() =>
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        var inputFile = files[i];
                        var fileName = Path.GetFileNameWithoutExtension(inputFile);
                        var extension = Path.GetExtension(inputFile);
                        var outputFile = Path.Combine(outputDir, $"{fileName}_NoFormulas{extension}");

                        try
                        {
                            this.Invoke(new Action(() =>
                            {
                                statusLabel.Text = $"Processing: {Path.GetFileName(inputFile)}";
                            }));

                            ConvertFormulasToValues(inputFile, outputFile);

                            this.Invoke(new Action(() =>
                            {
                                progressBar.Value = i + 1;
                            }));
                        }
                        catch (Exception ex)
                        {
                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show($"Error processing {Path.GetFileName(inputFile)}:\n{ex.Message}",
                                              "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }));
                        }
                    }
                });

                statusLabel.Text = $"Completed! Processed {files.Count} files.";
                MessageBox.Show($"Successfully processed {files.Count} files!\nOutput directory: {outputDir}",
                              "Process Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                processButton.Enabled = true;
                progressBar.Visible = false;
                UpdateUI();
            }
        }

        private void ConvertFormulasToValues(string inputFile, string outputFile)
        {
            // Copy file first to preserve original
            File.Copy(inputFile, outputFile, true);

            using var document = SpreadsheetDocument.Open(outputFile, true);
            
            foreach (var worksheetPart in document.WorkbookPart!.WorksheetParts)
            {
                var worksheet = worksheetPart.Worksheet;
                var cells = worksheet.Descendants<Cell>().Where(c => c.CellFormula != null).ToList();

                foreach (var cell in cells)
                {
                    try
                    {
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

            document.Save();
        }
    }

    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}