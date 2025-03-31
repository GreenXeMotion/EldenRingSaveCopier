using System.ComponentModel;
using System.Text;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using EldenRingSaveCopy.Saves.Model;
using EldenRingSaveCopy.Backup;

namespace EldenRingSaveCopy
{
    public partial class MainForm : Form
    {
        private FileManager _fileManager;
        private readonly BindingList<ISaveGame> sourceSaveGames = new BindingList<ISaveGame>();
        private readonly BindingList<ISaveGame> targetSaveGames = new BindingList<ISaveGame>();
        private ISaveGame selectedSourceSave = new NullSaveGame();
        private ISaveGame selectedTargetSave = new NullSaveGame();
        private readonly IBackupManager _backupManager;
        private const string DEFAULT_ERROR_MESSAGE = "An unexpected error occurred";

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        private const int MESSAGE_ERROR = 0;
        private const int MESSAGE_INFO = 1;
        private const int MESSAGE_SUCCESS = 2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd,
                 int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private readonly ToolTip _toolTip = new ToolTip();
        private readonly StatusStrip _statusStrip = new StatusStrip();
        private readonly ToolStripStatusLabel _statusLabel = new ToolStripStatusLabel();

        private void MainForm_MouseDown(object? sender,
        System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && sender != null)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private string? FindSaveFile()
        {
            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string[] saveFiles = Directory.GetFiles(exePath, "ER0000.sl2")
                                          .Concat(Directory.GetFiles(exePath, "ER0000.co2"))
                                          .ToArray();
            return saveFiles.Length > 0 ? saveFiles[0] : null;
        }

        private bool LoadSaveFile(string path, bool isSource)
        {
            try
            {
                var fileBytes = File.ReadAllBytes(path);
                InitializeFileManager(path, fileBytes, isSource);
                LoadSaveSlots(fileBytes, isSource);
                return UpdateUIAfterLoad(isSource);
            }
            catch (Exception ex)
            {
                HandleLoadError(ex, isSource);
                return false;
            }
        }

        private void InitializeFileManager(string path, byte[] fileBytes, bool isSource)
        {
            if (isSource)
            {
                _fileManager.SourcePath = path;
                _fileManager.SourceFile = fileBytes;
                sourceFilePath.Text = path;
            }
            else
            {
                _fileManager.TargetPath = path;
                _fileManager.TargetFile = fileBytes;
                targetFilePath.Text = path;
            }
        }

        private void LoadSaveSlots(byte[] fileBytes, bool isSource)
        {
            var saveGames = isSource ? sourceSaveGames : targetSaveGames;
            saveGames.Clear();

            for (int i = 0; i < 10; i++)
            {
                var newSave = new SaveGame();
                newSave.LoadData(fileBytes, i);

                if (!isSource || newSave.Active)
                {
                    if (!isSource && !newSave.Active)
                    {
                        newSave.CharacterName = $"Slot {i + 1}";
                    }
                    saveGames.Add(newSave);
                }
            }
        }

        private bool UpdateUIAfterLoad(bool isSource)
        {
            var saveGames = isSource ? sourceSaveGames : targetSaveGames;
            var comboBox = isSource ? fromSaveSlot : toSaveSlot;

            if (saveGames.Count > 0)
            {
                comboBox.SelectedIndex = 0;
                if (isSource)
                {
                    selectedSourceSave = comboBox.SelectedItem as ISaveGame ?? new NullSaveGame();
                    showAdditionalInfoMessage(MESSAGE_INFO, "Source savegame file loaded successfully.");
                }
                else
                {
                    selectedTargetSave = comboBox.SelectedItem as ISaveGame ?? new NullSaveGame();
                    showAdditionalInfoMessage(MESSAGE_INFO, "Target savegame file loaded successfully.");
                }
                return true;
            }
            else if (isSource)
            {
                showAdditionalInfoMessage(MESSAGE_ERROR, "No active save slots found in source file.");
            }
            return false;
        }

        private void HandleLoadError(Exception ex, bool isSource)
        {
            var fileType = isSource ? "source" : "target";
            LogError($"Failed to load {fileType} savegame file", ex);
            showAdditionalInfoMessage(MESSAGE_ERROR, $"Failed to load {fileType} savegame file: {ex.Message}");

            if (isSource)
                sourceFilePath.Text = "Failed to load";
            else
                targetFilePath.Text = "Failed to load";
        }

        public MainForm()
        {
            InitializeComponent();
            _fileManager = new FileManager();
            _backupManager = new BackupManager();  // Changed from DefaultBackupManager to BackupManager

            // Configure tooltips
            _toolTip.SetToolTip(fromSaveSlot, "Select the character you want to copy from");
            _toolTip.SetToolTip(toSaveSlot, "Select the slot where you want to copy the character to");
            _toolTip.SetToolTip(copyButton, "Copy the selected character to the destination slot");

            // Add minimize button
            Button minimizeButton = new Button
            {
                Size = new Size(25, 25),
                Text = "_",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            minimizeButton.Click += (s, e) => this.WindowState = FormWindowState.Minimized;
            minimizeButton.Location = new Point(exitButton.Location.X - 30, exitButton.Location.Y);
            this.Controls.Add(minimizeButton);

            // Configure status strip
            _statusStrip.BackColor = Color.FromArgb(32, 32, 32);
            _statusStrip.ForeColor = Color.White;
            _statusLabel.Spring = true;
            _statusStrip.Items.Add(_statusLabel);
            this.Controls.Add(_statusStrip);

            // Auto-detect and load .sl2 or .co2 file if present
            string? savePath = FindSaveFile();
            if (savePath != null)
            {
                LoadSaveFile(savePath, true);
            }
            else
            {
                UpdateStatus("Select Source and Destination files and characters", MessageType.Info);
            }
        }

        private void UpdateStatus(string message, MessageType type)
        {
            _statusLabel.Text = message;
            _statusLabel.ForeColor = type switch
            {
                MessageType.Error => Color.DarkOrange,
                MessageType.Success => Color.Gold,
                _ => Color.White
            };
            showAdditionalInfoMessage((int)type, message);
        }

        private enum MessageType
        {
            Error = MESSAGE_ERROR,
            Info = MESSAGE_INFO,
            Success = MESSAGE_SUCCESS
        }

        // Tries to read the current windows user name and if it suceeds at it, it uses it in the default Elden Ring savefile location
        private void setCurrentUserDirectory(ref OpenFileDialog currentDialog)
        {
            string? nameDirectory = null;
            try
            {
                nameDirectory = "C:\\Users\\" + Environment.UserName + "\\AppData\\Roaming\\EldenRing";
                currentDialog.InitialDirectory = nameDirectory;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void sourceFileBrowse(object sender, EventArgs e)
        {
            sourceSaveGames.Clear();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            setCurrentUserDirectory(ref openFileDialog);
            openFileDialog.Filter = "Elden Ring Save File |ER0000.sl2|Elden Ring Coop Save File |ER0000.co2";
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                _fileManager.SourcePath = openFileDialog.FileName;
                try
                {
                    _fileManager.SourceFile = File.ReadAllBytes(_fileManager.SourcePath);
                    sourceFilePath.Text = _fileManager.SourcePath;

                    for (int i = 0; i < 10; i++)
                    {
                        var newSave = new SaveGame();
                        newSave.LoadData(_fileManager.SourceFile, i);
                        if(newSave.Active)
                        {
                            sourceSaveGames.Add(newSave);
                        }
                    }

                    if(sourceSaveGames.Count > 0)
                    {
                        fromSaveSlot.SelectedIndex = 0;
                        this.selectedSourceSave = fromSaveSlot.SelectedItem as ISaveGame ?? new NullSaveGame();
                        showAdditionalInfoMessage(MESSAGE_INFO, "Source savegame file loaded correctly.");
                    }
                }
                catch (IOException)
                {
                    sourceFilePath.Text = "Failed to load";
                    showAdditionalInfoMessage(MESSAGE_ERROR, "Source savegame file failed to load.");
                }
            }
            CheckButtonState();
        }

        private void targetButtonBrowse(object sender, EventArgs e)
        {
            targetSaveGames.Clear();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            setCurrentUserDirectory(ref openFileDialog);
            openFileDialog.Filter = "Elden Ring Save File |ER0000.sl2|Elden Ring Coop Save File |ER0000.co2";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                _fileManager.TargetPath = openFileDialog.FileName;
                try
                {
                    _fileManager.TargetFile = File.ReadAllBytes(_fileManager.TargetPath);
                    targetFilePath.Text = _fileManager.TargetPath;

                    for (int i = 0; i < 10; i++)
                    {
                        var newSave = new SaveGame();
                        newSave.LoadData(_fileManager.TargetFile, i);
                        if (!newSave.Active)
                        {
                            newSave.CharacterName = $"Slot {i + 1}";
                        }
                        targetSaveGames.Add(newSave);
                    }
                    if (targetSaveGames.Count > 0)
                    {
                        toSaveSlot.SelectedIndex = 0;
                        this.selectedTargetSave = toSaveSlot.SelectedItem as ISaveGame ?? new NullSaveGame();
                        showAdditionalInfoMessage(MESSAGE_INFO, "Destination savegame file loaded correctly.");
                    }
                }
                catch (IOException)
                {
                    sourceFilePath.Text = "Failed to load";
                    showAdditionalInfoMessage(MESSAGE_ERROR, "Destination savegame file failed to load.");
                }
            }
            CheckButtonState();
        }

        private void CheckButtonState()
        {
            if (_fileManager?.SourceFile != null && _fileManager.TargetFile != null &&
                _fileManager.SourceFile.Length > 0 && _fileManager.TargetFile.Length > 0 &&
                _fileManager.SourcePath != _fileManager.TargetPath &&
                selectedSourceSave?.Id != Guid.Empty && selectedTargetSave?.Id != Guid.Empty)
            {
                copyButton.Enabled = true;
                copyButton.BackColor = Color.Goldenrod;
                copyButton.Text =
                    "Copy source character "
                    + (this.selectedSourceSave?.CharacterName?.Contains("Slot ") == true ? this.selectedSourceSave.CharacterName :
                    this.selectedSourceSave?.CharacterName?.Split('\0')[0] ?? "Unknown")
                    + (this.selectedTargetSave?.CharacterName?.Contains("Slot ") == true ? " on destination file " + this.selectedTargetSave.CharacterName :
                    " over destination character " + (this.selectedTargetSave?.CharacterName?.Split('\0')[0] ?? "Unknown"));
            }
            else
            {
                copyButton.Enabled = false;
                copyButton.BackColor = Color.DarkOrange;
                copyButton.Text = "Select Source and Destination file and characters";
            }
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            titleBar.MouseDown += MainForm_MouseDown;

            fromSaveSlot.DisplayMember = "CharacterName";
            fromSaveSlot.DataSource = new BindingSource() { DataSource = this.sourceSaveGames }.DataSource;
            toSaveSlot.DisplayMember = "CharacterName";
            toSaveSlot.DataSource = new BindingSource() { DataSource = this.targetSaveGames }.DataSource;
        }

        private void fromSaveSlot_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ISaveGame save)
            {
                this.selectedSourceSave = save;
                CheckButtonState();
            }
        }

        private void toSaveSlot_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ISaveGame save)
            {
                this.selectedTargetSave = save;
                CheckButtonState();
            }
        }

        private void CreateFileBackup(string path, byte[] file)
        {
            var backupPath = _backupManager.GetNextBackupPath(path);
            File.WriteAllBytes(backupPath, file);
            UpdateStatus($"Backup created at {backupPath}", MessageType.Info);
        }

        private int SlotStartIndex(SaveGame save)
        {
            return (SaveGame.SLOT_START_INDEX + (save.Index * 0x10) + (save.Index * SaveGame.SLOT_LENGTH));
        }

        private int HeaderStartIndex(SaveGame save)
        {
            return (SaveGame.SAVE_HEADER_START_INDEX + (save.Index * SaveGame.SAVE_HEADER_LENGTH));
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            copyButton.Enabled = false;
            try
            {
                CreateFileBackup(_fileManager.TargetPath, _fileManager.TargetFile);

                var sourceSave = (SaveGame)this.selectedSourceSave;
                var targetSave = (SaveGame)this.selectedTargetSave;
                byte[] newSave = new byte[_fileManager.TargetFile.Length];

                //Create temp working file
                Array.Copy(_fileManager.TargetFile, newSave, _fileManager.TargetFile.Length);

                //Replace steam id in the source save with the steam id of the target save
                foreach (int idLocation in sourceSave.SaveData.StartingIndex(_fileManager.SourceID))
                {
                    Array.Copy(_fileManager.TargetID, 0, sourceSave.SaveData, idLocation, _fileManager.TargetID.Length);
                }

                //Copy source save slot to target save slot in temp file
                Array.Copy(sourceSave.SaveData, 0, newSave, SlotStartIndex(targetSave), SaveGame.SLOT_LENGTH);

                //Copy save header to temp file
                Array.Copy(sourceSave.HeaderData, 0, newSave, HeaderStartIndex(targetSave), SaveGame.SAVE_HEADER_LENGTH);

                //Mark target slot as active
                newSave[SaveGame.CHAR_ACTIVE_STATUS_START_INDEX + targetSave.Index] = 0x01;

                //Calculate checksums
                using (var md5 = MD5.Create())
                {
                    //Get slot checksum
                    md5.ComputeHash(sourceSave.SaveData);
                    //Write checksum to temp target file
                    if (md5.Hash != null)
                        Array.Copy(md5.Hash, 0, newSave, SlotStartIndex(targetSave) - 0x10, 0x10);
                    else
                        throw new InvalidOperationException("Failed to compute MD5 hash");
                    //get header checksum
                    md5.ComputeHash(newSave.Skip(SaveGame.SAVE_HEADERS_SECTION_START_INDEX).Take(SaveGame.SAVE_HEADERS_SECTION_LENGTH).ToArray());
                    //Write headers checksum
                    Array.Copy(md5.Hash, 0, newSave, SaveGame.SAVE_HEADERS_SECTION_START_INDEX - 0x10, 0x10);
                }

                //Write temp file to target file
                File.WriteAllBytes(_fileManager.TargetPath, newSave);

                //Delete old backup file to avoid corrupt save error
                File.Delete(_fileManager.TargetPath + ".bak");

                //Copy working file to source file to ensure each character is written to target file in the event of multiple characters being copied.
                Array.Copy(newSave, _fileManager.TargetFile, newSave.Length);

                this.targetSaveGames.RemoveAt(targetSave.Index);
                this.targetSaveGames.Insert(sourceSave.Index, sourceSave);
                toSaveSlot.SelectedIndex = targetSave.Index;

                //Indicate successful copy
                UpdateStatus("Copy successful! Ensure the ER0000.bak file has been deleted from save folder prior to loading.", MessageType.Success);
                copyButton.Text = "Copy Successful!";
                copyButton.BackColor = Color.Gold;
            }
            catch (Exception ex)
            {
                UpdateStatus($"Copy failed: {ex.Message}", MessageType.Error);
                copyButton.Text = "Copy Failed!";
                copyButton.BackColor = Color.DarkOrange;

                byte[] err = Encoding.Default.GetBytes(ex.Message);
                File.WriteAllBytes(_fileManager.TargetPath.Replace("ER0000.sl2", "Error.log"), err);
            }
        }

        private void showAdditionalInfoMessage(int type, string message)
        {
            additionalInfoLabel.Text = message;
            switch (type)
            {
                case MESSAGE_ERROR:
                    additionalInfoLabel.ForeColor = Color.DarkOrange;
                    break;
                case MESSAGE_INFO:
                    additionalInfoLabel.ForeColor = Color.White;
                    break;
                case MESSAGE_SUCCESS:
                    additionalInfoLabel.ForeColor = Color.Gold;
                    break;
                default:
                    additionalInfoLabel.ForeColor= Color.White;
                    break;

            }
        }

        private void exitButtonClick(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ShowError(string message, Exception? ex = null)
        {
            var errorMessage = ex != null ? $"{message}: {ex.Message}" : message;
            UpdateStatus(errorMessage, MessageType.Error);
            LogError(errorMessage, ex);
        }

        private void LogError(string message, Exception? ex = null)
        {
            var baseDir = _fileManager?.TargetPath != null ? Path.GetDirectoryName(_fileManager.TargetPath) : AppDomain.CurrentDomain.BaseDirectory;
            var logPath = Path.Combine(baseDir ?? AppDomain.CurrentDomain.BaseDirectory, "error.log");
            var logMessage = $"[{DateTime.Now}] {message}\n";
            if (ex != null)
                logMessage += $"Exception: {ex}\n";
            File.AppendAllText(logPath, logMessage);
        }

        private bool ExecuteWithErrorHandling(Action action, string successMessage, string? errorMessage = null)
        {
            try
            {
                action();
                UpdateStatus(successMessage, MessageType.Success);
                return true;
            }
            catch (Exception ex)
            {
                ShowError(errorMessage ?? DEFAULT_ERROR_MESSAGE, ex);
                return false;
            }
        }
    }
}
