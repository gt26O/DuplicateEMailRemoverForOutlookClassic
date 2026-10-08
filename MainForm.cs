// Released under the MIT License. Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

// DISCLAIMER
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THIS CODE LIES SOLELY WITH THE USER.

// Created by Ryan Scott White on Nov. 2023, Updated 12/20/2024 (note: no AI used for coding as of 12/20/2024, may change with the future drafts)


using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using DuplicateEMailRemover.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace DuplicateEMailRemover
{
    public partial class MainForm : Form
    {
        // Enabling will allow emails to be acted on(like delete or moved).
        private readonly bool ALLOW_MOVES_DELETES = true;

        private int authorizedTabIndex = 0;
        private int totalFoldersToSearch = 0;
        private int totalItemsToSearch = 0;
        private readonly string logToPath = $"{Path.GetTempPath()}\\DuplicateEMailRemoverOutput.txt";
        private CancellationTokenSource? tokenSource;
        private static readonly Outlook.Application outlook = new();

        // If this is specified, then it will save each email to the specified folder.
        private string? filePathToSaveEmails;

        // Live progress + restore UI. These controls are created in Form1_Load
        // (not in the generated designer file) and added to the "Go" tab.
        private readonly System.Diagnostics.Stopwatch scanStopwatch = new();
        private readonly System.Windows.Forms.Timer progressTimer = new();
        private int lastProcessedCount;
        private string? lastBackupFolder;
        private Label lblCurrent = null!;
        private Label lblElapsed = null!;
        private Label lblRate = null!;
        private Label lblEta = null!;
        private Button btnRestore = null!;
        private ListView lvDeleted = null!;
        private CheckBox checkBoxMatchOnMessageId = null!;

        // Folders we should skip - these folders should not really be scanned for duplicate emails.
        private static readonly SortedSet<string> FoldersToSkip = [
            "Calendar",
            "Contacts",
            "Conflicts",
            "Conversation Action Settings",
            "Conversation History",
            "Drafts",
            "ExternalContacts",
            "Files",
            "Journal",
            "Local Failures",
            "Lync Contacts",
            "MeContact",
            "News Feed",
            "Notes",
            "Outbox",
            "PersonMetadata",
            "Quick Step Settings",
            "RSS Feeds",
            "Server Failures",
            "Suggested Contacts",
            "Sync Issues",
            "Tasks",
            "Yammer Root",
            "Your feeds",
            ];

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //tabControl1.Appearance = TabAppearance.FlatButtons;
            //tabControl1.ItemSize = new Size(0, 1);
            //tabControl1.SizeMode = TabSizeMode.Fixed;
            radioButtonLogOnly.Text = $"Log only: {logToPath}";

            if (!ALLOW_MOVES_DELETES)
            {
                Text += " (DEBUG MODE)";
            }

            BuildProgressAndRestoreUi();
            BuildMessageIdMatchOption();
        }

        // Adds a "Message-ID" matching option to the criteria group. The Internet
        // Message-ID uniquely identifies an email, so combining it with date +
        // sender + subject is the most reliable way to detect true duplicates.
        private void BuildMessageIdMatchOption()
        {
            checkBoxMatchOnMessageId = new CheckBox
            {
                AutoSize = true,
                Location = new Point(421, 116),
                Text = "Message-ID",
                UseVisualStyleBackColor = true
            };
            toolTip1.SetToolTip(checkBoxMatchOnMessageId,
                "Match on the email's unique Internet Message-ID header. " +
                "Best combined with Date Sent + Sender + Subject.");
            checkBoxMatchOnMessageId.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnMessageId);
        }

        // Adds the live progress details, the "deleted emails" list, and the
        // restore button to the Go tab at runtime (keeps the designer file simple).
        private void BuildProgressAndRestoreUi()
        {
            lblCurrent = new Label { Location = new Point(12, 268), AutoSize = true, Text = "Current folder: -" };
            lblElapsed = new Label { Location = new Point(12, 292), AutoSize = true, Text = "Elapsed: 00:00:00" };
            lblRate = new Label { Location = new Point(200, 292), AutoSize = true, Text = "Rate: 0 /s" };
            lblEta = new Label { Location = new Point(360, 292), AutoSize = true, Text = "ETA: -" };

            btnRestore = new Button
            {
                Location = new Point(410, 258),
                Size = new Size(220, 34),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Text = "Restore deleted emails…"
            };
            btnRestore.Click += btnRestore_Click;

            Label lblDeleted = new() { Location = new Point(12, 318), AutoSize = true, Text = "Deleted emails (this run):" };

            lvDeleted = new ListView
            {
                Location = new Point(12, 340),
                Size = new Size(616, 160),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            _ = lvDeleted.Columns.Add("Subject", 230);
            _ = lvDeleted.Columns.Add("Sender", 150);
            _ = lvDeleted.Columns.Add("Folder", 150);
            _ = lvDeleted.Columns.Add("Id", 180);
            _ = lvDeleted.Columns.Add("Deleted (UTC)", 140);

            tabPage6.Controls.Add(lblCurrent);
            tabPage6.Controls.Add(lblElapsed);
            tabPage6.Controls.Add(lblRate);
            tabPage6.Controls.Add(lblEta);
            tabPage6.Controls.Add(btnRestore);
            tabPage6.Controls.Add(lblDeleted);
            tabPage6.Controls.Add(lvDeleted);

            progressTimer.Interval = 1000;
            progressTimer.Tick += ProgressTimer_Tick;
        }

        private void ProgressTimer_Tick(object? sender, EventArgs e)
        {
            TimeSpan el = scanStopwatch.Elapsed;
            lblElapsed.Text = $"Elapsed: {el:hh\\:mm\\:ss}";

            double sec = el.TotalSeconds;
            double rate = sec > 0 ? lastProcessedCount / sec : 0;
            lblRate.Text = $"Rate: {rate:0.#} /s";

            if (rate > 0 && totalItemsToSearch > lastProcessedCount)
            {
                double remain = (totalItemsToSearch - lastProcessedCount) / rate;
                lblEta.Text = $"ETA: {TimeSpan.FromSeconds(remain):hh\\:mm\\:ss}";
            }
            else
            {
                lblEta.Text = "ETA: -";
            }
        }

        private void AddDeletedRow(DeletedEmailRecord r)
        {
            ListViewItem item = new(r.Subject);
            _ = item.SubItems.Add(string.IsNullOrEmpty(r.SenderName) ? r.SenderEmail : r.SenderName);
            _ = item.SubItems.Add(r.FolderPath);
            _ = item.SubItems.Add(r.Id);
            _ = item.SubItems.Add(r.DeletedUtc);
            _ = lvDeleted.Items.Add(item);
            item.EnsureVisible();
        }

        private void btnRestore_Click(object? sender, EventArgs e)
        {
            try
            {
                using RestoreForm rf = new(outlook);
                _ = rf.ShowDialog(this);
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            _ = MessageBox.Show("Avoid opening or closing Outlook while running this tool.");

            SuspendLayout();
            Outlook.Folders mailboxes = outlook.GetNamespace("MAPI").Folders;
            int folderCount = 0;
            //treeView1.BeginUpdate();
            foreach (Outlook.Folder f in mailboxes)
            {
                try
                {
                    int mailboxNameLen = f.FolderPath.Length + 1;
                    FolderTreeNode node = new(f);
                    node.Expand();
                    visitTopLevelFolders(f, node, mailboxNameLen, ref folderCount);
                    _ = treeView1.Nodes.Add(node);
                }
                catch (Exception ee)
                {
                    _ = MessageBox.Show($"""
                            There was an error with the mailbox {f.Name}-{f.AddressBookName}-{f.FolderPath}.
                            The error message was: {ee.Message}");
                            {ee.InnerException}
                            """);
                }

            }
            //treeView1.EndUpdate();
            ResumeLayout();

            // Make sure some folders were found.
            if (folderCount == 0)
            {
                _ = MessageBox.Show("No folders found. Please make sure you have office installed.");
                System.Windows.Forms.Application.Exit();
            }
            //MessageBox.Show("Form1_END");
        }

        private void visitTopLevelFolders(Outlook.Folder folder, TreeNode node, int mailboxNameLen, ref int folderCount)
        {
            foreach (Outlook.Folder f in folder.Folders)
            {
                string folderPath = f.FolderPath[mailboxNameLen..];
                if (FoldersToSkip.Contains(folderPath))
                {
                    continue;
                }

                FolderTreeNode childFolderNode = new(f);
                folderCount++;
                _ = node.Nodes.Add(childFolderNode);

                try
                {
                    visitSubFolder(f, childFolderNode, ref folderCount);
                }
                catch (Exception error)
                {
                    _ = MessageBox.Show($"""
                            Errors can sometimes happen if Outlook is opened/closed while 
                            $"this app is open.
                            There was an error with the mailbox {f.Name}-{f.AddressBookName}-{f.FolderPath}
                            The error message was: {error.Message}");
                            {error.InnerException}
                            """);
                    //Application.Exit();
                }
            }
        }


        private void visitSubFolder(Outlook.Folder folder, TreeNode node, ref int folderCount)
        {
            foreach (Outlook.Folder f in folder.Folders)
            {
                FolderTreeNode childFolderNode = new(f);
                folderCount++;
                _ = node.Nodes.Add(childFolderNode);
                visitSubFolder(f, childFolderNode, ref folderCount);
            }
        }

        private void btnTreeViewSelectNextPanel_Click(object sender, EventArgs e)
        {
            totalItemsToSearch = 0;
            totalFoldersToSearch = 0;

            try
            {
                foreach (TreeNode nd in treeView1.Nodes)
                {
                    CallRecursiveOnTreeNodeToFillOrderByListBox(nd);
                }
            }
            catch (Exception error)
            {
                _ = MessageBox.Show($"Errors can sometimes happen if Outlook is opened/closed while " +
                    $"this app is open.\n\n Details: {error.Message})");
                Application.Exit();

            }

            if (totalFoldersToSearch == 0)
            {
                _ = MessageBox.Show("There should be at least one folder selected.");
                return;
            }

            // If there is only one item selected then we don't need to do any sorting.
            else if (totalFoldersToSearch == 1)
            {
                authorizedTabIndex = 2;
                tabControl1.SelectedIndex = 2;
            }

            // there was at last two folders selected so we should proceed to the folder-ordering tab
            else
            {

                // Generally we want anything under Deleted Items last
                for (int i = listBoxOfFoldersToProcess.Items.Count - 1; i >= 0; i--)
                {
                    FolderTreeNode node = (FolderTreeNode)listBoxOfFoldersToProcess.Items[i];
                    if (Regex.IsMatch(node.Text, @"\\\\[^\\]+?\\Deleted Items(\\|$)"))
                    {
                        _ = listBoxOfFoldersToProcess.Items.Add(node);
                        listBoxOfFoldersToProcess.Items.RemoveAt(i);
                    }
                }

                authorizedTabIndex = 1;
                tabControl1.SelectedIndex = 1; // = (tabControl1.SelectedIndex + 1 < tabControl1.TabCount) ?   tabControl1.SelectedIndex + 1 : tabControl1.SelectedIndex;
            }

            txtItemsTotalCount.Text = totalItemsToSearch.ToString();
            txtFoldersTotalCount.Text = totalFoldersToSearch.ToString();
        }

        private void CallRecursiveOnTreeNodeToFillOrderByListBox(TreeNode nd)
        {
            if (nd.Checked)
            {
                _ = listBoxOfFoldersToProcess.Items.Add(nd);
                totalItemsToSearch += (nd as FolderTreeNode)?.OutlookFolder.Items.Count ?? 0;
                totalFoldersToSearch++;
            }

            _ = cboDestinationFolder.Items.Add(nd);

            foreach (TreeNode n in nd.Nodes)
            {
                CallRecursiveOnTreeNodeToFillOrderByListBox(n);
            }
        }

        private void treeView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Select the clicked node
                treeView1.SelectedNode = treeView1.GetNodeAt(e.X, e.Y);

                if (treeView1.SelectedNode != null)
                {
                    contextMenuStrip1.Show(treeView1, e.Location);
                }
            }
        }

        private void toolStripMenuItemCheckAllChildNodes_Click(object sender, EventArgs e)
        {
            CheckOrUncheckSubnodes(treeView1.SelectedNode, true);
        }

        private void toolStripMenuItemUnCheckAllChildNodes_Click(object sender, EventArgs e)
        {
            CheckOrUncheckSubnodes(treeView1.SelectedNode, false);
        }

        private void CheckOrUncheckSubnodes(TreeNode node, bool value)
        {
            node.Checked = value;
            foreach (TreeNode childNode in node.Nodes)
            {
                CheckOrUncheckSubnodes(childNode, value);
            }
        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxOfFoldersToProcess.SelectedItem == null)
            {
                return;
            }

            _ = listBoxOfFoldersToProcess.DoDragDrop(listBoxOfFoldersToProcess.SelectedItem, DragDropEffects.Move);
        }

        private void listBox1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            Point point = listBoxOfFoldersToProcess.PointToClient(new Point(e.X, e.Y));
            int index = listBoxOfFoldersToProcess.IndexFromPoint(point);
            if (index < 0)
            {
                index = listBoxOfFoldersToProcess.Items.Count - 1;
            }

            object? data = e.Data?.GetData(typeof(FolderTreeNode)) ?? null;
            if (data != null)
            {
                listBoxOfFoldersToProcess.Items.Remove(data);
                listBoxOfFoldersToProcess.Items.Insert(index, data);
            }
        }

        private void btnNextToOptions_Click(object sender, EventArgs e)
        {
            authorizedTabIndex = 2;
            tabControl1.SelectedIndex = 2;
        }

        private void radioButtonActionUpdateGUI_CheckedChanged(object sender, EventArgs e)
        {
            cboDestinationFolder.Enabled = radioButtonMoveToFolder.Checked | radioButtonCopyToFolder.Checked;
        }

        private void btnMoveToBeginProcessing_Click(object sender, EventArgs e)
        {
            // Lets make sure the destination folder is valid.
            if ((radioButtonMoveToFolder.Checked | radioButtonCopyToFolder.Checked)
                && cboDestinationFolder.SelectedIndex == -1)
            {
                _ = MessageBox.Show("Please select a destination folder.");
                return;
            }

            if (!checkBoxMatchOnBody.Checked && !checkBoxMatchOnHTMLBody.Checked
                && checkBoxMatchOnMessageId?.Checked != true)
            {
                DialogResult result = MessageBox.Show("Neither the \"Email Body\" or \"Body(HTLM)\" were selected.\n" +
                    "This can cause unintended non-duplicate messages to be moved/deleted.\n\n" +
                    "Are you sure you would like to continue? ", "Confirm", MessageBoxButtons.YesNo);

                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            if (radioButtonOpenEachInOutlook.Checked)
            {
                _ = MessageBox.Show("Please open Outlook and leave it open. Closing it will cause errors.");
            }


            authorizedTabIndex = 3;
            tabControl1.SelectedIndex = 3;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only needed if user somehow clicks back to a previous tab
            if (tabControl1.SelectedIndex == 0)
            {
                listBoxOfFoldersToProcess.Items.Clear();
            }
        }
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (ALLOW_MOVES_DELETES)
            {
                if (tabControl1.SelectedIndex != authorizedTabIndex)
                {
                    e.Cancel = true;
                }
            }
        }

        private void checkBoxMatchOn_CheckedChanged(object sender, EventArgs e)
        {

            int checkCount =
                  (checkBoxMatchOnMessageId?.Checked == true ? 1 : 0)
                + (checkBoxMatchOnSentOn.Checked ? 1 : 0)
                + (checkBoxMatchOnReceivedTime.Checked ? 1 : 0)
                + (checkBoxMatchOnLastModTime.Checked ? 1 : 0)
                + (checkBoxMatchOnSenderEmail.Checked ? 1 : 0)
                + (checkBoxMatchOnTo.Checked ? 1 : 0)
                + (checkBoxMatchOnCC.Checked ? 1 : 0)
                + (checkBoxMatchOnBCC.Checked ? 1 : 0)
                + (checkBoxMatchOnSubject.Checked ? 1 : 0)
                + (checkBoxMatchOnBody.Checked ? 1 : 0)
                + (checkBoxMatchOnHTMLBody.Checked ? 1 : 0)
                + (checkBoxMatchOnAttachment.Checked ? 1 : 0);

            btnMoveToBeginProcessing.Enabled = checkCount > 1;
        }

        // some help from here https://stackoverflow.com/a/18033198/2352507 for async and progress bars
        private async void btnStart_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            btnRestore.Enabled = false;
            btnStart.Text = "Running";
            lvDeleted.Items.Clear();
            lastProcessedCount = 0;
            lastBackupFolder = null;
            scanStopwatch.Restart();
            progressTimer.Start();
            try
            {
                Progress<int> folderProcessedCount = new(s => txtFoldersProcessed.Text = s.ToString());
                Progress<int> mailProcessedCount = new(s => { txtEmailsProcessed.Text = s.ToString(); lastProcessedCount = s; });
                Progress<int> duplicatesFound = new(s => txtDuplicatesFound.Text = s.ToString());
                Progress<int> nonMailItemsSkipped = new(s => txtNonMailItemsSkipped.Text = s.ToString());
                Progress<string> currentFolder = new(s => lblCurrent.Text = $"Current folder: {s}");
                Progress<DeletedEmailRecord> deletedProgress = new(AddDeletedRow);

                // Capture every UI-dependent setting here, on the UI thread, so the
                // background scan never touches WinForms controls across threads
                // (which can throw InvalidOperationException on large runs).
                ScanSettings settings = CaptureScanSettings();

                tokenSource = new CancellationTokenSource();
                CancellationToken token = tokenSource.Token;
                await Task.Run(() => ScanFolders(settings, folderProcessedCount, mailProcessedCount, duplicatesFound, nonMailItemsSkipped, currentFolder, deletedProgress, token));
                lastBackupFolder = settings.Delete ? settings.ExportRoot : null;
            }
            catch (Exception exception)
            {
                _ = MessageBox.Show(exception.Message);
            }

            progressTimer.Stop();
            scanStopwatch.Stop();
            ProgressTimer_Tick(this, EventArgs.Empty); // final totals

            // Lets show the result
            _ = new Process
            {
                StartInfo = new ProcessStartInfo(logToPath)
                {
                    UseShellExecute = true
                }
            }.Start();

            if (!string.IsNullOrEmpty(lastBackupFolder) && Directory.Exists(lastBackupFolder))
            {
                _ = MessageBox.Show($"Deleted emails were backed up to:\n{lastBackupFolder}\n\n" +
                    $"A list with a unique id for every deleted email is in {DeletedEmailManifest.CsvFileName} / {DeletedEmailManifest.JsonFileName}.\n" +
                    "Use \"Restore deleted emails…\" to put any of them back.");
            }

            btnStart.Text = "Start";
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            btnRestore.Enabled = true;
        }

        // Snapshot of everything the scan needs, read once on the UI thread.
        private ScanSettings CaptureScanSettings()
        {
            ScanSettings s = new();

            foreach (object? obj in listBoxOfFoldersToProcess.Items)
            {
                s.Folders.Add(((FolderTreeNode)obj).OutlookFolder);
            }

            s.MatchMessageId = checkBoxMatchOnMessageId?.Checked == true;
            s.MatchSentOn = checkBoxMatchOnSentOn.Checked;
            s.MatchReceivedTime = checkBoxMatchOnReceivedTime.Checked;
            s.MatchLastModTime = checkBoxMatchOnLastModTime.Checked;
            s.MatchSenderEmail = checkBoxMatchOnSenderEmail.Checked;
            s.MatchTo = checkBoxMatchOnTo.Checked;
            s.MatchCC = checkBoxMatchOnCC.Checked;
            s.MatchBCC = checkBoxMatchOnBCC.Checked;
            s.MatchSubject = checkBoxMatchOnSubject.Checked;
            s.MatchBody = checkBoxMatchOnBody.Checked;
            s.MatchHTMLBody = checkBoxMatchOnHTMLBody.Checked;
            s.MatchAttachment = checkBoxMatchOnAttachment.Checked;

            s.Copy = radioButtonCopyToFolder.Checked;
            s.Move = radioButtonMoveToFolder.Checked;
            s.Open = radioButtonOpenEachInOutlook.Checked;
            s.Delete = radioButtonDeleteEmails.Checked;

            // Fixed: previously this test referenced "Copy" twice, so choosing
            // "Move to folder" left the destination null and Move(null) threw.
            if (s.Copy || s.Move)
            {
                FolderTreeNode? item = (FolderTreeNode?)cboDestinationFolder.SelectedItem;
                if (item != null)
                {
                    s.Destination = item.OutlookFolder;
                }
            }

            s.SaveFilePath = filePathToSaveEmails;

            // For deletions, always keep a restorable backup. Use the folder the
            // user picked for saving duplicates if any, otherwise a per-run folder
            // under the user's Documents.
            if (s.Delete)
            {
                string baseDir = string.IsNullOrEmpty(filePathToSaveEmails)
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DuplicateEmailRemover")
                    : filePathToSaveEmails;
                s.ExportRoot = Path.Combine(baseDir, "Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            }

            return s;
        }

        public void ScanFolders(ScanSettings settings,
            IProgress<int> foldersProgress,
            IProgress<int> fileProgress,
            IProgress<int> duplicateCount,
            IProgress<int> nonMailItemCount,
            IProgress<string> currentFolderProgress,
            IProgress<DeletedEmailRecord> deletedProgress,
            CancellationToken token)
        {
            int folderCount = settings.Folders.Count;

            // As a safeguard, make sure we do not have any duplicate folders selected.
            for (int i = 0; i < folderCount; i++)
            {
                for (int j = i + 1; j < folderCount; j++)
                {
                    if (settings.Folders[i].FolderPath == settings.Folders[j].FolderPath)
                    {
                        _ = MessageBox.Show("Duplicate folders found in the processing list. Please remove duplicate folders.");
                        return;
                    }
                }
            }

            Outlook.NameSpace ns = outlook.GetNamespace("MAPI");

            // Duplicate detection is delegated to the (unit-tested) Core library.
            MatchOptions matchOptions = new()
            {
                MessageId = settings.MatchMessageId,
                SentOn = settings.MatchSentOn,
                ReceivedTime = settings.MatchReceivedTime,
                LastModTime = settings.MatchLastModTime,
                SenderEmail = settings.MatchSenderEmail,
                To = settings.MatchTo,
                CC = settings.MatchCC,
                BCC = settings.MatchBCC,
                Subject = settings.MatchSubject,
                Body = settings.MatchBody,
                HtmlBody = settings.MatchHTMLBody,
                Attachment = settings.MatchAttachment
            };
            DuplicateClassifier classifier = new(matchOptions);

            // Delete/Move/Copy are collected here and applied AFTER the scan.
            // Acting on items while enumerating a folder's Items collection silently
            // skips items in Outlook COM, so duplicates would be missed otherwise.
            List<DeferredAction> deferred = [];

            bool warningNotificationEnabled = true;
            bool canceled = false;
            int totalItemCount = 0;
            int totalDuplicateCount = 0;
            int totalNonMailItemCount = 0;

            // When deleting, prepare a restorable backup folder + manifest.
            DeletedEmailManifest? manifest = null;
            StreamWriter? csvWriter = null;
            if (settings.Delete && settings.ExportRoot != null)
            {
                manifest = new DeletedEmailManifest
                {
                    CreatedUtc = DateTime.UtcNow.ToString("u"),
                    BackupFolder = settings.ExportRoot
                };
                _ = Directory.CreateDirectory(manifest.MsgFolder);
                csvWriter = new StreamWriter(manifest.CsvPath, false, Encoding.UTF8);
                csvWriter.WriteLine(DeletedEmailRecord.CsvHeader());
            }

            using StreamWriter sw = new(logToPath, false, Encoding.UTF8, 65536);
            StringBuilder sb = new(1024);

            try
            {
            for (int i = 0; i < folderCount && !canceled; i++)
            {
                Outlook.Folder folder = settings.Folders[i];
                string folderPath = folder.FolderPath;
                Outlook.Items items = folder.Items;
                foldersProgress.Report(i);
                currentFolderProgress.Report(folderPath);

                foreach (object? item in items)
                {
                    if (token.IsCancellationRequested)
                    {
                        sw.Write("  User canceled via Stop button.\r\n");
                        canceled = true;
                        break;
                    }

                    fileProgress.Report(++totalItemCount);

                    if (item is not Outlook.MailItem mailItem)
                    {
                        nonMailItemCount.Report(++totalNonMailItemCount);
                        if (item != null)
                        {
                            _ = Marshal.ReleaseComObject(item);
                        }
                        continue;
                    }

                    // Classify this email (duplicate detection lives in Core).
                    ClassificationResult classification = classifier.Classify(new OutlookMailItemAdapter(mailItem), folderPath);
                    string hash = classification.Hash;

                    if (classification.Disposition == MailDisposition.Duplicate)
                    {
                        string origFolder = classification.OriginalFolderPath;
                        duplicateCount.Report(++totalDuplicateCount);

                        // We have a duplicate, lets first log it.
                        _ = sb.Clear();
                        _ = sb.AppendLine($"===================================\nDuplicate Found:" +
                            $"\n  Sent On: {mailItem.SentOn} " +
                            $"\n  Sender:  {mailItem.SenderEmailAddress} " +
                            $"\n  Subject: {mailItem.Subject ?? ""}");
                        if (origFolder != folderPath)
                        {
                            _ = sb.AppendLine($"\n  Original was here:\n    {origFolder}");
                        }

                        _ = sb.AppendLine($"\n  Duplicate found here:\n    {folderPath}");
                        _ = sb.AppendLine($"\n  HASH: {hash}");  //for debug
                        string msg = sb.ToString();
                        sw.WriteLine(msg);

                        // Lets check with the user if we should take an action on this duplicate email.
                        if (warningNotificationEnabled)
                        {
                            DialogResult dialog = MessageBox.Show($"Do you want to continue to be notified?\r\n Yes = Continue to notify me on each message\n" +
                                $" No = Stop notifying and process ALL remaining emails\n\n{msg}", "Duplicate found", MessageBoxButtons.YesNoCancel);
                            if (dialog == DialogResult.No)
                            {
                                warningNotificationEnabled = false;
                            }
                            else if (dialog == DialogResult.Cancel)
                            {
                                sw.WriteLine("  User canceled via Cancel button.\r\n");
                                canceled = true;
                                _ = Marshal.ReleaseComObject(mailItem);
                                break;
                            }
                        }

                        // For non-delete actions, optionally save a copy to a mirrored
                        // file-system folder (legacy behaviour). Deletions use the
                        // restorable backup below instead.
                        if (settings.SaveFilePath != null && !settings.Delete)
                        {
                            try
                            {
                                string sentOnDate = mailItem.SentOn.ToString();
                                string filenameFriendlyDate = string.Join("_", sentOnDate.Split(Path.GetInvalidFileNameChars()));
                                string path2 = $"{settings.SaveFilePath}{folderPath}\\{filenameFriendlyDate} {hash}.msg";
                                _ = Directory.CreateDirectory(settings.SaveFilePath + folderPath);
                                mailItem.SaveAs(path2);
                                sw.WriteLine($" Email saved to {path2}");
                            }
                            catch (Exception error)
                            {
                                LogErrorMessage(sw, error, "Email saved to file");
                            }
                        }

                        DeletedEmailRecord? record = null;

                        // Before deleting, back up the email to a unique .msg file and
                        // record it in the manifest so it can be listed and restored.
                        if (settings.Delete && manifest != null)
                        {
                            string id = Guid.NewGuid().ToString("N");
                            string msgRelative = Path.Combine(DeletedEmailManifest.MsgSubFolder, id + ".msg");
                            string msgFull = Path.Combine(manifest.BackupFolder, msgRelative);

                            record = new DeletedEmailRecord
                            {
                                Id = id,
                                FolderPath = folderPath,
                                StoreId = folder.StoreID,
                                OriginalEntryId = mailItem.EntryID,
                                Subject = mailItem.Subject ?? "",
                                SenderName = mailItem.SenderName ?? "",
                                SenderEmail = mailItem.SenderEmailAddress ?? "",
                                SentOn = mailItem.SentOn.ToString(),
                                ReceivedTime = mailItem.ReceivedTime.ToString(),
                                Size = mailItem.Size,
                                Hash = hash,
                                MsgFile = msgRelative
                            };

                            try
                            {
                                mailItem.SaveAs(msgFull, Outlook.OlSaveAsType.olMSGUnicode);
                            }
                            catch (Exception error)
                            {
                                LogErrorMessage(sw, error, "Email backed up to file");
                                record = null; // cannot guarantee a restorable backup; do not delete
                            }
                        }

                        if (settings.Open)
                        {
                            // Opening a window does not modify the Items collection, so it is safe inline.
                            try
                            {
                                mailItem.Display(true);
                            }
                            catch (Exception error)
                            {
                                LogErrorMessage(sw, error, "Email opened in Outlook");
                            }
                        }
                        else if (settings.Delete)
                        {
                            if (record != null)
                            {
                                // Defer the delete; finalize the manifest entry once removed.
                                deferred.Add(new DeferredAction(mailItem.EntryID, folder.StoreID, folderPath, record));
                                sw.Write("  Email queued for deletion\r\n");
                            }
                            else
                            {
                                sw.Write("  Skipped deletion (backup failed, email kept)\r\n");
                            }
                        }
                        else if (settings.Move || settings.Copy)
                        {
                            // Defer any action that mutates the folder's Items collection.
                            deferred.Add(new DeferredAction(mailItem.EntryID, folder.StoreID, folderPath, null));
                            sw.Write(settings.Move ? "  Email queued to move\r\n" : "  Email queued to copy\r\n");
                        }
                    }

                    _ = Marshal.ReleaseComObject(mailItem);
                }

                _ = Marshal.ReleaseComObject(items);
            }

            if (canceled)
            {
                sw.WriteLine($"\r\nRun canceled. {deferred.Count} queued action(s) were NOT applied; nothing was deleted, moved, or copied.");
                return;
            }

            ApplyDeferredActions(ns, settings, deferred, sw, manifest, csvWriter, deletedProgress);
            }
            finally
            {
                if (csvWriter != null)
                {
                    csvWriter.Flush();
                    csvWriter.Dispose();
                }

                if (manifest != null)
                {
                    try
                    {
                        manifest.SaveJson();
                        sw.WriteLine($"\r\nBacked up {manifest.Records.Count} deleted email(s) to: {manifest.BackupFolder}");
                        sw.WriteLine($"Manifest (restore list with unique ids): {manifest.JsonPath}");
                    }
                    catch (Exception error)
                    {
                        LogErrorMessage(sw, error, "writing the manifest");
                    }
                }
            }
        }

        // Phase 2: apply the delete/move/copy actions once scanning is finished,
        // re-fetching each item by its stable EntryID. Deletions finalize their
        // manifest entry (unique id + CSV row) only after the item is truly removed.
        private void ApplyDeferredActions(Outlook.NameSpace ns, ScanSettings settings, List<DeferredAction> deferred,
            StreamWriter sw, DeletedEmailManifest? manifest, StreamWriter? csvWriter, IProgress<DeletedEmailRecord> deletedProgress)
        {
            if (!ALLOW_MOVES_DELETES)
            {
                sw.WriteLine($"  Simulated (DEBUG MODE): {deferred.Count} action(s) would be applied.");
                return;
            }

            int appliedSinceSave = 0;
            foreach (DeferredAction act in deferred)
            {
                Outlook.MailItem? mailItem = null;
                try
                {
                    mailItem = ns.GetItemFromID(act.EntryId, act.StoreId) as Outlook.MailItem;
                    if (mailItem == null)
                    {
                        continue;
                    }

                    if (settings.Copy)
                    {
                        Outlook.MailItem copiedItem = mailItem.Copy();
                        copiedItem.Move(settings.Destination);
                        _ = Marshal.ReleaseComObject(copiedItem);
                    }
                    else if (settings.Move)
                    {
                        mailItem.Move(settings.Destination);
                    }
                    else if (settings.Delete)
                    {
                        mailItem.Delete();
                        sw.Write($"  Email Deleted ({act.FolderPath})\r\n");

                        // Only now is the deletion real: record it in the manifest.
                        if (act.Record != null)
                        {
                            act.Record.DeletedUtc = DateTime.UtcNow.ToString("u");
                            manifest?.Records.Add(act.Record);
                            csvWriter?.WriteLine(act.Record.ToCsvLine());
                            deletedProgress.Report(act.Record);

                            // Persist the manifest periodically so a crash loses little.
                            if (++appliedSinceSave >= 500)
                            {
                                appliedSinceSave = 0;
                                csvWriter?.Flush();
                                try { manifest?.SaveJson(); } catch { /* non-fatal */ }
                            }
                        }
                    }
                }
                catch (Exception error)
                {
                    LogErrorMessage(sw, error, settings.Copy
                        ? "Email copied to mailbox folder"
                        : settings.Move ? "Email moved to mailbox folder" : "Email Deleted");
                }
                finally
                {
                    if (mailItem != null)
                    {
                        _ = Marshal.ReleaseComObject(mailItem);
                    }
                }
            }
        }

        private static void LogErrorMessage(StreamWriter sw, Exception error, string action)
        {
            string message = $"  Failed {action} " +
                $"Message: {error.Message}\r\n" +
                $"Source: {error.Source}\r\n" +
                $"InnerException: {error.InnerException}\r\n";
            sw.Write(message);
            _ = MessageBox.Show(message);
        }

        private void checkBoxSaveEachDupToFilePath_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxSaveEachDupToFilePath.Checked)
            {
                using FolderBrowserDialog fbd = new();
                DialogResult result = fbd.ShowDialog(this);
                if (result == DialogResult.OK && Directory.Exists(fbd.SelectedPath))
                {
                    filePathToSaveEmails = fbd.SelectedPath;
                    txtSaveEachDupToFilePath.Text = fbd.SelectedPath;
                }
                else
                {
                    filePathToSaveEmails = null;
                    txtSaveEachDupToFilePath.Text = "";
                }
            }
            else
            {
                filePathToSaveEmails = null;
                txtSaveEachDupToFilePath.Text = "";
            }

            checkBoxSaveEachDupToFilePath.Text = "Save each duplicate to " + (filePathToSaveEmails ?? "a file.");
            txtSaveEachDupToFilePath.Visible = checkBoxSaveEachDupToFilePath.Checked;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            tokenSource?.Cancel();
        }

        private void btnI_Agree_Click(object sender, EventArgs e)
        {
            authorizedTabIndex = 4;
            tabControl1.SelectedIndex = 4;
        }
    }


    // A UI-thread snapshot of all settings the background scan needs, so the
    // scan never reads WinForms controls from a non-UI thread.
    internal sealed class ScanSettings
    {
        public List<Outlook.Folder> Folders { get; } = [];

        public bool MatchMessageId { get; set; }
        public bool MatchSentOn { get; set; }
        public bool MatchReceivedTime { get; set; }
        public bool MatchLastModTime { get; set; }
        public bool MatchSenderEmail { get; set; }
        public bool MatchTo { get; set; }
        public bool MatchCC { get; set; }
        public bool MatchBCC { get; set; }
        public bool MatchSubject { get; set; }
        public bool MatchBody { get; set; }
        public bool MatchHTMLBody { get; set; }
        public bool MatchAttachment { get; set; }

        public bool Copy { get; set; }
        public bool Move { get; set; }
        public bool Open { get; set; }
        public bool Delete { get; set; }

        public Outlook.MAPIFolder? Destination { get; set; }
        public string? SaveFilePath { get; set; }

        // When Delete is selected, every deleted email is backed up (as a .msg)
        // into this folder, together with a manifest listing them by unique id.
        public string? ExportRoot { get; set; }
    }

    // A duplicate action to apply after the scan, identified by its stable EntryID.
    // For deletions, Record carries the manifest entry to finalize once the item
    // is actually removed.
    internal sealed record DeferredAction(string EntryId, string StoreId, string FolderPath, DeletedEmailRecord? Record);

    internal sealed class FolderTreeNode : TreeNode
    {
        public override string ToString()
        {
            return $"{Text.TrimStart('\\')} ({ItemCount})";
        }

        public Outlook.Folder OutlookFolder { get; }

        public Outlook.Items OutlookItems { get; }

        public int ItemCount { get; }

        public string Path
        {
            get => Text;
            set => Text = value;
        }

        public FolderTreeNode(Outlook.Folder outlookFolder)
        {
            OutlookFolder = outlookFolder;
            Text = outlookFolder.FolderPath;
            OutlookItems = OutlookFolder.Items;
            ItemCount = OutlookFolder.Items.Count;
        }
    }
}