// Released under the MIT License. Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

// DISCLAIMER
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THIS CODE LIES SOLELY WITH THE USER.

// Created by Ryan Scott White on Nov. 2023, Updated 12/20/2024 (note: no AI used for coding as of 12/20/2024, may change with the future drafts)


using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
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

            if (!checkBoxMatchOnBody.Checked && !checkBoxMatchOnHTMLBody.Checked)
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
                  (checkBoxMatchOnSentOn.Checked ? 1 : 0)
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
            btnStart.Text = "Running";
            try
            {
                Progress<int> folderProcessedCount = new(s => txtFoldersProcessed.Text = s.ToString());
                Progress<int> mailProcessedCount = new(s => txtEmailsProcessed.Text = s.ToString());
                Progress<int> duplicatesFound = new(s => txtDuplicatesFound.Text = s.ToString());
                Progress<int> nonMailItemsSkipped = new(s => txtNonMailItemsSkipped.Text = s.ToString());

                // Capture every UI-dependent setting here, on the UI thread, so the
                // background scan never touches WinForms controls across threads
                // (which can throw InvalidOperationException on large runs).
                ScanSettings settings = CaptureScanSettings();

                tokenSource = new CancellationTokenSource();
                CancellationToken token = tokenSource.Token;
                await Task.Run(() => ScanFolders(settings, folderProcessedCount, mailProcessedCount, duplicatesFound, nonMailItemsSkipped, token));
            }
            catch (Exception exception)
            {
                _ = MessageBox.Show(exception.Message);
            }

            // Lets show the result
            _ = new Process
            {
                StartInfo = new ProcessStartInfo(logToPath)
                {
                    UseShellExecute = true
                }
            }.Start();

            btnStart.Text = "Start";
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        // Snapshot of everything the scan needs, read once on the UI thread.
        private ScanSettings CaptureScanSettings()
        {
            ScanSettings s = new();

            foreach (object? obj in listBoxOfFoldersToProcess.Items)
            {
                s.Folders.Add(((FolderTreeNode)obj).OutlookFolder);
            }

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
            return s;
        }

        public void ScanFolders(ScanSettings settings,
            IProgress<int> foldersProgress,
            IProgress<int> fileProgress,
            IProgress<int> duplicateCount,
            IProgress<int> nonMailItemCount, CancellationToken token)
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

            // hash -> folder path of the first (kept) occurrence.
            Dictionary<string, string> hs = [];

            // Delete/Move/Copy are collected here and applied AFTER the scan.
            // Acting on items while enumerating a folder's Items collection silently
            // skips items in Outlook COM, so duplicates would be missed otherwise.
            List<DeferredAction> deferred = [];

            bool warningNotificationEnabled = true;
            bool canceled = false;
            int totalItemCount = 0;
            int totalDuplicateCount = 0;
            int totalNonMailItemCount = 0;

            using StreamWriter sw = new(logToPath, false, Encoding.UTF8, 65536);
            StringBuilder sb = new(1024);

            for (int i = 0; i < folderCount && !canceled; i++)
            {
                Outlook.Folder folder = settings.Folders[i];
                string folderPath = folder.FolderPath;
                Outlook.Items items = folder.Items;
                foldersProgress.Report(i);

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

                    // Build the attributes to check for duplicates. (Subject/Received date/Sender/etc.)
                    _ = sb.Clear();
                    if (settings.MatchSentOn)
                    {
                        _ = sb.Append(mailItem.SentOn.ToString());
                    }

                    if (settings.MatchReceivedTime)
                    {
                        _ = sb.Append(mailItem.ReceivedTime.ToString());
                    }

                    if (settings.MatchLastModTime)
                    {
                        _ = sb.Append(mailItem.LastModificationTime.ToString());
                    }

                    if (settings.MatchSenderEmail)
                    {
                        _ = sb.Append(mailItem.SenderEmailAddress);
                    }

                    if (settings.MatchTo)
                    {
                        _ = sb.Append(mailItem.To);
                    }

                    if (settings.MatchCC)
                    {
                        _ = sb.Append(mailItem.CC);
                    }

                    if (settings.MatchBCC)
                    {
                        _ = sb.Append(mailItem.BCC);
                    }

                    if (settings.MatchSubject)
                    {
                        _ = sb.Append(mailItem.Subject);
                    }

                    if (settings.MatchBody)
                    {
                        _ = sb.Append(mailItem.Body);
                    }

                    if (settings.MatchHTMLBody)
                    {
                        _ = sb.Append(mailItem.HTMLBody);
                    }

                    if (settings.MatchAttachment)
                    {
                        Outlook.Attachments attachments = mailItem.Attachments;
                        foreach (Outlook.Attachment attachment in attachments)
                        {
                            _ = sb.Append(attachment.FileName);
                            _ = Marshal.ReleaseComObject(attachment);
                        }
                        _ = Marshal.ReleaseComObject(attachments);
                    }

                    byte[] buffer = Encoding.UTF8.GetBytes(sb.ToString());
                    byte[] digest = MD5.HashData(buffer);

                    _ = sb.Clear();
                    for (int j = 0; j < digest.Length; j++)
                    {
                        _ = sb.Append(digest[j].ToString("X2"));
                    }
                    string hash = sb.ToString();

                    if (hs.TryGetValue(hash, out string? origFolder))
                    {
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

                        // If the user requested to save the email to a folder, do that first before moving or deleting.
                        // (SaveAs does not modify the Items collection, so it is safe inline.)
                        if (settings.SaveFilePath != null)
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
                        else if (settings.Delete || settings.Move || settings.Copy)
                        {
                            // Defer any action that mutates the folder's Items collection.
                            deferred.Add(new DeferredAction(mailItem.EntryID, folder.StoreID, folderPath));
                            sw.Write(settings.Delete
                                ? "  Email queued for deletion\r\n"
                                : settings.Move ? "  Email queued to move\r\n" : "  Email queued to copy\r\n");
                        }
                    }
                    else
                    {
                        hs[hash] = folderPath;
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

            ApplyDeferredActions(ns, settings, deferred, sw);
        }

        // Phase 2: apply the delete/move/copy actions once scanning is finished,
        // re-fetching each item by its stable EntryID.
        private void ApplyDeferredActions(Outlook.NameSpace ns, ScanSettings settings, List<DeferredAction> deferred, StreamWriter sw)
        {
            if (!ALLOW_MOVES_DELETES)
            {
                sw.WriteLine($"  Simulated (DEBUG MODE): {deferred.Count} action(s) would be applied.");
                return;
            }

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
    }

    // A duplicate action to apply after the scan, identified by its stable EntryID.
    internal sealed record DeferredAction(string EntryId, string StoreId, string FolderPath);

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