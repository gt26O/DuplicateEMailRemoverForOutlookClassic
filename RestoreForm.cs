// Released under the MIT License (see LICENSE.txt).
//
// Restore window: opens a DeletedEmails.json manifest produced by a delete run,
// lists the deleted emails, and restores selected ones back into Outlook from
// their backup .msg files. Restore does not depend on Outlook EntryIDs, so it
// works even for Exchange / Office 365 accounts where IDs change over time.

using Outlook = Microsoft.Office.Interop.Outlook;

namespace DuplicateEMailRemover
{
    public sealed class RestoreForm : Form
    {
        private readonly Outlook.Application outlook;
        private readonly Outlook.NameSpace ns;

        private readonly Button btnOpen = new();
        private readonly Label lblManifest = new();
        private readonly ListView listView = new();
        private readonly Button btnRestoreSelected = new();
        private readonly Button btnRestoreAll = new();
        private readonly Label lblStatus = new();

        private DeletedEmailManifest? manifest;

        public RestoreForm(Outlook.Application outlookApp)
        {
            outlook = outlookApp;
            ns = outlook.GetNamespace("MAPI");

            Text = "Restore Deleted Emails";
            Width = 900;
            Height = 560;
            MinimumSize = new Size(640, 400);
            StartPosition = FormStartPosition.CenterParent;

            btnOpen.Text = "Open manifest (DeletedEmails.json)…";
            btnOpen.Location = new Point(12, 12);
            btnOpen.Size = new Size(300, 32);
            btnOpen.Click += BtnOpen_Click;

            lblManifest.Location = new Point(320, 18);
            lblManifest.AutoSize = true;
            lblManifest.Text = "No manifest loaded.";

            listView.Location = new Point(12, 56);
            listView.Size = new Size(860, 400);
            listView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listView.View = View.Details;
            listView.CheckBoxes = true;
            listView.FullRowSelect = true;
            listView.GridLines = true;
            _ = listView.Columns.Add("Restored", 80);
            _ = listView.Columns.Add("Subject", 260);
            _ = listView.Columns.Add("Sender", 180);
            _ = listView.Columns.Add("Sent On", 140);
            _ = listView.Columns.Add("Original Folder", 220);
            _ = listView.Columns.Add("Id", 220);

            btnRestoreSelected.Text = "Restore checked";
            btnRestoreSelected.Location = new Point(12, 466);
            btnRestoreSelected.Size = new Size(160, 36);
            btnRestoreSelected.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnRestoreSelected.Enabled = false;
            btnRestoreSelected.Click += (s, e) => RestoreRecords(onlyChecked: true);

            btnRestoreAll.Text = "Restore all pending";
            btnRestoreAll.Location = new Point(182, 466);
            btnRestoreAll.Size = new Size(160, 36);
            btnRestoreAll.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnRestoreAll.Enabled = false;
            btnRestoreAll.Click += (s, e) => RestoreRecords(onlyChecked: false);

            lblStatus.Location = new Point(360, 474);
            lblStatus.AutoSize = true;
            lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            lblStatus.Text = "Each email is restored to its original folder (Inbox if that folder no longer exists).";

            Controls.Add(btnOpen);
            Controls.Add(lblManifest);
            Controls.Add(listView);
            Controls.Add(btnRestoreSelected);
            Controls.Add(btnRestoreAll);
            Controls.Add(lblStatus);
        }

        private void BtnOpen_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog ofd = new()
            {
                Title = "Open a DeletedEmails.json manifest",
                Filter = "Deleted email manifest (*.json)|*.json|All files (*.*)|*.*",
                FileName = DeletedEmailManifest.JsonFileName
            };

            if (ofd.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            try
            {
                manifest = DeletedEmailManifest.LoadJson(ofd.FileName);
                lblManifest.Text = $"{manifest.Records.Count} record(s) — {manifest.BackupFolder}";
                PopulateList();
                btnRestoreSelected.Enabled = true;
                btnRestoreAll.Enabled = true;
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(this, $"Could not open the manifest:\n{ex.Message}", "Restore");
            }
        }

        private void PopulateList()
        {
            listView.BeginUpdate();
            listView.Items.Clear();
            if (manifest != null)
            {
                foreach (DeletedEmailRecord r in manifest.Records)
                {
                    ListViewItem item = new(r.Restored ? "Yes" : "No");
                    _ = item.SubItems.Add(r.Subject);
                    _ = item.SubItems.Add(string.IsNullOrEmpty(r.SenderName) ? r.SenderEmail : r.SenderName);
                    _ = item.SubItems.Add(r.SentOn);
                    _ = item.SubItems.Add(r.FolderPath);
                    _ = item.SubItems.Add(r.Id);
                    item.Tag = r;
                    item.ForeColor = r.Restored ? SystemColors.GrayText : SystemColors.ControlText;
                    _ = listView.Items.Add(item);
                }
            }
            listView.EndUpdate();
        }

        private void RestoreRecords(bool onlyChecked)
        {
            if (manifest == null)
            {
                return;
            }

            List<ListViewItem> targets = [];
            foreach (ListViewItem item in listView.Items)
            {
                DeletedEmailRecord r = (DeletedEmailRecord)item.Tag!;
                if (r.Restored)
                {
                    continue;
                }

                if (!onlyChecked || item.Checked)
                {
                    targets.Add(item);
                }
            }

            if (targets.Count == 0)
            {
                _ = MessageBox.Show(this, "Nothing to restore (no pending items selected).", "Restore");
                return;
            }

            btnRestoreSelected.Enabled = false;
            btnRestoreAll.Enabled = false;

            int done = 0;
            int failed = 0;
            foreach (ListViewItem item in targets)
            {
                DeletedEmailRecord r = (DeletedEmailRecord)item.Tag!;
                lblStatus.Text = $"Restoring {done + failed + 1} of {targets.Count}…";
                Refresh();

                try
                {
                    RestoreOne(r);
                    item.Text = "Yes";
                    item.Checked = false;
                    item.ForeColor = SystemColors.GrayText;
                    done++;
                }
                catch (Exception ex)
                {
                    failed++;
                    _ = MessageBox.Show(this, $"Failed to restore \"{r.Subject}\":\n{ex.Message}", "Restore");
                }

                // Persist progress after every item so a crash cannot lose the record of what was restored.
                try
                {
                    manifest.SaveJson();
                }
                catch
                {
                    // Non-fatal: restoring continues even if the manifest cannot be rewritten.
                }
            }

            lblStatus.Text = $"Done. Restored {done}, failed {failed}.";
            btnRestoreSelected.Enabled = true;
            btnRestoreAll.Enabled = true;
        }

        private void RestoreOne(DeletedEmailRecord record)
        {
            string msgPath = manifest!.ResolveMsgPath(record);
            if (!File.Exists(msgPath))
            {
                throw new FileNotFoundException($"Backup file not found: {msgPath}");
            }

            object shared = ns.OpenSharedItem(msgPath);
            if (shared is not Outlook.MailItem mail)
            {
                throw new InvalidOperationException("The backup file is not a mail item.");
            }

            Outlook.MAPIFolder target = FindFolderByPath(record.FolderPath)
                ?? ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            _ = mail.Move(target);
            record.Restored = true;
            record.RestoredUtc = DateTime.UtcNow.ToString("u");
        }

        private Outlook.MAPIFolder? FindFolderByPath(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return null;
            }

            foreach (Outlook.Folder top in ns.Folders)
            {
                Outlook.MAPIFolder? found = SearchFolder(top, folderPath);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private static Outlook.MAPIFolder? SearchFolder(Outlook.MAPIFolder folder, string folderPath)
        {
            if (string.Equals(folder.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase))
            {
                return folder;
            }

            foreach (Outlook.MAPIFolder sub in folder.Folders)
            {
                Outlook.MAPIFolder? found = SearchFolder(sub, folderPath);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }
    }
}
