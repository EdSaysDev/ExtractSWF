using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Taskbar;
using ProgressBarControlUtils;
using System.Threading;

namespace ExtractSWF
{
    public partial class MainForm : Form
    {
        List<string> pptFilenames;
        string folderName=null;
        string tmpFolder;

        Thread exportThread;

        public MainForm()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string resourceName = new AssemblyName(args.Name).Name + ".dll";
                string resource = Array.Find(this.GetType().Assembly.GetManifestResourceNames(), element => element.EndsWith(resourceName));

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
                {
                    byte[] assemblyData = new byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return Assembly.Load(assemblyData);
                }
            };

            InitializeComponent();

            tmpFolder = Path.GetTempPath();
            pptFilenames = new List<string>();
        }

        /*
        private void chooseFileBtn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "PowerPoint Presentations (*.pptx)|*.pptx|All files (*.*)|*.*";
                //openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(openFileDialog.FileName))
                {
                    //Get the path of specified file
                    pptFilename = openFileDialog.FileName;
                    filenameLbl.Text = pptFilename;
                    baseFileName = Path.GetFileNameWithoutExtension(pptFilename);
                }
            }
        }
        */

        private void exportBtn_Click(object sender, EventArgs e)
        {
            if (exportThread != null) {
                if (exportThread.ThreadState == ThreadState.Running)
                {
                    {
                        MessageBox.Show("You must wait for the previous export to finish first", "Have some patience", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            if (pptFilenames.Count == 0)
            {
                MessageBox.Show("You must select at least one PowerPoint Presentation first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (folderName == null)
            {
                MessageBox.Show("You must select an output folder first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            exportThread = new Thread(() => { runExport(); });
            exportThread.Start();

        }
        void runExport() {

            int errors = 0;
            bool aborted = false;

            int fileCount = 0;
            int binFilesCount = 0;
            int binFilesTotal = 0;
            foreach (string currentPPTFilename in pptFilenames)
            {
                fileCount++;
                string baseFileName = Path.GetFileNameWithoutExtension(currentPPTFilename);
                while (true)
                {
                    int flashFileCount = 0;
                    try
                    {
                        updateProgressBars(TaskbarProgressBarState.Normal, fileCount, 0, pptFilenames.Count, 1);
                        using (ZipArchive zip = ZipFile.Open(currentPPTFilename, ZipArchiveMode.Read))
                        {
                            List<ZipArchiveEntry> binFiles = new List<ZipArchiveEntry>();


                            foreach (ZipArchiveEntry entry in zip.Entries)
                            {
                                // check file is in correct directory with correct file extension
                                string fullName = entry.FullName;
                                if (!(fullName.StartsWith("ppt/activeX") && fullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)))
                                {
                                    continue;
                                }
                                binFiles.Add(entry);
                            }

                            binFilesCount = 0;
                            binFilesTotal = binFiles.Count;
                            updateProgressBars(TaskbarProgressBarState.Normal, fileCount, 0, pptFilenames.Count, binFilesTotal);
                            foreach (ZipArchiveEntry entry in binFiles)
                            {
                                binFilesCount++;

                                // extract file
                                string outputFileName = folderName + "\\" + baseFileName + "_" + Path.GetFileNameWithoutExtension(entry.Name) + ".swf";
                                string tmpFileName = tmpFolder + "\\" + baseFileName + "_" + entry.Name;
                                if (File.Exists(tmpFileName))
                                {
                                    File.Delete(tmpFileName);
                                }

                                entry.ExtractToFile(tmpFileName);

                                // check file contains flash magic number
                                Stream tmpFile = File.Open(tmpFileName, FileMode.Open);

                                int dataLen = 3;
                                long lastRead = tmpFile.Length - dataLen;
                                int counter = 0;
                                byte[] data = new byte[dataLen];
                                bool found = false;
                                while (counter < lastRead)
                                {
                                    tmpFile.Read(data, 0, 3);
                                    tmpFile.Seek(counter, SeekOrigin.Begin);
                                    if (Encoding.ASCII.GetString(data) == "FWS")
                                    {
                                        found = true;
                                        break;
                                    }
                                    counter++;
                                }

                                if (found)
                                {

                                    // it's a flash file, so rename, trim, and move it
                                    Stream outputFile = File.Open(outputFileName, FileMode.OpenOrCreate);
                                    tmpFile.Position -= 1;

                                    // create a buffer to hold the bytes 
                                    byte[] buffer = new byte[1024];
                                    int bytesRead;

                                    // while the read method returns bytes
                                    // keep writing them to the output stream
                                    while ((bytesRead = tmpFile.Read(buffer, 0, 1024)) > 0)
                                    {
                                        outputFile.Write(buffer, 0, bytesRead);
                                    }
                                    outputFile.Close();

                                    flashFileCount++;
                                }


                                tmpFile.Close();
                                File.Delete(tmpFileName);

                            }
                            updateProgressBars(TaskbarProgressBarState.Normal, fileCount, binFilesCount, pptFilenames.Count, binFilesTotal);
                        }
                    } catch (Exception exception)
                    {
                        string caption = "Error loading '" + Path.GetFileName(currentPPTFilename) + "':\n\n" + exception.Message;
                        updateProgressBars(TaskbarProgressBarState.Error);
                        DialogResult result = MessageBox.Show(caption, "Extraction Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3);
                        if (result == DialogResult.Abort)
                        {
                            errors++;
                            aborted = true;
                            break;
                        }
                        if (result == DialogResult.Retry)
                        {
                            continue;
                        }
                        if (result == DialogResult.Ignore)
                        {
                            errors++;
                            break;
                        }
                    }

                    if (flashFileCount == 0)
                    {
                        updateProgressBars(TaskbarProgressBarState.Paused);

                        // allow application to redraw
                        Application.DoEvents();

                        MessageBox.Show("No flash files found in '" + Path.GetFileName(currentPPTFilename) + "'.", "No Flash Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    break;
                }

                
                if (aborted)
                {
                    break;
                }

                updateProgressBars(TaskbarProgressBarState.Normal, fileCount, binFilesCount, pptFilenames.Count, binFilesTotal);
            }

            string error = errors.ToString() + " error";
            if (errors != 1)
            {
                error += "s";
            }
            error += ".";

            if (aborted)
            {
                // allow application to redraw
                Application.DoEvents();

                MessageBox.Show("Extraction failed with " + error, "Extraction Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updateProgressBars(TaskbarProgressBarState.Error);
                return;
            }

            if (binFilesTotal == 0)
            {
                binFilesTotal = 1;
            }
            updateProgressBars(TaskbarProgressBarState.Normal, pptFilenames.Count, binFilesTotal, pptFilenames.Count, binFilesTotal);
            updateProgressBars(TaskbarProgressBarState.NoProgress, false);

            // allow application to redraw
            Application.DoEvents();

            if (errors > 0)
            {
                MessageBox.Show("Extraction failed successfully with " + error, "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } else {
                MessageBox.Show("Extraction completed.", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void dragPanel_DragEnter(object sender, DragEventArgs e)
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

        bool CaseInsensitiveContains(List<string> lst, string s)
        {
            return lst.FindIndex(x => x.Equals(s, StringComparison.OrdinalIgnoreCase)) != -1;
        }

        private void dragPanel_DragDrop(object sender, DragEventArgs e)
        {
            bool valid = true;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                bool dirListedAlready = false;
                string[] data = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach(string fn in data)
                {
                    if (File.Exists(fn))
                    {
                        if (!CaseInsensitiveContains(pptFilenames, fn))
                        {
                            pptFilenames.Add(fn);
                        } else
                        {
                            valid = false;
                        }
                    }

                    else if (Directory.Exists(fn))
                    {
                        if (dirListedAlready)
                        {
                            valid = false;
                        } else
                        {
                            folderName = fn;
                            dirListedAlready = true;
                        }
                    }

                    else
                    {
                        // doesn't exist
                        valid = false;
                    }
                }
            }

            folderLbl.Text = folderName;

            fileList.Items.Clear();
            foreach (string pptName in pptFilenames)
            {
                fileList.Items.Add(pptName);
            }

            if (!valid)
            {
                Application.DoEvents(); // so control contents are updated
                MessageBox.Show("Sorry, not all of that data is not allowed to be dragged and dropped.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void clearListBtn_Click(object sender, EventArgs e)
        {
            fileList.Items.Clear();
            pptFilenames.Clear();
        }

        private void openFolderBtn_Click(object sender, EventArgs e)
        {
            if (folderName == null)
            {
                MessageBox.Show("You must choose an output folder first!", "Oh dear...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            System.Diagnostics.Process.Start("explorer.exe", folderName);
        }

        void updateProgressBars(TaskbarProgressBarState state, bool updateFormPBars = true)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new MethodInvoker (() => { updateProgressBars(state, updateFormPBars); }));
            }
            else
            {
                TaskbarManager windowsTaskbar = TaskbarManager.Instance;
                windowsTaskbar.SetProgressState(state);

                if (updateFormPBars)
                {
                    // update progress bars and task bar states
                    ProgressBarState pBarState = ProgressBarState.Normal;
                    if (state == TaskbarProgressBarState.Error)
                    {
                        pBarState = ProgressBarState.Error;
                    }
                    else if (state == TaskbarProgressBarState.Paused)
                    {
                        pBarState = ProgressBarState.Paused;
                    }
                    else if (state == TaskbarProgressBarState.NoProgress)
                    {
                        flashFileProgressBar.Value = 0;
                        fileProgressBar.Value = 0;
                    }

                    ModifyProgressBarColour.SetState(flashFileProgressBar, pBarState);
                    ModifyProgressBarColour.SetState(fileProgressBar, pBarState);
                }

                Invalidate();
            }
        }

        void updateProgressBars(TaskbarProgressBarState state, int fileNum, int flashNum, int fileMax, int flashMax)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new MethodInvoker(() => { updateProgressBars(state, fileNum, flashNum, fileMax, flashMax); }));
            }
            else
            {
                TaskbarManager windowsTaskbar = TaskbarManager.Instance;

                // check the values are in an acceptable range
                fileMax = Math.Max(fileMax, 1);
                fileNum = Math.Min(Math.Max(fileNum, 0), fileMax);

                flashMax = Math.Max(flashMax, 1);
                flashNum = Math.Min(Math.Max(flashNum, 0), flashMax);

                // update text boxes
                fileNumLbl.Text = "File " + fileNum + "/" + fileMax;
                flashNumLbl.Text = "Flash file " + flashNum + "/" + flashMax;

                updateProgressBars(state);

                // update values
                windowsTaskbar.SetProgressValue(fileNum, fileMax);

                flashFileProgressBar.Maximum = flashMax;
                flashFileProgressBar.Value = flashNum;

                fileProgressBar.Maximum = fileMax;
                fileProgressBar.Value = fileNum;

                // force repaint
                Invalidate();

                // respond to events so changes are drawn
                //Application.DoEvents();
            }
        }

        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This application was developed by Edward Lancaster.\n\nCopyright © Edward Lancaster 2021, All Rights Reserved.", "About ExtractSWF", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = false;
        }

        /*private void chooseFolderBtn_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    folderName = fbd.SelectedPath;
                    folderLbl.Text = folderName;
                }
            }
        }
        */
    }
}
