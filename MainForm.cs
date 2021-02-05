using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtractSWF
{
    public partial class MainForm : Form
    {
        List<string> pptFilenames;
        string folderName=null;
        string tmpFolder;

        public MainForm()
        {
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

            int errors = 0;
            bool aborted = false;

            foreach (string currentPPTFilename in pptFilenames)
            {
                string baseFileName = Path.GetFileNameWithoutExtension(currentPPTFilename);
                while (true)
                {
                    try
                    {
                        using (ZipArchive zip = ZipFile.Open(currentPPTFilename, ZipArchiveMode.Read))
                        {
                            foreach (ZipArchiveEntry entry in zip.Entries)
                            {
                                // check file is in correct directory with correct file extension
                                string fullName = entry.FullName;
                                if (!(fullName.StartsWith("ppt/activeX") && fullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)))
                                {
                                    continue;
                                }

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
                                }


                                tmpFile.Close();
                                File.Delete(tmpFileName);

                            }
                        }
                    } catch (Exception exception)
                    {
                        string caption = "Error loading '" + Path.GetFileName(currentPPTFilename) + "':\n\n" + exception.Message;
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

                    Application.DoEvents();
                }
                if (aborted)
                {
                    break;
                }
            }

            string error = errors.ToString() + " error";
            if (errors != 1)
            {
                error += "s";
            }
            error += ".";

            if (aborted)
            {
                MessageBox.Show("Extraction failed with " + error, "Extraction Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

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

        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This application was developed by Edward Lancaster.\n\nCopyright © Edward Lancaster 2021, All Rights Reserved.", "About ExtractSWF", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
