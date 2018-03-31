using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;



namespace CASPER
{
    public partial class UserForm : Form
    {
        //For moving form.////////////////////////////////////////////////////////
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        /////////////////////////////////////////////////////////////////////////
        public UserForm()
        {
            MaximizeBox = false;
            MinimizeBox = false;
            InitializeComponent();
            circleBar.BackColor = System.Drawing.Color.Transparent;
        }
        private void btn_About_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("C:\\CASPER\\ReadMe.txt"); //Opens ReadMe text file.
        }
        private void btn_Input_Click(object sender, EventArgs e)
        {
            //For input directory 
            var dialogBox = new Ionic.Utils.FolderBrowserDialogEx();
            dialogBox.Description = "Select an output destination.";
            dialogBox.ShowNewFolderButton = true;
            dialogBox.ShowEditBox = true;
            dialogBox.SelectedPath = txtBox_Input.Text;
            dialogBox.ShowFullPathInEditBox = true;
            dialogBox.RootFolder = System.Environment.SpecialFolder.MyComputer;

            System.Windows.Forms.DialogResult result = dialogBox.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txtBox_Input.Text = dialogBox.SelectedPath;
            }
            if (txtBox_Output.Text == "")
            {
                lbl_Progress.Text = "Waiting on output path...";
                lbl_Progress.Update();
            }
            else
            {
                lbl_Progress.Text = "Waiting to start extraction...";
                lbl_Progress.Update();
            }
            dialogBox = null;
        }
        private void btn_Output_Click(object sender, EventArgs e)
        {
            //For output Directory
            var dialogBox = new Ionic.Utils.FolderBrowserDialogEx();
            if (System.IO.Directory.Exists(txtBox_Input.Text))
            {
                dialogBox.SelectedPath = txtBox_Input.Text;
            }
            dialogBox.Description = "Select an output destination.";
            dialogBox.ShowNewFolderButton = true;
            dialogBox.ShowEditBox = true;
            //dialogBox.SelectedPath = txtBox_Output.Text;
            dialogBox.ShowFullPathInEditBox = true;
            dialogBox.RootFolder = System.Environment.SpecialFolder.MyComputer;

            System.Windows.Forms.DialogResult result = dialogBox.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txtBox_Output.Text = dialogBox.SelectedPath;
            }
            if (txtBox_Input.Text == "")
            {
                lbl_Progress.Text = "Waiting on input path...";
                lbl_Progress.Update();
            }
            else
            {
                lbl_Progress.Text = "Waiting to start extraction...";
                lbl_Progress.Update();
            }
            dialogBox = null;
        }
        private void btn_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btn_Minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void btn_Extract_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            lbl_percent.Text = "";
            lbl_percent.Update();
            //Main function
            circleBar.Value = 0;
            circleBar.Update();
            lbl_percent.Text = "0%";
            lbl_percent.Update();
            if (txtBox_Input.Text == "")
            {//Error message if user does not select an input directory
                System.Windows.Forms.MessageBox.Show("Please enter something in the input textbox.");
                return;
            }
            if (txtBox_Output.Text == "")
            {//Error message if user does not select an output directory. 
                System.Windows.Forms.MessageBox.Show("Please enter something in the output textbox.");
                return;
            }
            string pathToConfigFile = "C:\\CASPER\\ConfigFile.txt";
            if (!System.IO.File.Exists(pathToConfigFile))
            {//Error message if there is no configuration text file found. 
                System.Windows.Forms.MessageBox.Show("The directory to the configuration file is incorrect");
                return;
            }
            lbl_Progress.Text = "Beginning extraction.";
            lbl_Progress.Update();
            int numOfZips = 0;
            string parentDirectory = txtBox_Input.Text;

            ///////////////////////////////////////////////// For Config File
            List<string> folderList = new List<string>();
            List<string> fileList = new List<string>();
            List<string> zipList = new List<string>();
            List<string> internalZipList = new List<string>();
            bool keepUnprocessedData = true;
            /////////////////////////////////////////////////

            readConfigFile(pathToConfigFile, ref zipList, ref folderList, ref fileList, ref internalZipList, ref keepUnprocessedData);

            if (zipList[0] == "")
            {//Error message if the user does not include any zip file names in the configuration text file. 
                lbl_Progress.Text = "Error: No zip files specified in configuration file.";
                lbl_Progress.Update();
                System.Windows.Forms.MessageBox.Show("There is no zip files to look for in the configuration file.");
                return;
            }

            ///////////////////////////////////////////////// For extraction
            List<string> pathsOfUnreadableZips = new List<string>();
            List<string> pathsOfEmptyZipsToDelete = new List<string>();
            List<string> pathsWithOriginalZip = new List<string>();
            List<string> pathsWithoutOriginalZip = new List<string>();
            List<string> nameOfZipsNoPath = new List<string>();
            List<string> newFolderPaths = new List<string>();
            List<string> pathOfNewZips = new List<string>();
            List<string> pathOfOriginalXML = new List<string>();
            List<string> XMLLastDirectory = new List<string>();
            List<Ionic.Zip.ZipFile> newZipsCreated = new List<Ionic.Zip.ZipFile>();
            string finalPath = ""; 
            /////////////////////////////////////////////////

            getPathsOfOrginalZip(parentDirectory, ref pathsWithOriginalZip, ref pathsWithoutOriginalZip, ref numOfZips, ref pathOfOriginalXML, ref XMLLastDirectory, zipList);
            getNamesOfOriginalZip(ref nameOfZipsNoPath, pathsWithOriginalZip, pathsWithoutOriginalZip);
            createNewPaths(pathsWithoutOriginalZip, ref newFolderPaths, ref finalPath);
            createNewZips(numOfZips, newFolderPaths, nameOfZipsNoPath, ref newZipsCreated, ref pathOfNewZips, pathOfOriginalXML, XMLLastDirectory);

            if (numOfZips == 0)
            {//Error message if no zip files with a name specified in the configuration text file are found in any directory. 
                lbl_Progress.Text = "Error: No zip files found with the same name.";
                lbl_Progress.Update();
                System.Windows.Forms.MessageBox.Show("No zip files found.");
                return;
            }

            if (pathsWithOriginalZip.Count != numOfZips || newFolderPaths.Count != numOfZips || newZipsCreated.Count != numOfZips)
            {//Error message if something else went wrong such as an invalid output path. 
                System.Windows.Forms.MessageBox.Show("An error occurred. Make sure you have a valid destination path or valid zip files.");
                return;
            }
            circleBar.Maximum = numOfZips;
            double updateVal = 0;
            double increment = 100.00 / numOfZips;
            lbl_Progress.Text = "Extracting data.";
            lbl_Progress.Update();

            List<string> summaryText = new List<string>();
            summaryText.Add("CASPER Version: 1.1.0");
            summaryText.Add("FileName" + "Status".PadLeft(97));

            txtbox_Summary.Text += "FileName" + "Status".PadLeft(97) + "\r\n";
            txtbox_Summary.Update();

            try
            {
                for (int i = 0; i < numOfZips; i++)
                {
                    extractData(pathsWithOriginalZip[i], newFolderPaths[i], newZipsCreated[i], folderList, fileList, zipList, internalZipList, pathOfNewZips[i], keepUnprocessedData, ref summaryText);    //Extract Data into the new directories
                    updateVal += increment;
                    if (updateVal < 1)
                    {
                        lbl_percent.Text = "0%";
                        lbl_percent.Update();
                    }
                    else
                    {
                        lbl_percent.Text = updateVal.ToString("#") + "%";
                        lbl_percent.Update();
                    }
                    circleBar.Value += 1;
                    circleBar.Update();
                }
            }
            catch(System.OutOfMemoryException o)
            {
                MessageBox.Show(o.Message);
                Application.Exit();
            }
        
            finalPath += getLastDirectoryName(txtBox_Input.Text);
            if (System.IO.Directory.Exists(finalPath))
            {
                finalCleanUp(finalPath);
            }

            string summaryTextPath = finalPath + "\\CASPERLOG.txt";
            System.IO.File.WriteAllLines(summaryTextPath, summaryText);

            Cursor.Current = Cursors.Default;
            lbl_percent.Text = "100%";
            lbl_percent.Update();
            lbl_Progress.Text = "Processing Successful.";
            lbl_Progress.Update();
        }
        private static string makeRelative(string _filePath, string _referencePath)
        {
            var fileUri = new System.Uri(_filePath);
            var referenceUri = new System.Uri(_referencePath);
            return System.Uri.UnescapeDataString(referenceUri.MakeRelativeUri(fileUri).ToString());
        }
        private static string getMeshFileName(string _filePath)
        {
            List<string> parts = _filePath.Split('/').ToList();
            return parts[parts.Count - 1];
        }
        private static string getLastDirectoryName(string _filePath)
        {
            List<string> parts = _filePath.Split('\\').ToList();
            return parts[parts.Count - 1];
        }
        private static void getNamesOfOriginalZip(ref List<string> _zipNames, List<string> _pathWithZips, List<string> _pathWithoutZips)
        {//Gets just the names of the zip files without the full path. Stores the names in a list. 
            int count = _pathWithoutZips.Count;
            for (int i = 0; i < count; i++)
            {
                _zipNames.Add(makeRelative(_pathWithZips[i], _pathWithoutZips[i] + "\\"));
            }
        }
        private static string goUpOneDirectoryRelative(string _path)
        {
            System.Text.StringBuilder updatedPath = new System.Text.StringBuilder(_path);
            int numOfPaths = _path.Length - 1;
            for (int i = numOfPaths; i >= 0; i--)
            {
                if (_path[i] == '/')
                {
                    updatedPath.Remove(i, 1);
                    break;
                }
                else
                {
                    updatedPath.Remove(i, 1);
                }
            }
            return updatedPath.ToString();
        }
        private static string goUpOneDirectoryAbsolute(string _path)
        {
            System.Text.StringBuilder updatedPath = new System.Text.StringBuilder(_path);
            int numOfPaths = _path.Length - 1;
            for (int i = numOfPaths; i >= 0; i--)
            {
                if (_path[i] == '\\')
                {
                    updatedPath.Remove(i, 1);
                    break;
                }
                else
                {
                    updatedPath.Remove(i, 1);
                }
            }
            return updatedPath.ToString();
        }
        private static void readConfigFile(string _pathToConfigFile, ref List<string> _zipList, ref List<string> _folderList, ref List<string> _fileList, ref List<string> _internalZipList, ref bool _keepUnprocessedData)
        {//Reads the configuration text file and creates four lists: zip, folder, file, and internal zip. 
            string zips;
            string folders;
            string files;
            string internalZips;
            string keepUnprocessed;
            string skip;


            const System.Int32 BufferSize = 128;
            using (var fileStream = System.IO.File.OpenRead(_pathToConfigFile))
            using (var streamReader = new System.IO.StreamReader(fileStream, System.Text.Encoding.UTF8, true, BufferSize))
            {
                skip = streamReader.ReadLine();
                zips = streamReader.ReadLine().Replace(" ", string.Empty).ToLower();
                skip = streamReader.ReadLine();
                folders = streamReader.ReadLine().Replace(" ", string.Empty).ToLower();
                skip = streamReader.ReadLine();
                files = streamReader.ReadLine().Replace(" ", string.Empty).ToLower();
                skip = streamReader.ReadLine();
                internalZips = streamReader.ReadLine().Replace(" ", string.Empty).ToLower();
                skip = streamReader.ReadLine();
                keepUnprocessed = streamReader.ReadLine().Replace(" ", string.Empty).ToLower();
            }

            _zipList = zips.Split(',').ToList();
            _folderList = folders.Split(',').ToList();
            _fileList = files.Split(',').ToList();
      
            if(keepUnprocessed == "yes")
            {
                _keepUnprocessedData = true;
            }
            else
            {
                _keepUnprocessedData = false;
            }

            if (internalZips == "")
            {
                return;
            }
            else
            {
                _internalZipList = internalZips.Split(',').ToList();
            }
        }
        private static bool isMatchingZip(string _fileName, List<string> _zipList)
        {//Returns true if the entry matches anything in the zip list, false otherwise. 
            string currentZip;
            int numOfZips = _zipList.Count;

            for (int i = 0; i < numOfZips; i++)
            {
                currentZip = _zipList[i];
                if (currentZip.Contains("*"))
                {
                    if (currentZip.StartsWith("*"))
                    {
                        int startIndex = currentZip.IndexOf('*');
                        currentZip = currentZip.Remove(startIndex, 1);
                        if (_fileName.EndsWith(currentZip))
                        {
                            return true;
                        }
                    }
                    else if (currentZip.EndsWith("*"))
                    {
                        int startIndex = currentZip.IndexOf('*');
                        currentZip = "\\" + currentZip.Remove(startIndex, 1);
                        if (_fileName.Contains(currentZip))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        List<string> parts = currentZip.Split('*').ToList();
                        int numOfParts = parts.Count;
                        for (int j = 0; j < numOfParts; j++)
                        {
                            if (!_fileName.Contains(parts[j]))
                            {
                                return false;
                            }
                        }
                        return true;
                    }
                }
                else
                {
                    string czip = "\\" + currentZip;
                    if(_fileName.EndsWith(czip))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private static bool isThereMatch(string _fileName, List<string> _folderList, List<string> _fileList)
        {//Returns true if entry matches anything in the file or folder lists, false otherwise. 
            string currentFolder;
            string currentFile;
            int numOfFolders = _folderList.Count;
            int numOfFiles = _fileList.Count;
            _fileName = "/" + _fileName + "/";
            if (_folderList[0] == "")
            {
                goto Files;   //empty folder list
            }
            for (int i = 0; i < numOfFolders; i++)
            {
                currentFolder = _folderList[i];
                if (currentFolder.Contains("*"))
                {
                    if (currentFolder.StartsWith("*"))
                    {
                        int startIndex = currentFolder.IndexOf('*');
                        currentFolder = currentFolder.Remove(startIndex, 1) + "/";
                        if (_fileName.EndsWith(currentFolder))
                        {
                            return true;
                        }
                    }
                    else if (currentFolder.EndsWith("*"))
                    {
                        int startIndex = currentFolder.IndexOf('*');
                        currentFolder = "/" + currentFolder.Remove(startIndex, 1);
                        if (_fileName.Contains(currentFolder))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        List<string> parts = currentFolder.Split('*').ToList();
                        int lastIndex = parts.Count - 1;
                        parts[0] = "/" + parts[0];
                        parts[lastIndex] = parts[lastIndex] + "/";
                        int numOfParts = parts.Count;
                        for (int j = 0; j < numOfParts; j++)
                        {
                            if (!_fileName.Contains(parts[j]))
                            {
                                return false;
                            }
                        }
                        return true;
                    }
                }
                else
                {
                    string cfolder = "/" + currentFolder + "/";
                    if(_fileName.Contains(cfolder))
                    {
                        return true;
                    }
                }
            }

        Files:
            if (_fileList[0] == "")
            {
                return false;
            }

            for (int i = 0; i < numOfFiles; i++)
            {
                currentFile = _fileList[i];
                if (currentFile.Contains("*"))
                {
                    if (currentFile.StartsWith("*"))
                    {
                        int startIndex = currentFile.IndexOf('*');
                        currentFile = currentFile.Remove(startIndex, 1) + "/";
                        if (_fileName.EndsWith(currentFile))
                        {
                            return true;
                        }
                    }
                    else if (currentFile.EndsWith("*"))
                    {
                        int startIndex = currentFile.IndexOf('*');
                        currentFile = "/" + currentFile.Remove(startIndex, 1);
                        if (_fileName.Contains(currentFile))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        List<string> parts = currentFile.Split('*').ToList();
                        int numOfParts = parts.Count;
                        int lastIndex = parts.Count - 1;
                        parts[lastIndex] = parts[lastIndex] + "/";
                        parts[0] = "/" + parts[0];
                        for (int j = 0; j < numOfParts; j++)
                        {
                            if (!_fileName.Contains(parts[j]))
                            {
                                return false;
                            }
                        }
                        return true;
                    }
                   
                }
                else
                {
                    string cfile = "/" + currentFile + "/";
                    if(_fileName.Contains(cfile))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private void getPathsOfOrginalZip(string _path, ref List<string> _pathToZips, ref List<string> _pathWithoutZip, ref int _numOfZips, ref List<string> _pathOfOriginalXML, ref List<string> _xmlLastDirectory, List<string> _zipList)
        {//Gets full file path of original zip files. 
            lbl_Progress.Text = "Inspecting paths to zip files.";
            lbl_Progress.Update();

            string[] files = System.IO.Directory.GetFiles(_path);
            string[] subDirectories = System.IO.Directory.GetDirectories(_path);
            int numOfFiles = files.Length;
            int numOfSubDirectories = subDirectories.Length;

            for (int i = 0; i < numOfFiles; i++)
            {
                if (isMatchingZip(files[i].ToLower(), _zipList))
                {
                    _pathWithoutZip.Add(_path);
                    _pathToZips.Add(files[i]);
                    _numOfZips++;
                }
                if(files[i].EndsWith(".xml"))
                {
                    string directoryTemp = goUpOneDirectoryAbsolute(files[i]);
                    string temp2 = getLastDirectoryName(directoryTemp);
                    _pathOfOriginalXML.Add(files[i]);
                    _xmlLastDirectory.Add(temp2);
                }
            }
            if (numOfSubDirectories == 0)
            {
                return;
            }
            for (int i = 0; i < numOfSubDirectories; i++)
            {
                getPathsOfOrginalZip(subDirectories[i], ref _pathToZips, ref _pathWithoutZip, ref _numOfZips, ref _pathOfOriginalXML, ref _xmlLastDirectory, _zipList);
            }
        }
        private void getVisiTagPath(string _path, ref string _visiTagPath)
        {//Gets full file path of original zip files. 
            string[] subDirectories = System.IO.Directory.GetDirectories(_path);;
            int numOfSubDirectories = subDirectories.Length;
            if (numOfSubDirectories == 0)
            {
                return;
            }
            for(int a = 0; a < numOfSubDirectories; a++)
            {
                if(subDirectories[a].EndsWith("WiseTag"))
                {
                    _visiTagPath = subDirectories[a];
                }
            }
            for (int i = 0; i < numOfSubDirectories; i++)
            {
                getVisiTagPath(subDirectories[i], ref _visiTagPath);
            }
        }
        private void createNewZips(int _numOfZips, List<string> _newPaths, List<string> _zipNames, ref List<Ionic.Zip.ZipFile> _newZips, ref List<string> _pathOfNewZips, List<string> _pathOfXML, List<string> _xmlLastDir) 
        {
            lbl_Progress.Text = "Configuring new zip files";
            lbl_Progress.Update();

            if (_newPaths.Count != _numOfZips || _zipNames.Count != _numOfZips)
            {
                return;
            }

            string path;

            for (int i = 0; i < _numOfZips; i++)
            {
                path = (_newPaths[i] + "\\" + _zipNames[i]).Replace('/', '\\');
                Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile(path);
                _pathOfNewZips.Add(path);
                _newZips.Add(zip);
                zip.Save();
                for(int j = 0; j < _xmlLastDir.Count; j++)
                {
                    if(path.Contains(_xmlLastDir[j]))
                    {
                        System.IO.File.Copy(_pathOfXML[j], _newPaths[i] + "\\StudyBackup.xml", true);
                    }
                }
            }
        }
        private void createNewPaths(List<string> _pathWithoutZip, ref List<string> _newFolderPath, ref string _finalPath)
        {
            lbl_Progress.Text = "Generating new paths to zip files.";
            lbl_Progress.Update();

            string path = txtBox_Output.Text + "\\Processed_";
            string directoryLast = getLastDirectoryName(txtBox_Input.Text);
            string testPath = path + directoryLast;
            if(System.IO.Directory.Exists(testPath))
            {
                for(int i = 1; i < 999999; i++)
                {
                    string duplicatePath = path + "(" + i.ToString() + ")";
                    testPath = duplicatePath + directoryLast;
                    if(!System.IO.Directory.Exists(testPath))
                    {
                        path = duplicatePath;
                        break;
                    }
                }
            }
            _finalPath = path;
            int numOfPathsWithOutZip = _pathWithoutZip.Count;
            for (int i = 0; i < numOfPathsWithOutZip; i++)
            {
                string replacementPath;
                string extensionPath = makeRelative(_pathWithoutZip[i], txtBox_Input.Text).Replace('/', '\\');
                if (extensionPath == "")
                {
                    replacementPath = path + getLastDirectoryName(txtBox_Input.Text);
                }
                else
                {
                    replacementPath = (path + makeRelative(_pathWithoutZip[i], txtBox_Input.Text)).Replace('/', '\\');
                }
                _newFolderPath.Add(replacementPath);
                System.IO.Directory.CreateDirectory(replacementPath);
            }
        }
        private void deleteTempEntries(string _destinationPath, List<string> _zipList)
        {//Deletes anything that is not a matching zip file in the configuration text file.
            if (System.IO.Directory.Exists(_destinationPath))
            {
                string[] entries = System.IO.Directory.GetFileSystemEntries(_destinationPath);
                int numOfEntries = entries.Length;
                for (int i = 0; i < numOfEntries; i++)
                {
                    System.IO.FileAttributes attr = System.IO.File.GetAttributes(entries[i]);
                    if ((!isMatchingZip(entries[i].ToLower(), _zipList)) && ((attr & System.IO.FileAttributes.Directory) != System.IO.FileAttributes.Directory))
                    {
                        if (!entries[i].EndsWith(".xml"))
                        {
                            System.IO.File.Delete(entries[i]);
                        }
                    }
                    else if (System.IO.Directory.Exists(entries[i]))
                    {
                        System.IO.Directory.Delete(entries[i], true);
                    }
                }
            }
           else
            {
                return;
            }
        }
        private void finalCleanUp(string _directory)
        {
            foreach (var directory in System.IO.Directory.GetDirectories(_directory))
            {
                finalCleanUp(directory);
                if (System.IO.Directory.GetFiles(directory).Length == 0 && System.IO.Directory.GetDirectories(directory).Length == 0)
                {
                    System.IO.Directory.Delete(directory, false);
                }
            }
        }
        private void visiTagToText(uint[] _visiTagIDs, string _pathToVisiTag, string _destinationPath, string _fileName, Ionic.Zip.ZipFile _newZip, ref bool _isVisiTagLoadError)
        {
            string visiTagExportDirectory = _pathToVisiTag + "\\VisiTagExport";
            string fileNameWithoutWiseTag = goUpOneDirectoryRelative(goUpOneDirectoryRelative(_fileName));
            string txtFolderDestination = fileNameWithoutWiseTag + "/VisiTagExport";

            coretechlib.CartoMapReaderDotNet cmr = new coretechlib.CartoMapReaderDotNet();
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(visiTagExportDirectory);

            string allPosPath = visiTagExportDirectory + "\\AllPositions.txt";
            //string EndEmperiumPosPath = visiTagExportDirectory + "\\EndEmperium.txt";      
            string VisiTagInfoPath = visiTagExportDirectory + "\\VisiTagInfo.txt";
            string VisiTagSessionsPath = visiTagExportDirectory + "\\VisiTagSessions.txt";

            var allPositionsText = new List<string>();
            //var endEmperiumPosText = new List<string>();
            var visiTagInfoText = new List<string>();
            var visiTagSessionsText = new List<string>();
            

            int numOfFilesRead = 0;

            allPositionsText.Add("\tVisiTagID\tSessionID\tAllPosTimeStamp\tAllPosValidStatus\tX\tY\tZ");
            //endEmperiumPosText.Add("\tVisiTagID\tSessionID\tEndEmpPosTimeStamp\tEndEmpValidStatus\tX\tY\tZ");
            visiTagSessionsText.Add("\tVisiTagID\tSessionID\tStartTs\tEndTs\tPresetID\tMapID");

            for (int i = 0; i < _visiTagIDs.Length; i++)
            {
                try
                {
                    object[] loadVisiTag = cmr.LoadVisiTagData(_pathToVisiTag + "\\", _visiTagIDs[i]);
                    string statusText = (string)loadVisiTag[0];    //Success or Error
                    if (statusText == "Success")
                    {
                        numOfFilesRead++;
                        Int32 sessionID = (Int32)loadVisiTag[1];       //Session ID for synchronization with RF and CF recordings
                        string ablatingCatheterName = (string)loadVisiTag[3];  //Ablation catheter name as String value
                        Int32[,] ablatingChannelsID = (Int32[,])loadVisiTag[4];    //unipolar ID, bipolar ID (if not a bipolar channel, bipolarID = -1)
                        Int32[] startEndTimestamp = (Int32[])loadVisiTag[5];   //Start and end VisiTag timestamps
                        Int32[] mapIDs = (Int32[])loadVisiTag[6];  //the selected map when the ablation was performed
                        Int32[][,] ablationIntervalsPerChannel = (Int32[][,])loadVisiTag[7];   //used in multi electrode ablation - single list for focal catheter
                        bool isForceCatheter = (bool)loadVisiTag[8];   // Indicates if this is a Biosense batheter with eeprom or external on(Only Bionsense Catheter are allowed)
                        bool isMultiElectrodeCatheter = (bool)loadVisiTag[9];  //Used for nMarq catheter
                        bool isTGA = (bool)loadVisiTag[10];    //Is temperature guided ablation used.
                        Int32 presetID = (Int32)loadVisiTag[11];   //User setup settings ID. The settings stored in CARTO Data Table - CONFIG_VISI_TAG_PRESET_TABLE
                        string[] touchAtStartSession = (string[])loadVisiTag[12];  //Could be one of: "In Touch", "Not In Touch", "Unknown", "Not Supported" -this is part of the TPI capability
                        //All the catheter positions included in this session. 
                        UInt32[][][] allPositionTimestamps = (UInt32[][][])loadVisiTag[13];
                        Double[][][,] allPositions = (Double[][][,])loadVisiTag[14];
                        Int32[][][] allPositionValidStatuses = (Int32[][][])loadVisiTag[15];   //Valid status per timestamp
                        //All the catheter positions included in this session. 
                        //UInt32[][][] endExperiumTimestamps = (UInt32[][][])loadVisiTag[16];
                        //Double[][][,] endExperiumPositions = (Double[][][,])loadVisiTag[17];
                        //Int32[][][] endExperiumValidStatuses = (Int32[][][])loadVisiTag[18];
                        //Three typed data indices for synchronization between Visitag Vatheter positions, RF Ablation data, and Catheter Force Data.
                        for (int x = 0; x < allPositions.Length; x++)
                        {
                            for (int y = 0; y < allPositions[x].Length; y++)
                            {
                                for (int z0 = 0; z0 < allPositions[x][y].GetLength(0); z0++)
                                {
                                    for (int z1 = 0; z1 < allPositions[x][y].GetLength(1); z1 += 3)
                                    {
                                        allPositionsText.Add("\t" + _visiTagIDs[i] + "\t" + sessionID + "\t" + allPositionTimestamps[x][y][z0] + "\t" + allPositionValidStatuses[x][y][z0] + "\t" + string.Format("{0:0.000}", allPositions[x][y][z0, z1]) + "\t" + string.Format("{0:0.000}", allPositions[x][y][z0, z1 + 1]) + "\t" + string.Format("{0:0.000}", allPositions[x][y][z0, z1 + 2]));
                                    }
                                }
                            }
                        }

                        /*for (int x = 0; x < endExperiumPositions.Length; x++)
                        {
                            for (int y = 0; y < endExperiumPositions[x].Length; y++)
                            {
                                for (int z0 = 0; z0 < endExperiumPositions[x][y].GetLength(0); z0++)
                                {
                                    for (int z1 = 0; z1 < endExperiumPositions[x][y].GetLength(1); z1 += 3)
                                    {
                                        endEmperiumPosText.Add("\t" + _visiTagIDs[i] + "\t" + sessionID + "\t" + endExperiumTimestamps[x][y][z0] + "\t" + endExperiumValidStatuses[x][y][z0] + "\t" + string.Format("{0:0.000}", endExperiumPositions[x][y][z0, z1]) + "\t" + string.Format("{0:0.000}", endExperiumPositions[x][y][z0, z1 + 1]) + "\t" + string.Format("{0:0.000}", endExperiumPositions[x][y][z0, z1 + 2]));
                                    }
                                }
                            }
                        }*/
                        visiTagSessionsText.Add("\t" + _visiTagIDs[i] + "\t" + sessionID + "\t" + startEndTimestamp[0] + "\t" + startEndTimestamp[1] + "\t" + presetID + "\t" + mapIDs[0]);
                        if (i == 0)
                        {
                            string ablatingChannelsIDString = "AblatingChannelsIDs: ";
                            visiTagInfoText.Add("AblatingCatheterName: " + ablatingCatheterName);
                            if (touchAtStartSession.Length > 0)
                            {
                                visiTagInfoText.Add("TouchAtStartSession: " + touchAtStartSession[0]);
                            }
                            else
                            {
                                visiTagInfoText.Add("TouchAtStartSession: null");
                            }
                            for (int j = 0; j < ablatingChannelsID.GetLength(0); j++)
                            {
                                for (int k = 0; k < ablatingChannelsID.GetLength(1); k++)
                                {
                                    ablatingChannelsIDString += ablatingChannelsID[i, j].ToString() + "\t";
                                }
                            }
                            visiTagInfoText.Add(ablatingChannelsIDString);
                            visiTagInfoText.Add("IsForceCatheter: " + isForceCatheter);
                            visiTagInfoText.Add("IsMultiElectrodeCatheter: " + isMultiElectrodeCatheter);
                            visiTagInfoText.Add("IsTGA: " + isTGA);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.SEHException)
                {
                }
                catch(System.IndexOutOfRangeException)
                {
                }
            }

            System.IO.File.WriteAllLines(allPosPath, allPositionsText);
            //System.IO.File.WriteAllLines(EndEmperiumPosPath, endEmperiumPosText.ToArray());
            System.IO.File.WriteAllLines(VisiTagSessionsPath, visiTagSessionsText.ToArray());
            System.IO.File.WriteAllLines(VisiTagInfoPath, visiTagInfoText.ToArray());
           
            if (!System.IO.Directory.Exists(txtFolderDestination))
            {
                if (!txtBox_Input.Text.Contains("Processed_"))
                {
                    _newZip.AddItem(visiTagExportDirectory, txtFolderDestination);
                }
            }
            if(numOfFilesRead == 0)
            {
                _isVisiTagLoadError = true;
            }
            _newZip.Save();
            cmr = null;
        }
        private void meshToText(string _fileName, string _destinationPath, Ionic.Zip.ZipFile _newZip)
        {
            string path = _destinationPath + "\\" + _fileName;
            string meshName = getMeshFileName(_fileName);
            string txtFolderDestination = goUpOneDirectoryRelative(_fileName) + "/" + meshName.Replace('.', '_');
            string newDirectoryPath = _destinationPath + "\\" + meshName.Replace('.', '-');
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(newDirectoryPath);
            coretechlib.CartoMapReaderDotNetOld cmr = new coretechlib.CartoMapReaderDotNetOld();
            try
            {
                System.Object[] cmrLoad = cmr.LoadCartoMap(path);
                System.Object[] cmrInfo = cmr.GetCartoMapInfo();
                System.Object[] cmrExtend = cmr.GetCartoMapExtendedData();

                string vertexPath = newDirectoryPath + "\\" + "VertexInfo.txt";
                string trianglePath = newDirectoryPath + "\\" + "TriangleInfo.txt";
                string colorGroupPath = newDirectoryPath + "\\" + "colorGroupInfo.txt";
                string huePath = newDirectoryPath + "\\" + "hueInfo.txt";
                string mapInfoPath = newDirectoryPath + "\\" + "mapInfo.txt";
                string verticesStatePath = newDirectoryPath + "\\" + "verticesStateInfo.txt";
                string triangleStatePath = newDirectoryPath + "\\" + "triangleStateInfo.txt";
                string vertexAttrPath = newDirectoryPath + "\\" + "vertexAttrInfo.txt";

                if (cmrLoad[10] != null)
                {
                    string[] vertexInfo = cmrLoad[10].ToString().Split(',');
                    System.IO.File.WriteAllLines(vertexPath, vertexInfo);
                }
                if (cmrLoad[11] != null)
                {
                    string[] triangleInfo = cmrLoad[11].ToString().Split(',');
                    System.IO.File.WriteAllLines(trianglePath, triangleInfo);
                }
                if (cmrLoad[12] != null)
                {
                    string[] colorGroupinfo = cmrLoad[12].ToString().Split(',');
                    System.IO.File.WriteAllLines(colorGroupPath, colorGroupinfo);
                }
                if (cmrLoad[13] != null)
                {
                    string[] hueInfo = cmrLoad[13].ToString().Split(',');
                    System.IO.File.WriteAllLines(huePath, hueInfo);
                }
                if (cmrInfo[21] != null)
                {
                    string[] mapInfo = cmrInfo[21].ToString().Split(',');
                    System.IO.File.WriteAllLines(mapInfoPath, mapInfo);
                }
                if (cmrExtend[2] != null)
                {
                    string[] verticesState = cmrExtend[2].ToString().Split(',');
                    System.IO.File.WriteAllLines(verticesStatePath, verticesState);
                }
                if (cmrExtend[3] != null)
                {
                    string[] triangleState = cmrExtend[3].ToString().Split(',');
                    System.IO.File.WriteAllLines(triangleStatePath, triangleState);
                }
                if (cmrExtend[4] != null)
                {
                    string[] vertexAttributes = cmrExtend[4].ToString().Split(',');
                    System.IO.File.WriteAllLines(vertexAttrPath, vertexAttributes);
                }

                if (!System.IO.Directory.Exists(txtFolderDestination) && cmrLoad[10] != null)
                {
                    if (!txtBox_Input.Text.Contains("Processed_"))
                    {
                        _newZip.AddItem(newDirectoryPath, txtFolderDestination);
                    }
                }
            }
            catch (System.Runtime.InteropServices.SEHException)
            {
                //Catch so we can prevent error. 
            }
            cmr = null;
        }
        private void extractData(string _pathToZip, string _destinationPath, Ionic.Zip.ZipFile _newZip, List<string> _folderList, List<string> _fileList, List<string> _zipList, List<string> _internalZipList,  string _pathOfNewZip, bool _keepUnproccessedData, ref List<string> _summaryText)
        {
            string pathToAdd = makeRelative(_pathToZip, txtBox_Input.Text);
            string visiTagRelativePath = "";
            int padLength = 105 - pathToAdd.Length;
            int numOfExtractions = 0;
            bool isVisiTagLoadError = false;

            if (padLength < 0)
            {
                padLength = 0;
            }
            try
            {
                using (Ionic.Zip.ZipFile zip = Ionic.Zip.ZipFile.Read(_pathToZip))
                {
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed;

                    foreach (Ionic.Zip.ZipEntry ze in zip)
                    {//Look at each entry and see if it is a match to extract.
                        string fileName = ze.FileName;
                        bool isThereItemToExtract = isThereMatch(fileName.ToLower(), _folderList, _fileList);
                        bool isItInternalZip = isMatchingZip(fileName.ToLower(), _internalZipList);
                
                        if(fileName.EndsWith("WiseTag/"))
                        {
                            visiTagRelativePath = ze.FileName;
                        }

                        if (isThereItemToExtract || isItInternalZip)
                        {
                            numOfExtractions++;
                            string pathOfFileToExtract = (_destinationPath + "\\" + ze.FileName).Replace('/', '\\');
                            string pathInNewZipFile = goUpOneDirectoryRelative(ze.FileName);
                            ze.Extract(_destinationPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                            if (fileName.EndsWith(".mesh"))
                            {
                                meshToText(fileName, _destinationPath, _newZip);
                                if(_keepUnproccessedData)
                                {
                                    _newZip.AddItem(pathOfFileToExtract, pathInNewZipFile);
                                }
                            }
                            else
                            {
                                if (!(fileName.Contains("/WiseTag/") && (_keepUnproccessedData == false)))
                                {
                                    _newZip.AddItem(pathOfFileToExtract, pathInNewZipFile);
                                }
                            }
                            if (isItInternalZip)
                            {
                                using (Ionic.Zip.ZipFile internalZip = Ionic.Zip.ZipFile.Read(pathOfFileToExtract))
                                {
                                    internalZip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed;
                                    foreach (Ionic.Zip.ZipEntry zie in internalZip)
                                    {
                                        string internalFileName = zie.FileName;
                                        if (isThereMatch(internalFileName.ToLower(), _folderList, _fileList))
                                        {
                                            string extractionPath = goUpOneDirectoryAbsolute(pathOfFileToExtract);  //path of where to extract the internal zip contents. 
                                            string filePath = (extractionPath + "\\" + zie.FileName).Replace('/', '\\');    //path of directory to be saved inside the zipfile
                                            zie.Extract(extractionPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                                            _newZip.AddItem(filePath, pathInNewZipFile);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    _newZip.Save();
                }
            }
            catch(Ionic.Zip.BadReadException)
            {//If zip file is unreadable.
                string name = getLastDirectoryName(_pathToZip);
                if (name.Contains(".zip"))
                {
                    _summaryText.Add(pathToAdd + "Unable to Read".PadLeft(padLength));
                    txtbox_Summary.AppendText(pathToAdd + "Unable to Read".PadLeft(padLength) + "\r\n");
                    txtbox_Summary.Update();
                    System.IO.File.Delete(_newZip.Name);
                }
            }
            catch(Ionic.Zip.ZipException)
            {//If zip file is unreadable.
                string name = getLastDirectoryName(_pathToZip);
                if (name.Contains(".zip"))
                {
                    _summaryText.Add(pathToAdd + "Unable to Read".PadLeft(padLength));
                    txtbox_Summary.AppendText(pathToAdd + "Unable To Read".PadLeft(padLength) + "\r\n");
                    txtbox_Summary.Update();
                    System.IO.File.Delete(_newZip.Name);
                }
            }
            string visiTagDirectory = "";
            getVisiTagPath(_destinationPath, ref visiTagDirectory);
            try
            {
                if (System.IO.Directory.Exists(visiTagDirectory))
                {
                    coretechlib.CartoMapReaderDotNet cmrdn = new coretechlib.CartoMapReaderDotNet();
                    string visiTagDirectoryOpen = visiTagDirectory + "\\";
                    uint[] visiTagIDs = cmrdn.GetAllVisiTagIDs(visiTagDirectoryOpen);
                    if (visiTagIDs.Length != 0)
                    {
                        visiTagToText(visiTagIDs, visiTagDirectory, _destinationPath, visiTagRelativePath, _newZip, ref isVisiTagLoadError);
                    }
                    cmrdn = null;
                }
            }
            catch(System.Runtime.InteropServices.SEHException)
            {

            }
            if (numOfExtractions == 0)
            {
                _summaryText.Add(pathToAdd + "No Data".PadLeft(padLength));
                txtbox_Summary.AppendText(pathToAdd + "No Data".PadLeft(padLength) + "\r\n");
                txtbox_Summary.Update();
                System.IO.File.Delete(_newZip.Name);
            }
            else if(isVisiTagLoadError)
            {
                _summaryText.Add(pathToAdd + "Unable to Read VisiTag".PadLeft(padLength));
                txtbox_Summary.AppendText(pathToAdd + "Unable to read VisiTag".PadLeft(padLength) + "\r\n");
                txtbox_Summary.Update();
            }
            else
            {
                _summaryText.Add(pathToAdd + "Success".PadLeft(padLength));
                txtbox_Summary.AppendText(pathToAdd + "Success".PadLeft(padLength) + "\r\n");
                txtbox_Summary.Update();
            }
            deleteTempEntries(_destinationPath, _zipList);
        }
    }
}

