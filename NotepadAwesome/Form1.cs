using System;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;


namespace NotepadAwesome
{
    public partial class Form1 : Form
    {
        // Parser
        [DllImport("shell32.dll", SetLastError = true)]
        static extern IntPtr CommandLineToArgvW(
            [MarshalAs(UnmanagedType.LPWStr)] string lpCmdLine, out int pNumArgs);

        public string userDir = "";
        public string appName = System.Windows.Forms.Application.ProductName;
        public string cmdParse = "*";
        public string title = "";
        public string filename = "";
        public bool txtChanged = false;
        public string[] cmdLineSaves;
        int numSaves = 0;

        public Form1()
        {
            InitializeComponent();
            //cmdLineSaves = new string[] { };
            Program.numForms =+ 1;
            // Get user directory - set directory
            // Right now - manually created directory
            userDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + appName;
            Directory.SetCurrentDirectory(userDir);

            // Set the title
            getTitle();
        }

        private void Form1_Load(object sender, EventArgs e)
        {   
            mainNotes.Text = userDir;
        }

        private void autoSave()
        {
            commandText.Text = "";
            commandText.Text = "Saved";
            txtChanged = false;
        }

        private void commandText_TextChanged(object sender, EventArgs e)
        {
        }

        private void commandText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                // This is where we parse commands - after this we need to move to a class
                string cmdText = "";
                cmdText = commandText.Text;
                //var args = new string[2];
                var retArg = CommandLineToArgs(cmdText);
                // Save the commands
                //cmdLineSaves[numSaves] = cmdText;
                //numSaves++;

                // Need a robust parser - but now just take it all
                switch (retArg[0].ToLower())
                {
                    case "n": // New window
                        Form1 newPad = new Form1();
                        newPad.Show();
                        commandText.Text = "";
                        break;
                    case "e": // Exit
                        this.Close();
                        if(Program.numForms == 0)
                        {
                            System.Windows.Forms.Application.Exit();
                        }
                        break;
                    case "c": // Clear ecommand line
                        commandText.Text = "";
                        break;
                    case "t": // Set window title
                        setTitle(retArg[1]);
                        commandText.Text = "";
                        break;
                    case "s": // save file
                        saveFile();
                        commandText.Text = "";
                        break;
                    case "f": // get files in directory
                        commandText.Text = "";
                        // get files
                        string[] fileNames = Directory.GetFiles(userDir, "*.txt");
                        // If we want it in a new window (switch -n)
                        if (retArg.Length == 1)
                        {
                            mainNotes.Lines = fileNames;
                        }
                        else
                        {
                            if (retArg[1] == "-n")
                            {
                                // open window sending in files
                                newPad = new Form1();
                                newPad.mainNotes.Lines = fileNames;
                                newPad.Show();
                            }
                        }
                        break;
                    case "d": // Delete file
                        try
                        {
                            File.Delete(userDir + "\\" + retArg[1] + ".txt");
                        } catch
                        {
                            commandText.Text = commandText.Text + " " + "file doesn't exist";
                        }
                        commandText.Text = "";
                        break;
                    case "cls": // Clear main screen
                        mainNotes.Text = "";
                        commandText.Text = "";
                        break;
                    case "em": // Send email with text
                        // Make sure we have enough retArg
                        int length = retArg.Length;
                        Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                        Microsoft.Office.Interop.Outlook._MailItem oMailItem = (Microsoft.Office.Interop.Outlook._MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        switch (length)
                        {
                            case (1): // Nothing - just open mail with text in the body
                                oMailItem.Body = mainNotes.Text;
                                break;
                            case (2): // Always - the to
                                oMailItem.Body = mainNotes.Text;
                                oMailItem.To = retArg[1];
                                break;
                            case (3): // Subject
                                oMailItem.Body = mainNotes.Text;
                                oMailItem.To = retArg[1];
                                oMailItem.Subject = retArg[2];
                                break;
                        }
                        oMailItem.Display(true);
                        commandText.Text = "";
                        break;
                    case "of": // Open a file
                        // Open a new window with
                        if(retArg.Length > 1)
                        {
                            try { 
                                string file = File.ReadAllText(retArg[1] + ".txt");
                                if (file.Length > 1)
                                {
                                    Form1 form = new Form1();
                                    form.mainNotes.Text = file;
                                    form.setTitle(retArg[1]);
                                    form.Show();
                                    commandText.Text = "";
                                }
                            } catch
                                {
                                    commandText.Text = commandText.Text + ' ' + "wrong file name.";
                                }
                        }
                        break;
                    case "h": // open the help file to a new window - need to ensure it exists
                        try
                        {
                            string file = File.ReadAllText("helpfile.txt");
                            Form1 form = new Form1();
                            form.mainNotes.Text = file;
                            form.setTitle("helpfile");
                            form.Show();
                        } catch
                        {
                            commandText.Text = commandText.Text + " " + "Help file doesn't exist";
                        }
                        break;
                    case "cs":// Command line saves
                        break;
                }
            }
        }

        private void openTextFile(string fileName)
        {
            string file = File.ReadAllText(fileName);
            mainNotes.Text = fileName;
        }

        private void getTitle()
        {
            // On new - title = Day+General Time (afternoon)+yaer+_+tiimestamp
            string day = DateTime.Now.DayOfWeek.ToString();
            string month = String.Format("{0:MMMM}", DateTime.Now);
            string appxTime = DateTime.Now.Hour.ToString();
            string hour = DateTime.Now.Hour.ToString();
            string minute = DateTime.Now.Minute.ToString();
            string sec = DateTime.Now.Second.ToString();
            
            // Get general time
            int hourInt = DateTime.Now.Hour;
            if((hourInt > 6) && (hourInt <  11))
            {
               appxTime = "Morning";
            }
            else if ((hourInt > 11) && (hourInt < 16))
            {
                appxTime = "Afternoon";
            }
            else if ((hourInt > 16) && (hourInt < 22))
            {
                appxTime = "Evening";
            }
            else if((hourInt > 22) && (hourInt< 6))
            {
                appxTime = "DeadNight";
            }

            setTitle(day + '_' + appxTime  + '_' + month + '_' + hour + '_' + minute + '_' + sec);
        }

        private void setTitle(string newTitle)
        {
            title = newTitle;
            this.Text = title;
            filename = title + ".txt";
        }

        private void saveFile()
        {
            // Save everything in the text box - override the existing file if necessary
            string fileWrite = userDir + '\\' + filename;
            using (StreamWriter outputfile = new StreamWriter(fileWrite))
            {
                outputfile.Write(mainNotes.Text);
            }
        }

        private void mainNotes_TextChanged(object sender, EventArgs e)
        {
            // If the text has been changed AND it has been x minutes since last save, save the document
            txtChanged = true;
        }
        
        private void Form1_Load_1(object sender, EventArgs e)
        {
            //mainNotes.Text = userDir;
        }

        public static string[] CommandLineToArgs(string commandLine)
        {
            int argc;
            var argv = CommandLineToArgvW(commandLine, out argc);
            if (argv == IntPtr.Zero)
                throw new System.ComponentModel.Win32Exception();
            try
            {
                var args = new string[argc];
                for (var i = 0; i < args.Length; i++)
                {
                    var p = Marshal.ReadIntPtr(argv, i * IntPtr.Size);
                    args[i] = Marshal.PtrToStringUni(p);
                }

                return args;
            }
            finally
            {
                Marshal.FreeHGlobal(argv);
            }
        }

        private void pasteToolStripButton_Click(object sender, EventArgs e)
        {
            mainNotes.Paste();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainNotes.Paste();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainNotes.Copy();
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            mainNotes.Copy();
        }

        private void cutToolStripButton_Click(object sender, EventArgs e)
        {
            mainNotes.Cut();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainNotes.Cut();
        }

        private void mainNotes_KeyDown(object sender, KeyEventArgs e)
        {
            // Check if the tab/ctrl is pressed - if so, move to commandText
            if (e.KeyData == Keys.F1){
                commandText.Focus();
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            saveFile();
        }
    }
}