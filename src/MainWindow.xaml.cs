using System;
using System.Windows;
using System.Speech.Synthesis;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using IWshRuntimeLibrary;

#region Reference Links

//https://www.codeproject.com/Articles/146757/Add-Remove-Startup-Folder-Shortcut-to-Your-App

#endregion

namespace Challenge_TalkClock
{
    public partial class MainWindow : Window
    {
        Timer timer = new Timer();
        static EventLog log = new EventLog("System");
        private NotifyIcon myTray = new NotifyIcon();

        //NEW: Option to start automatically
        string appname = Assembly.GetExecutingAssembly().FullName.Remove(Assembly.GetExecutingAssembly().FullName.IndexOf(","));
        private string DesktopPathName;
        private string StartupPathName;
        //


        public MainWindow()
        {
            InitializeComponent();
            Hide();

            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                Close();
                return;

            } //Allows only one persistance of the program to run

            CreateIcon();
            MyWinFormsTimer(true);
            DesktopPathName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), appname + ".lnk");
            StartupPathName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), appname + ".lnk");

            if (Properties.Settings.Default.firstTime)
            {
                myTray.ShowBalloonTip(1000, "Talking Clock", "Your talking clock is now active and working!", ToolTipIcon.None);
                Properties.Settings.Default.firstTime = false;
                Properties.Settings.Default.Save();

            } //startup

        }

        #region Menu

        public void CreateIcon()
        {
            ContextMenu myMenu = new ContextMenu();
            MenuItem myItem1 = new MenuItem();
            MenuItem myItem2 = new MenuItem();
            MenuItem myItem3 = new MenuItem();

            //creates the icon and message
            myTray.Icon = new Icon(@"Clock.ico");
            myTray.Visible = true;
            myTray.Text = "Talking Clock";

            //creates a list of menu items in context menu
            myMenu.MenuItems.AddRange(new MenuItem[] { myItem1, myItem2, myItem3 });
            myItem1.Index = 0;
            myItem1.Text = "Exit";
            myItem1.Click += new EventHandler(ExitClicked);
            myItem2.Index = 1;
            myItem2.Text = "About";
            myItem2.Click += new EventHandler(AboutClicked);
            myItem3.Index = 2;
            myItem3.Text = "Launch Automatically";
            myItem3.Click += new EventHandler(LaunchClicked);
            myTray.ContextMenu = myMenu;


        } //Creates a simple icon in the system tray

        private void AboutClicked(object sender, EventArgs e)
        {
            myTray.ShowBalloonTip(1000, "About Talking Clock", "This program will keep you aware of the time while doing other things.", ToolTipIcon.Info);

        } //Info on program

        private void LaunchClicked(object sender, EventArgs e)
        {
            if (!Properties.Settings.Default.startUp)
            {
                    CreateShortcut(StartupPathName, true); //Create a shortcut in the startup folder
                    myTray.ShowBalloonTip(300, "Talking Clock", "Now running on startup", ToolTipIcon.None);
            }
            else
            {
                myTray.ShowBalloonTip(300, "Talking Clock", "No longer running on startup.", ToolTipIcon.None);
                CreateShortcut(StartupPathName, false); //Remove the shortcut in the startup folder
            }
            Properties.Settings.Default.startUp = !Properties.Settings.Default.startUp;
            Properties.Settings.Default.Save();

        } //Startup user control

        private void ExitClicked(object sender, EventArgs e)
        {
            myTray.Visible = false;
            Close();
        } //Exit Option

        private void CreateShortcut(string shortcutPathName, bool create)
        {
            if (create)
            {
                try
                {
                    string shortcutTarget = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, appname + ".exe");
                    WshShell myShell = new WshShell();
                    WshShortcut myShortcut = (WshShortcut)myShell.CreateShortcut(shortcutPathName);
                    myShortcut.TargetPath = shortcutTarget; //The exe file this shortcut executes when double clicked
                    myShortcut.IconLocation = shortcutTarget + ",0"; //Sets the icon of the shortcut to the exe`s icon
                    myShortcut.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //The working directory for the exe
                    myShortcut.Arguments = ""; //The arguments used when executing the exe
                    myShortcut.Save(); //Creates the shortcut
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
            else
            {
                try
                {
                    if (System.IO.File.Exists(shortcutPathName))
                        System.IO.File.Delete(shortcutPathName);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }

        #endregion

        #region Timer Creation

        public void MyWinFormsTimer(bool On = false)
        {
            if (!On) // default
            {
                timer.Enabled = false;
            }
            else
            {
                timer.Enabled = true;
                timer.Interval = MilliSecondsLeftTilTheHour();
                timer.Tick += new EventHandler(Timer_Tick);
            }

        } // Sets up or disables a timer to occur at regular interval

        private int MilliSecondsLeftTilTheHour()
        {
            int interval;
            int minutesRemaining = 59 - DateTime.Now.Minute;
            int secondsRemaining = 59 - DateTime.Now.Second;
            interval = ((minutesRemaining * 60) + secondsRemaining) * 1000;
            if (interval == 0) //quick calculation when caught up
            {
                interval = 60 * 60 * 1000;
            }
            return interval;
        } //returns an integer in miliseconds left until the next hour

        private void Timer_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            now = now.AddMinutes(1);
            int preSpeech = now.Hour;
            if (preSpeech == 24) preSpeech = 0;
            String postSpeech;
            postSpeech = "The time is ";
            if (preSpeech == 12)
            {
                postSpeech += "Noon.";
            }
            else if (preSpeech == 0)
            {
                postSpeech += "Midnight.";
            }
            else if ((preSpeech - 12) > 1)
            {
                preSpeech -= 12;
                postSpeech += preSpeech.ToString();
                postSpeech += " pm";
            }
            else
            {
                postSpeech += preSpeech.ToString();
                postSpeech += " am";
            }
            SpeakNow(postSpeech);
            timer.Interval = MilliSecondsLeftTilTheHour();
        } //Speaks time every hour depending on timer expiration

        private static void SpeakNow(string String)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.SelectVoiceByHints(VoiceGender.Female);
            synthesizer.Rate = 1;     // -10...10
            synthesizer.SpeakAsync(String);
        } //Synthesizes string into audio

        private void log_EntryWritten(object sender, EntryWrittenEventArgs e)
        {
            if (e.Entry.InstanceId == 1 && e.Entry.EntryType == EventLogEntryType.Information)
                timer.Interval = MilliSecondsLeftTilTheHour();
        } //Keep Program on track when time is changed

        #endregion
    }
}

