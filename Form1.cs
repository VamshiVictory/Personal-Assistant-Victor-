using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Xml;
using System.Speech.AudioFormat;
using System.Threading;
using System.Management;
using System.Diagnostics;
using System.Windows;

namespace Victor
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer VICTOR = new SpeechSynthesizer();
        string ProcWindow;        
        Random rnd = new Random();

        public Form1()
        {
            InitializeComponent();            
        }
        private void Form1_Load_1(object sender, EventArgs e)
        {
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammar(new DictationGrammar());
            _recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"E:\commands.vic")))));
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);            
        }

       public void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            DateTime t2 = DateTime.Now;
            DateTime t1 = Convert.ToDateTime("00:00:01 AM");
            DateTime t3 = Convert.ToDateTime("11:59:59 AM");
            DateTime t4 = Convert.ToDateTime("12:00:01 PM");
            DateTime t5 = Convert.ToDateTime("03:59:59 PM");
            DateTime t6 = Convert.ToDateTime("04:00:01 PM");
            DateTime t7 = Convert.ToDateTime("07:59:59 PM");
            DateTime t8 = Convert.ToDateTime("08:00:01 PM");
            DateTime t9 = Convert.ToDateTime("11:59:59 PM");
            int ranNum = rnd.Next(1, 10);
            string speech = e.Result.Text;
            String battstatus;
            PowerStatus pwr = SystemInformation.PowerStatus;
            battstatus = SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
            String battlife;
            battlife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            String batt;
            batt = SystemInformation.PowerStatus.BatteryLifeRemaining.ToString();
            String num = SystemInformation.PowerStatus.ToString();
           string sam =  SystemInformation.PowerStatus.PowerLineStatus.ToString();
            switch (speech)
            {                
                //GREETINGS                              
                case "Thank you Victor":
                case "Thanks Victor":
                    VICTOR.Speak("Anytime for you Sir");
                    break; 
                case "How are you Victor?":
                    VICTOR.Speak("I am very good Sir, Thanks for asking");
                    VICTOR.Speak("How are you Sir?");
                    break;
                case "I am good Victor":
                    VICTOR.Speak("I know you will be, Sir!");
                    break;
                case "Good Morning Victor":                  
                    if (t2==t1 || t2==t3 || t2 < t3) 
                        VICTOR.Speak("Very Good Morning Sir!");
                    else if (t2 == t4 || t2 == t5 || t2 < t5)
                    {
                        VICTOR.Speak("Sorry, It's Afternoon Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t6 || t2 == t7 || t2 < t7)
                    {
                        VICTOR.Speak("Sorry, It's Evening Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t8 || t2 == t9 || t2 < t9)
                    {
                        VICTOR.Speak("Sorry, It's Night time Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    break;
                case "Good Afternoon Victor":
                    if (t2 == t1 || t2 == t3 || t2 < t3)
                    {
                        VICTOR.Speak("Sorry, It's Morning Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t4 || t2 == t5 || t2 < t5)   
                        VICTOR.Speak("Very Good Afternoon Sir");
                    else if (t2 == t6 || t2 == t7 || t2 < t7)
                    {
                        VICTOR.Speak("Sorry, It's Evening time Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t8 || t2 == t9 || t2 < t9)    
                    {
                        VICTOR.Speak("Sorry, It's Night time Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    break;
                     case "Good Evening Victor":
                    if (t2 == t1 || t2 == t3 || t2 < t3)
                    {
                        VICTOR.Speak("Sorry, It's Morning Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t4 || t2 == t5 || t2 < t5)
                    {
                        VICTOR.Speak("Sorry, It's Afternoon Sir");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                    else if (t2 == t6 || t2 == t7 || t2 < t7)
                        VICTOR.Speak("Very Good Evening Sir");
                    else if (t2 == t8 || t2 == t9 || t2 < t9)
                    {
                        VICTOR.Speak("Sorry, It's Night time Sir!");
                        VICTOR.Speak("I suggest u to please adjust your time, Sir");
                    }
                        break;
                     case "Good Night Victor":
                        if (t2 == t1 || t2 == t3 || t2 < t3)
                        {
                            VICTOR.Speak("Sorry, It's Morning time Sir!");
                            VICTOR.Speak("I suggest u to please adjust your time, Sir");
                        }
                        else if (t2 == t4 || t2 == t5 || t2 < t5)
                        {
                            VICTOR.Speak("Sorry, It's Afternoon time Sir!");
                            VICTOR.Speak("I suggest u to please adjust your time, Sir");
                        }
                        else if (t2 == t6 || t2 == t7 || t2 < t7)
                        {
                            VICTOR.Speak("Sorry, It's Evening time Sir!");
                            VICTOR.Speak("I suggest u to please adjust your time, Sir");
                        }
                        else if (t2 == t8 || t2 == t9 || t2 < t9)
                        {
                            VICTOR.Speak("Good Night Sir, Have a nice sleep");
                            Close();
                        }
                      break;
                    //Files 

                case "Open Movies":
                        System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Movies");
                        break;
                case "Open Special":
                        System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Videos\Special");
                        break;
                case "Open Drive":                       
                       System.Diagnostics.Process.Start("explorer", "::{20d04fe0-3aea-1069-a2d8-08002b30309d}");
                       break;
                case "Open Videos":                       
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Videos");
                       break;
                case "Open Softwares":                      
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Softwares\My Softwares");
                       break;
                case "Open College":
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Softwares\Docs\College Stuff");
                       break;
                case "Open Music":                     
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Music");
                       break;
                case "Open Photos":
                case "Open Pictures":
                case "Open Album":                       
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Volume(E)\Pictures");
                       break;
                case "Open Downloads":                       
                       System.Diagnostics.Process.Start("explorer.exe", @"F:\Downloads");
                       break;
                case "Open Projects":
                       System.Diagnostics.Process.Start("explorer.exe", @"E:\Documents\Visual Studio 2012\Projects");
                       break;
                case "My gmail":
                       SendKeys.Send("vamshivictory777@gmail.com");
                       break;
                case "Message":
                       SendKeys.Send("Hii, ");
                       SendKeys.Send("How are you?");
                       SendKeys.Send("{ENTER}");
                       break;
                case "Wish":
                       SendKeys.Send("Many( )many( )more( )happy( )returns( )of( )the( )day");
                       SendKeys.Send("{ENTER}");
                       break;
                case "Space":
                       SendKeys.Send("( )");
                       break;
                case "Show Properties":
                       SendKeys.Send("%{ENTER}");
                       break;
                case "One":
                           SendKeys.Send("{DOWN}");
                           SendKeys.Send("{UP}");
                           break;
                case "Two":
                           for (int i = 1; i <= 1; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Three":
                           for (int i = 1; i <= 2; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Four":
                           for (int i = 1; i <= 3; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Five":
                           for (int i = 1; i <= 4; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Six":
                           for (int i = 1; i <= 5; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Seven":
                       for (int i = 1; i <= 6; i++)
                           SendKeys.Send("{RIGHT}");
                           break;
                case "Nine":
                           for (int i = 1; i <= 8; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Ten":
                           for (int i = 1; i <= 9; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Eleven":
                           for (int i = 1; i <= 10; i++)
                               SendKeys.Send("{RIGHT}");
                           break;
                case "Refresh":
                       SendKeys.Send("{F5}");
                       break;

                case "Close Window":
                       SendKeys.Send("%{F4}");
                       break;
                
                case "Minimize":
                       SendKeys.Send("%( )n");
                       break;

                case "Maximize":
                       SendKeys.Send("%( )x");
                       break;
                case "Restore Window":
                       SendKeys.Send("%( )r");
                       break;
                case "Close Tab":
                    SendKeys.Send("^w");
                    break;
                case "New Tab":
                    SendKeys.Send("^t");
                    break;
                case "Next Tab":
                    SendKeys.Send("^{tab}");
                    break;
                case "Previous Tab":
                    SendKeys.Send("^+tab");
                    break;
                case "New Window":
                case "Next song":
                    SendKeys.Send("^n");
                    break;
                case "Select All":
                    SendKeys.Send("^a");
                    break;
                case "Copy":
                    VICTOR.Speak("Copying Selected");
                    SendKeys.Send("^c");
                    break;
                case "Paste":
                    SendKeys.Send("^v");   
                    break;
                case "Cut":
                    VICTOR.Speak("Cutting Selected");
                    SendKeys.Send("^x");
                    break;
                case "Down":
                    SendKeys.Send("{DOWN}");
                    break;
                case "Up":
                    SendKeys.Send("{UP}");
                    break;
                case "Next":
                    SendKeys.Send("{RIGHT}");
                    break;
                case "Previous":
                    SendKeys.Send("{LEFT}");
                    break;
                case "Delete":
                    SendKeys.Send("{DELETE}");
                    break;
                case "Back":
                    SendKeys.Send("{BS}");
                    break;
                case "Pause":
                case "Stop":
                case"Resume":
                    SendKeys.Send("( )");
                    break;
                case "Yes":
                    SendKeys.Send("{ENTER}");
                    break;
                case "Delete Permanently":                    
                    SendKeys.Send("+{DEL}");
                    break;
                case "Forward":
                    SendKeys.Send("%{RIGHT}");
                    break;
                case "Backward":
                    SendKeys.Send("%{LEFT}");
                    break;
                case "Skip forward":
                    SendKeys.Send("^{RIGHT}");
                    break;
                case "Skip backward":
                    SendKeys.Send("^{LEFT}");
                    break;
                case "Save":
                    VICTOR.Speak("Saving");
                    SendKeys.Send("^s");
                    break;
                case "Mute":
                    SendKeys.Send("{F1}");
                    break;
                case "Show Downloads":
                    SendKeys.Send("^j");
                    break;
                case "Volume up":
                    SendKeys.Send("^{UP}");
                    break;
                case "Volume down":
                    SendKeys.Send("^{DOWN}");
                    break;
                case "Increase Volume":
                    SendKeys.Send("{F3}");
                    break;
                case "Hide Yourself":
                    this.Opacity = .0;
                    this.WindowState = FormWindowState.Minimized;
                    break;
                case "Show Yourself":
                    this.WindowState = FormWindowState.Minimized;
                    this.WindowState = FormWindowState.Normal;
                    this.Opacity = .99;
                    break;
                case "Define Yourself":
                    VICTOR.Speak("I am a weak Artificial Intelligence developed by my Dad Victor");
                    break;
                case "Page down":
                    SendKeys.Send("{PGDN}");
                    break;
                case "Page up":
                    SendKeys.Send("{PGUP}");
                    break;
                case "Previous song":
                    SendKeys.Send("^b");
                    break;
                case "Bookmarks":
                    SendKeys.Send("^b");
                    break;
                case "Change Voice":
                    VICTOR.SelectVoiceByHints(VoiceGender.Female);
                    break;
                case "hello":
                case "hello Victor":                   
                        VICTOR.Speak("Hi sir,  I am Online");
                         break;
                //case "goodbye":
                case "goodbye Victor":
               case "close":
                case "close Victor":
                    VICTOR.Speak("Until next time");
                    Close();
                    break;
                case "Victor":
                    VICTOR.Speak(num);
                    VICTOR.Speak(battstatus);
                    VICTOR.Speak(battlife + "Percentage");
                    VICTOR.Speak(batt + "seconds");                   
                    VICTOR.Speak("Status is" + sam);
                    break;
                //WEBSITES
               
                case "open facebook":
                    System.Diagnostics.Process.Start("www.facebook.com");                    
                    break;
                case "open google":
                    System.Diagnostics.Process.Start("www.google.co.in");                   
                    break;
                //SHELL COMMANDS
                case "open jarvis":
                    System.Diagnostics.Process.Start("E:/Jarvis.exe");                   
                    break;
                case "open command":
                    System.Diagnostics.Process.Start("C:/Windows/System32/cmd.exe");                    
                    break;
                case "open notepad":
                    System.Diagnostics.Process.Start("C:/Windows/system32/notepad.exe");                    
                    break;
                case "open firefox":
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Mozilla Firefox/firefox.exe");                    
                    break;
                //CLOSE PROGRAMS
                case "close command":
                    ProcWindow = "cmd";                    
                    StopWindow();
                    break;
                case "close notepad":
                    ProcWindow = "notepad";                    
                    StopWindow();
                    break;               
                case "quit jarvis":
                    ProcWindow = "jarvis";
                    VICTOR.Speak("I am killing Jarvis");
                    StopWindow();
                    break;
                case "close firefox":
                case "close facebook":
                    ProcWindow = "firefox";                    
                    StopWindow();
                    break;
                //CONDITION OF DAY
                case "Open":
                case "Ok":                   
                    SendKeys.Send("{ENTER}");
                    break;
                case "what day is it?":
                    VICTOR.Speak(DateTime.Today.ToString("dddd"));
                    break;
                case "whats the date":
                case "whats todays date":
                    VICTOR.Speak
                    (DateTime.Today.ToString("MM-dd-yyyy"));
                    break;
                //OTHER COMMANDS
                case "go fullscreen":
                    FormBorderStyle = FormBorderStyle.None;
                    WindowState = FormWindowState.Maximized;
                    TopMost = true;                   
                    break;
                case "what time is it":
                    DateTime date = DateTime.Now;
                    string time = date.ToString("hh:mm");
                    VICTOR.Speak(time);
                    break;  
                case "exit fullscreen":
                   // VICTOR.Speak("As your Wish Sir");
                    FormBorderStyle = FormBorderStyle.Sizable;
                    WindowState = FormWindowState.Normal;
                    TopMost = false;
                    break;
                case "show commands":
                    string[] commands = (File.ReadAllLines(@"E:\commands.vic"));
                    VICTOR.Speak("Right Away Sir");
                    lstCommands.Items.Clear();
                    lstCommands.SelectionMode = SelectionMode.None;
                    lstCommands.Visible = true;
                    foreach (string command in commands)
                    {
                        lstCommands.Items.Add(command);
                    }
                    break;
                case "hide listbox":
                    lstCommands.Visible = false;
                    break;
                case "show listbox":
                    lstCommands.Visible = true;
                    break;
                //SHUTDOWN RESTART LOG OFF                
                default:
                    break;
            }
        }
        private void StopWindow()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(ProcWindow);
            foreach (System.Diagnostics.Process proc in procs)
            {
                proc.CloseMainWindow();
            }
        }

    }
}
