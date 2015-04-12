using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;
using System.Xml.Linq;
using System.Xml;
using System.IO;

using System.Net.Mail;

namespace JARVISV2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer JARVIS = new SpeechSynthesizer();
        String Temperature;
        String Condition;
        String Humidity;
        String WinSpeed;
        String TFCond;
        String TFHigh;
        String TFLow;
        String WOEID = "1528488"; //<<<<-------- CHANGE THE WOEID CODE HERE
        String Town;
        String QEvent;
        DateTime now = DateTime.Now;
        String userName = Environment.UserName;
        Random rnd = new Random();
        int timer = 11;
        int count = 1;

        public Form1()
        {
            InitializeComponent();
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Commands.txt")))));
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
        }

        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string time = now.GetDateTimeFormats('t')[0];
            int ranNum;
            string speech = e.Result.Text;
            switch (speech)
            {

                //GREETINGS
                case "Hello":
                case "Hello Jarvis":
                    if (now.Hour >= 5 && now.Hour < 12)
                    { JARVIS.Speak("Goodmorning " + userName); }
                    if (now.Hour >= 12 && now.Hour < 18)
                    { JARVIS.Speak("Good afternoon " + userName); }
                    if (now.Hour >= 18 && now.Hour < 24)
                    { JARVIS.Speak("Good evening " + userName); }
                    if (now.Hour < 5)
                    { JARVIS.Speak("Hello, it is getting late " + userName); }
                    break;
                case "Goodbye":
                case "Goodbye Jarvis":
                case "Close":
                case "Close Jarvis":
                    JARVIS.Speak("Farewell");
                    Close();
                    break;
                case "Jarvis":
                    ranNum = rnd.Next(1, 3);
                    if (ranNum == 1) { QEvent = ""; JARVIS.Speak("Yes sir"); }
                    else if (ranNum == 2) { QEvent = ""; JARVIS.Speak("Yes?"); }
                    break;

                //CONDITION OF DAY
                case "What time is it":
                    JARVIS.Speak(time);
                    break;
                case "What day is it":
                    JARVIS.Speak(DateTime.Today.ToString("dddd"));
                    break;
                case "Whats the date":
                case "Whats todays date":
                    JARVIS.Speak(DateTime.Today.ToString("dd-MM-yyyy"));
                    break;
                case "Hows the weather":
                case "Whats the weather like":
                case "Whats it like outside":
                    GetWeather();
                    if (QEvent == "connected")
                    { JARVIS.Speak("The weather in " + Town + " is " + Condition + " at " + Temperature + " degrees. There is a humidity of " + Humidity + " and a windspeed of " + WinSpeed + " miles per hour"); }
                    else if (QEvent == "failed")
                    { JARVIS.Speak("I seem to be having a bit of trouble connecting to the server. Just look out the window"); }
                    break;
                case "What will tomorrow be like":
                case "Whats tomorrows forecast":
                case "Whats tomorrow like":
                    GetWeather();
                    if (QEvent == "connected")
                    { JARVIS.Speak("Tomorrows forecast is " + TFCond + " with a high of " + TFHigh + " and a low of " + TFLow); }
                    else if (QEvent == "failed")
                    { JARVIS.Speak("It's hard to tell without an stable internet connection"); }
                    break;
                case "Whats the temperature outside":
                    GetWeather();
                    if (QEvent == "connected")
                    { JARVIS.Speak(Temperature + " degrees"); }
                    else if (QEvent == "failed")
                    { JARVIS.Speak("I cannot access the server at this time"); }
                    break;

                //OTHER COMMANDS
                case "Switch Window":
                    SendKeys.Send("%{TAB " + count + "}");
                    count += 1;
                    break;
                case "Reset":
                    count = 1;
                    int timer = 11;
                    lblTimer.Visible = false;
                    ShutdownTimer.Enabled = false;
                    lstCommands.Visible = false;
                    break;
                case "Out of the way":
                    if (WindowState == FormWindowState.Normal)
                    {
                        WindowState = FormWindowState.Minimized;
                        JARVIS.Speak("My apologies");
                    }
                    break;
                case "Come back":
                    if (WindowState == FormWindowState.Minimized)
                    {
                        JARVIS.Speak("Alright?");
                        WindowState = FormWindowState.Normal;
                    }
                    break;
                case "Show commands":
                    string [] commands = File.ReadAllLines("Commands.txt");
                    JARVIS.Speak("Here we are");
                    lstCommands.Items.Clear();
                    lstCommands.SelectionMode = SelectionMode.None;
                    lstCommands.Visible = true;
                    foreach (string command in commands)
                    {
                        lstCommands.Items.Add(command);
                    }
                    break;
                case "Hide listbox":
                    lstCommands.Visible = false;
                    break;

                //SHUTDOWN RESTART LOG OFF
                case "Shutdown":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "shutdown";
                        JARVIS.Speak("I will shutdown shortly");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Log off":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "logoff";
                        JARVIS.Speak("Logging off");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Restart":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "restart";
                        JARVIS.Speak("I'll be back shortly");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Abort":
                    if (ShutdownTimer.Enabled == true)
                    {
                        timer = 11;
                        lblTimer.Text = timer.ToString();
                        ShutdownTimer.Enabled = false;
                        lblTimer.Visible = false;
                    }
                    break;
            }
        }
        private void GetWeather()
        {
            string query = String.Format("http://weather.yahooapis.com/forecastrss?w=" + WOEID);
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager man = new XmlNamespaceManager(wData.NameTable);
            man.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("rss").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("/rss/channel/item/yweather:forecast", man);

            Temperature = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", man).Attributes["temp"].Value;

            Condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", man).Attributes["text"].Value;
            
            Humidity = channel.SelectSingleNode("yweather:atmosphere", man).Attributes["humidity"].Value;

            WinSpeed = channel.SelectSingleNode("yweather:wind", man).Attributes["speed"].Value;

            Town = channel.SelectSingleNode("yweather:location", man).Attributes["city"].Value;

            TFCond = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["text"].Value;

            TFHigh = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["high"].Value;

            TFLow = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["low"].Value;
        }

        private void ShutdownTimer_Tick(object sender, EventArgs e)
        {
            if (timer == 0)
            {
                lblTimer.Visible = false;
                ComputerTermination();
                ShutdownTimer.Enabled = false;
            }
            else
            {
                timer = timer - 1;
                lblTimer.Text = timer.ToString();
            }
        }
        private void ComputerTermination()
        {
            switch (QEvent)
            {
                case "shutdown":
                    System.Diagnostics.Process.Start("shutdown", "-s");
                    break;
                case "logoff":
                    System.Diagnostics.Process.Start("shutdown", "-l");
                    break;
                case "restart":
                    System.Diagnostics.Process.Start("shutdown", "-r");
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}