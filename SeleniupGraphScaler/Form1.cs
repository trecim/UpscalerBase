using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SeleniupGraphScaler
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            try
            {
                Directory.Delete("Cache",true);
            }
            catch { }
            try
            {
                Directory.CreateDirectory("Cache");
            }
            catch{ Application.Exit(); }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FirefoxDriver driver = null;
            if (iszyroScaler.Checked)
            {
                var profile = new FirefoxProfile();
                profile.DeleteAfterUse = true;
                profile.Clean();
                profile.SetPreference("browser.download.folderList", 2);
                //profile.SetPreference("browser.download.manager.showWhenStarting", false);
                profile.SetPreference("browser.download.dir", Application.StartupPath + "\\Cache");
                profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "image/gif, image/pjpeg, image/jpeg, image/jpg, image/png, image/tiff");
                profile.SetPreference("general.useragent.override", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.111 Safari/537.36" + new Random().Next().GetHashCode().ToString());

                FirefoxOptions options = new FirefoxOptions();
                //options.AddArgument("user-agent=" + new Random().Next().GetHashCode().ToString() + "");
                options.Profile = profile;

                driver = new FirefoxDriver("C:\\Program Files\\Mozilla Firefox", options);
            }
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                List<string> ls = GetRecursFiles(folderBrowserDialog1.SelectedPath);
                foreach (string fname in ls)
                {
                    if (Path.GetExtension(fname) == ".png" || Path.GetExtension(fname) == ".jpg")
                    {
                        if (iszyroScaler.Checked)
                            ZyroUpscaleImage(fname, driver);
                        else
                            UpscaleImage(fname);
                    }
                }
            }
        }

        private void UpscaleImage(string path)
        {
            try
            {
                Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Mozilla", true);
                
                
            }
            catch { }
            try
            {
                Directory.Delete(Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) + "\\Local\\Mozilla", true);
            }
            catch { }
            try
            {
                Directory.Delete(Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) + "\\Local\\Temp", true);
                Directory.CreateDirectory(Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) + "\\Local\\Temp");
            }
            catch { }
            var tempfiles = GetRecursFiles((Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) + "\\Local\\Temp"));
            foreach (string filename in tempfiles)
            {
                try
                {
                    File.Delete(filename);
                }
                catch { }
            }

            var profile = new FirefoxProfile();
            profile.DeleteAfterUse = true;
            profile.Clean();
            profile.SetPreference("browser.download.folderList", 2);
            //profile.SetPreference("browser.download.manager.showWhenStarting", false);
            profile.SetPreference("browser.download.dir", Application.StartupPath + "\\Cache");
            profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "image/gif, image/pjpeg, image/jpeg, image/jpg, image/png, image/tiff");
            profile.SetPreference("general.useragent.override", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.111 Safari/537.36" + new Random().Next().GetHashCode().ToString());

            FirefoxOptions options = new FirefoxOptions();
            //options.AddArgument("user-agent=" + new Random().Next().GetHashCode().ToString() + "");
            options.Profile = profile;

            var driver = new FirefoxDriver("C:\\Program Files\\Mozilla Firefox", options);
            again:
            try
            {
                driver.Url = "https://icons8.com/upscaler";
                Thread.Sleep(2000);
                IWebElement elem = driver.FindElement(By.XPath("//input[@type='file']"));

                elem.SendKeys(path);

                bool stop = false;
                int count = 0;
                while (!stop)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//*[text() = 'Download']")).Click();
                        stop = true;
                    }
                    catch
                    {
                        if (count == 20)
                        {
                            IWebElement elem2 = driver.FindElement(By.XPath("//input[@type='file']"));
                            elem2.SendKeys(path);
                        }
                        count++;
                        Thread.Sleep(500);
                    }
                }
                while (!File.Exists(Application.StartupPath + "\\Cache\\" + Path.GetFileName(path)))
                    Thread.Sleep(500);
                driver.Close();
                File.Move(path, path + ".back");
                File.Move(Application.StartupPath + "\\Cache\\" + Path.GetFileName(path), path);
                Process[] runingProcess = Process.GetProcesses();
                for (int i = 0; i < runingProcess.Length; i++)
                {
                    // compare equivalent process by their name
                    if (runingProcess[i].ProcessName == "geckodriver")
                    {
                        // kill  running process
                        runingProcess[i].Kill();
                    }

                }
            }
            catch
            {
                Thread.Sleep(2000);
                goto again;
            }
            //driver.Url = "https://google.com";
            
        }

        private void ZyroUpscaleImage(string path, FirefoxDriver driver)
        {
            driver.Url = "https://zyro.com/tools/image-upscaler";
            Thread.Sleep(2000);
            IWebElement elem = driver.FindElement(By.XPath("//input[@type='file']"));

            elem.SendKeys(path);

            bool stop = false;
            bool skip = false;
            int count = 0;
            int count2 = 0;
            while (!stop)
            {
                try
                {
                    driver.FindElement(By.XPath("//button[@data-qa='image-upscaler-download-button']")).Click();
                    stop = true;
                }
                catch
                {
                    if (count == 10)
                    {
                        
                        IWebElement elem2 = driver.FindElement(By.XPath("//input[@type='file']"));
                        elem2.SendKeys(path);
                        count = 0;
                        count2++;
                    }
                    if(count2 > 2)
                    {
                        stop = true;
                        skip = true;
                    }
                    count++;
                    
                    Win32.POINT p = new Win32.POINT(count * (Screen.PrimaryScreen.Bounds.Width /10) , Screen.PrimaryScreen.Bounds.Height / 2);
                    Win32.ClientToScreen(this.Handle, ref p);
                    Win32.SetCursorPos(p.x, p.y);
                    Thread.Sleep(1000);
                }
            }
            if(!skip)
            {
                while (!File.Exists(Application.StartupPath + "\\Cache\\" + Path.GetFileName("zyro-image" + Path.GetExtension(path))))
                    Thread.Sleep(500);
                File.Move(path, path + ".back");
                File.Move(Application.StartupPath + "\\Cache\\" + Path.GetFileName("zyro-image" + Path.GetExtension(path)), path);
            }
            
        }

        private bool IsLocked(string filePath)
        {

            FileInfo f = new FileInfo(filePath);
            FileStream stream = null;

            try
            {
                stream = f.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException ex)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }


        private List<string> GetRecursFiles(string start_path)
        {
            List<string> ls = new List<string>();
            try
            {
                string[] folders = Directory.GetDirectories(start_path);
                foreach (string folder in folders)
                {
                    ls.Add("Папка: " + folder);
                    ls.AddRange(GetRecursFiles(folder));
                }
                string[] files = Directory.GetFiles(start_path);
                foreach (string filename in files)
                {
                    ls.Add(filename);
                }
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return ls;
        }
    }
    public class Win32
    {
        [DllImport("User32.Dll")]
        public static extern long SetCursorPos(int x, int y);

        [DllImport("User32.Dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT point);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;

            public POINT(int X, int Y)
            {
                x = X;
                y = Y;
            }
        }
    }
}
