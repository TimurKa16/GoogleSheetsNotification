using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;

namespace Uninstaller
{
    public partial class UninstallerForm : Form
    {
        public UninstallerForm()
        {
            InitializeComponent();
        }

        private void CloseButtonClick(object sender, EventArgs e)
        {
            CloseProcesses();
        }

        private void DeleteButtonClick(object sender, EventArgs e)
        {
            CloseProcesses();
            SetAutorunValue(false);
        }



        void CloseProcesses()
        {
            foreach (Process process in Process.GetProcesses())
            {
                string name = process.ProcessName;
                if (name == "NotificationAnvar" || name == "NotificationMarat" || name == "Notification")
                    process.Kill();
            }
        }

        const string applicationName1 = "NotificationAnvar";
        string exePath1 = "C:\\Program Files (x86)\\Timur Corporation\\TimurNotification\\NotificationAnvar.exe";

        const string applicationName2 = "NotificationMarat";
        string exePath2 = "C:\\Program Files (x86)\\Timur Corporation\\TimurNotification\\NotificationMarat.exe";

        const string applicationName3 = "Notification";
        string exePath3 = "C:\\Program Files (x86)\\Timur Corporation\\TimurNotification\\Notification.exe";
        public bool SetAutorunValue(bool autorun)
        {
            string ExePath = Application.ExecutablePath;
            RegistryKey reg;

            reg = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run\\");
            try
            {
                if (autorun)
                {
                    try
                    {
                        reg.SetValue(applicationName1, exePath1);
                    }
                    catch (Exception) { };
                    try
                    {
                        reg.SetValue(applicationName2, exePath2);
                    }
                    catch (Exception) { };
                    try
                    {
                        reg.SetValue(applicationName3, exePath3);
                    }
                    catch (Exception) { };
                }
                else
                {
                    try
                    {
                        reg.DeleteValue(applicationName1);
                    }
                    catch (Exception) { };
                    try
                    {
                        reg.DeleteValue(applicationName2);
                    }
                    catch (Exception) { };
                    try
                    {
                        reg.DeleteValue(applicationName3);
                    }
                    catch (Exception) { };
                }

                reg.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }

        private void UninstallerForm_Load(object sender, EventArgs e)
        {

        }
    }
}
