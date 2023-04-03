using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.ServiceProcess;
using System.IO;

namespace Printer_remover
{
    public partial class Form1 : Form
    {
        List<string> list = new List<string>();
        public Form1()
        {
            InitializeComponent();
            NactiRegistry();
        }

        private void NactiRegistry()
        {
            string verze = string.Empty;
            listView1.Items.Clear();

            if(Environment.Is64BitOperatingSystem)
            {
                verze = "Windows x64";
            }
            else
            {
                verze = "Windows NT x86";   
            }

            RegistryKey rk = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\" + verze + "\\Drivers\\Version-3");

            if(rk != null)
                foreach (string s in rk.GetSubKeyNames())
                {
                    //list klíčů
                    ListViewItem lvi = new ListViewItem(s);
                    lvi.Tag = "SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\" + verze + "\\Drivers\\Version-3";
                    listView1.Items.Add(lvi);
                }


            //načtení version4 v případě win 10
            RegistryKey osname = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion");
            string jmeno = (string)osname.GetValue("ProductName");


            if ((Environment.OSVersion.Version.Major >= 6 && Environment.OSVersion.Version.Minor >= 3) || jmeno.StartsWith("Windows 10"))
            {
                RegistryKey rk10 = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\" + verze + "\\Drivers\\Version-4");

                if(rk10 != null)
                    foreach (string s in rk10.GetSubKeyNames())
                    {
                        //list klíčů
                        ListViewItem lvi = new ListViewItem(s);
                        lvi.Tag = "SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\" + verze + "\\Drivers\\Version-4";
                        listView1.Items.Add(lvi);
                    }
                rk10.Close();
            }
            listView1.ListViewItemSorter = new ListViewItemComparer(0);
            listView1.Sort();
            rk.Close();
            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedIndices.Count > 0)
            {
                list.Clear();
                button1.Enabled = true;
                if (listView1.SelectedIndices.Count == 1)
                {
                    jmeno.Text = listView1.Items[listView1.SelectedIndices[0]].Text;
                    verze.Text = ((string)listView1.Items[listView1.SelectedIndices[0]].Tag).Remove(0, ((string)listView1.Items[listView1.SelectedIndices[0]].Tag).LastIndexOf('\\') + 1);
                }
                else
                {
                    jmeno.Text = "Multiple printers selected";
                    verze.Text = "Multiple printers selected";
                }

                int count = 0;

                foreach (int i in listView1.SelectedIndices)
                {
                    RegistryKey print = Registry.LocalMachine.OpenSubKey((string)(listView1.Items[i].Tag) + "\\" + listView1.Items[i].Text);
                    list.Add((string)(listView1.Items[i].Tag) + "\\" + listView1.Items[i].Text);
                    //var s = print.GetValue("Driver");

                    short p = 0;
                    if ((string)print.GetValue("Driver") != string.Empty)
                        p += 1;
                    if ((string)print.GetValue("Data File") != string.Empty)
                        p += 1;
                    if ((string)print.GetValue("Configuration File") != string.Empty)
                        p += 1;
                    count += (((string[])print.GetValue("Dependent Files")).Length + p);

                }
                soubory.Text = count.ToString();
            }
            else
                button1.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ServiceController service = new ServiceController("Spooler");

            if (service.Status.Equals(ServiceControllerStatus.Running))
            {
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped);
            }

            foreach(string s in list)
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey(s);
                string driver, iniFile, configFile, infPath;
                string[] driverFiles = (string[])key.GetValue("Dependent Files");
                driver = (string)key.GetValue("Driver");
                iniFile = (string)key.GetValue("Data File");
                configFile = (string)key.GetValue("Configuration File");
                infPath = (string)key.GetValue("InfPath");

                if(s.Contains("Version-4"))
                {
                    try
                    {
                        Directory.Delete(((string)key.GetValue("InfPath")).Remove(s.LastIndexOf('\\') + 1), true);
                    }
                    catch { /*MessageBox.Show("Unable to remove " + ((string)key.GetValue("InfPath")).Remove(s.LastIndexOf('\\') + 1) + ". Please check access rights");*/}
                }
                else
                {
                    try
                    {
                        if (Environment.Is64BitOperatingSystem)
                        {
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\x64\\3\\" + driver);
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\x64\\3\\" + iniFile);
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\x64\\3\\" + configFile);
                            foreach (string ss in driverFiles)
                            {
                                File.Delete("C:\\Windows\\System32\\spool\\drivers\\x64\\3\\" + ss);
                            }
                        }
                        else
                        {
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\W32X86\\3\\" + driver);
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\W32X86\\3\\" + iniFile);
                            File.Delete("C:\\Windows\\System32\\spool\\drivers\\W32X86\\3\\" + configFile);
                            foreach (string ss in driverFiles)
                            {
                                File.Delete("C:\\Windows\\System32\\spool\\drivers\\W32X86\\3\\" + ss);
                            }
                        }
                    }
                    catch(IOException ioe)
                    {
                        MessageBox.Show(ioe.Message + "\r\nPlease check if driver file is not used by some proces or service.");
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\r\nPlease check if driver file is not used by some proces or service.");
                    }

                    try
                    {
                        Directory.Delete(((string)key.GetValue("InfPath")).Remove(s.LastIndexOf('\\') + 1), true);
                    }
                    catch { /*MessageBox.Show("Unable to remove " + ((string)key.GetValue("InfPath")).Remove(s.LastIndexOf('\\') + 1) + ". Please check access rights");*/ }
                    
                }

                key.Close();
                RegistryKey del = Registry.LocalMachine.OpenSubKey(s.Remove(s.LastIndexOf('\\')), true);
                del.DeleteSubKeyTree(s.Remove(0, s.LastIndexOf('\\') + 1));
            }


            service.Start();
            service.WaitForStatus(ServiceControllerStatus.Running);
            NactiRegistry();
            button1.Enabled = false;
        }
    }

    class ListViewItemComparer : System.Collections.IComparer
    {
        private int col = 0;

        public ListViewItemComparer(int column)
        {
            col = column;
        }
        public int Compare(object x, object y)
        {
            return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
        }
    }
}
