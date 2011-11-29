using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security.AccessControl;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Xml;
using System.IO;
using System.Xml.Linq;
using Microsoft.Win32;


namespace Mac_Changer
{
    public partial class Form1 : Form
    {
        private XmlDocument _doc;
        private const string Path = @"MAC.xml";
        public Form1()
        {
            InitializeComponent();
            GetInterfaces();
            GetMac();
        }
        public class Language
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public override string ToString() { return Name; }
        }
        public void GetMac()
        {
            XDocument xmlDoc = XDocument.Load(Path);
            var q = from c in xmlDoc.Descendants("Mac") let xElement = c.Element("Value") where xElement != null let element = c.Element("Id") where element != null select new { Value = xElement.Value, Id = element.Value };
            var dict = q.ToDictionary(obj => obj.Value, obj => Convert.ToInt32(obj.Id));
            comboBox2.DataSource = new BindingSource(dict, null);
            comboBox2.DisplayMember = "Key";
            comboBox2.ValueMember = "Value";
            comboBox2.Focus();
        }
        public void GetInterfaces()
        {
            var dict = new Dictionary<string, string>();
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet || nic.NetworkInterfaceType == NetworkInterfaceType.FastEthernetFx || nic.NetworkInterfaceType == NetworkInterfaceType.FastEthernetT
)
                {
                    PhysicalAddress physicalAddress = nic.GetPhysicalAddress();
                    dict.Add(nic.Description, physicalAddress.ToString());
                }
            }
            comboBox1.DataSource = new BindingSource(dict, null);
            comboBox1.DisplayMember = "Key";
            comboBox1.ValueMember = "Value";
            comboBox1.Focus();
            label4.Text = comboBox1.SelectedValue.ToString();
        }
        public void SetMac(string macAddress)
        {
            const string Name = @"SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002bE10318}";
            using (RegistryKey key0 = Registry.LocalMachine.OpenSubKey(Name, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl))
            {

                string[] x = key0.GetSubKeyNames();
                foreach (string name in x)
                {

                        var var1 = Registry.LocalMachine.OpenSubKey(Name,RegistryKeyPermissionCheck.ReadWriteSubTree,RegistryRights.FullControl);
                        var v = var1.OpenSubKey(name, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
                        var z = v.GetValue("DriverDesc");
                        if (comboBox1.Text == z.ToString() )
                        {
                            v.SetValue("NetworkAddress",comboBox2.Text);
                            MessageBox.Show(z.ToString());
                        }
                        v.Close();
                        var1.Close();

                }
                key0.Close();
            }

        }
        public void Reset()
        {
            ManagementObjectSearcher mos = null;
            mos = new ManagementObjectSearcher(@"SELECT * 
                                     FROM   Win32_NetworkAdapter 
                                     WHERE  Description = '"+comboBox1.Text+"'");
            ManagementObjectCollection v = mos.Get();

            foreach (ManagementObject mo in v)
            {
                string[] args= {};
                mo.InvokeMethod("Disable",args);
                mo.InvokeMethod("Enable", args);
            }


        }

        private void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
        {
            label4.Text = comboBox1.SelectedValue.ToString();
        }

        public void Form1Load(object sender, EventArgs e)
        {

            //Create an xml document
            _doc = new XmlDocument();

            //If there is no current file, then create a new one
            if (!File.Exists(Path))
            {
                //Create neccessary nodes
                XmlDeclaration declaration = _doc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                XmlComment comment = _doc.CreateComment("This is an XML Generated File");
                XmlElement root = _doc.CreateElement("Root");
                _doc.AppendChild(declaration);
                _doc.AppendChild(comment);
                _doc.AppendChild(root);
                _doc.Save(Path);
            }
        }

        private void Button1Click(object sender, EventArgs e)
        {

            const string pattern = @"^([0-9a-fA-F][0-9a-fA-F]){5}([0-9a-fA-F][0-9a-fA-F])$";

            if (Regex.IsMatch((textBox1.Text), pattern))
            {
                //Load the XML File
                _doc.Load(Path);

                //Get the root element
                XmlElement root = _doc.DocumentElement;

                //XmlElement person = doc.CreateElement("Person");
                XmlElement mac = _doc.CreateElement("Mac");
                XmlElement value = _doc.CreateElement("Value");
                XmlElement id = _doc.CreateElement("Id");

                //Add the values for each nodes
                value.InnerText = PhysicalAddress.Parse(textBox1.Text.ToUpper()).ToString();
                var rnd = new Random();
                id.InnerText = rnd.Next(0, int.MaxValue).ToString();

                //Construct the value element
                mac.AppendChild(value);
                mac.AppendChild(id);

                //Add the New mac element to the end of the root element
                root.AppendChild(mac);

                //Save the document
                _doc.Save(Path);
                MessageBox.Show("Added Succesfully", "Mac Changer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Invalid Input Format\nEnter Address without spaces\nExample:  ff00ff00ff00", "MAC Changer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            GetMac();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetMac(comboBox2.SelectedValue.ToString());
            Reset();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Reset();
        }

    }
}
