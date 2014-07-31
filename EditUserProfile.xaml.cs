using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.Protocols;
using System.Linq;
using Microsoft.PowerShell;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Active_Directory_Interface
{
    /// <summary>
    /// Interaction logic for EditUserProfile.xaml
    /// </summary>
    public partial class EditUserProfile : Window
    {
        
        public EditUserProfile(string userName)
        {
            string globalUserame = userName;

            InitializeComponent();

            ///////////////////////////////////////////////////////////////////////////////////////////////
            ////  Sets combo box items for office location and departments.  Code omitted for sample.  ////
            ///////////////////////////////////////////////////////////////////////////////////////////////

            PrincipalContext insPrincipalContext = new PrincipalContext(ContextType.Domain, "[Domain]", "[Active Directory Location]");

            UserPrincipal insUserPrincipal = new UserPrincipal(insPrincipalContext);
            insUserPrincipal.Name = "*";
            GetUsers(insUserPrincipal);

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, userName);

            string userOU = userOUInfo.DistinguishedName.ToString();

            DirectoryEntry entry = new DirectoryEntry("LDAP://" + userOU);

            if (entry.Properties["manager"].Value != null)
            {
                string managerString = entry.Properties["manager"].Value.ToString();
                var managerDN1 = managerString.Split(',');
                var managerDN2 = managerDN1[0].ToString();
                var managerDN3 = managerDN2.Replace("CN=", "");

                managerCB.Items.Insert(0, managerDN3);
                managerCB.SelectedIndex = 0;
            }
            else
            {
                string managerDN3 = "No Manager";

                managerCB.Items.Insert(0, managerDN3);
                managerCB.SelectedIndex = 0;
            }

            ////  Stores employee information before changes.
            ////
            ////  This should be stored as a variable, but this is how I addressed it originally.
            #region Hidden TB
            hiddenFNTB.Text = entry.Properties["givenName"].Value.ToString();
            hiddenLNTB.Text = entry.Properties["sn"].Value.ToString();
            hiddenUNTB.Text = entry.Properties["sAMAccountName"].Value.ToString();

            if (entry.Properties["title"].Value != null)
            {
                hiddenJTTB.Text = entry.Properties["title"].Value.ToString();
            }
            else
            {
                hiddenJTTB.Clear();
            }

            if (entry.Properties["department"].Value != null)
            {
                hiddenDeptTB.Text = entry.Properties["department"].Value.ToString();

                departmentCB.Items.Remove(entry.Properties["department"].Value.ToString());
                departmentCB.Items.Insert(0, entry.Properties["department"].Value.ToString());
                departmentCB.SelectedIndex = 0;
            }
            else
            {
                hiddenDeptTB.Clear();
            }


            if (entry.Properties["manager"].Value != null)
            {
                string managerString = entry.Properties["manager"].Value.ToString();
                var managerDN1 = managerString.Split(',');
                var managerDN2 = managerDN1[0].ToString();
                var managerDN3 = managerDN2.Replace("CN=", "");

                hiddenManTB.Text = managerDN3;
            }
            else
            {
                string managerDN3 = "No Manager";

                hiddenManTB.Text = managerDN3;

            }
            #endregion


            firstNameTB.Text = entry.Properties["givenName"].Value.ToString();
            lastNameTB.Text = entry.Properties["sn"].Value.ToString();
            userNameTB.Text = entry.Properties["sAMAccountName"].Value.ToString();


            if (entry.Properties["title"].Value != null)
            {
                jobTitleTB.Text = entry.Properties["title"].Value.ToString();
            }
            else
            {
            }

            if (entry.Properties["physicalDeliveryOfficeName"].Value != null)
            {
                officeLocationCB.Items.Remove(entry.Properties["physicalDeliveryOfficeName"].Value.ToString());
                officeLocationCB.Items.Insert(0, entry.Properties["physicalDeliveryOfficeName"].Value.ToString());
                officeLocationCB.SelectedIndex = 0;
            }
            else
            {
            }
            
            emailTextBlock.Text = entry.Properties["mail"].Value.ToString();

            if (entry.Properties["description"].Value != null)
            {
                descriptionEditBox.Text = entry.Properties["description"].Value.ToString();
            }

        }





        /// <summary>
        ///   Retrieves employee names from Active Directory for the manager combo box.  Generates on window load.
        /// </summary>
        public void GetUsers(UserPrincipal parUserPrincipal)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            List<string> names = new List<string>();

            managerCB.Items.Clear();
            PrincipalSearcher insPrincipalSearcher = new PrincipalSearcher();
            insPrincipalSearcher.QueryFilter = parUserPrincipal;
            PrincipalSearchResult<Principal> results = insPrincipalSearcher.FindAll();
            foreach (Principal p in results)
            {
                string name = p.ToString();

                UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, p.ToString());
                string userOU = userOUInfo.DistinguishedName.ToString();
                DirectoryEntry checkUser = new DirectoryEntry("LDAP://" + userOU);

                string discardOU = checkUser.Properties["distinguishedName"].Value.ToString();

                if (!discardOU.Contains("SBSUsers"))
                {
                    names.Add(name);
                }

            }

            names.Sort();

            names.ForEach(delegate(String name)
            {
                managerCB.Items.Add(name);
            });
            
            managerCB.Items.Insert(0, "No Manager");

        }






        /// <summary>
        ///  This is the main work done when the user submits changes.
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string userName = hiddenFNTB.Text + " " + hiddenLNTB.Text;

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, userName);

            UserPrincipal managerOUInfo = UserPrincipal.FindByIdentity(ctx, managerCB.SelectedItem.ToString());

            string userOU = userOUInfo.DistinguishedName.ToString();

            DirectoryEntry editUserProfile = new DirectoryEntry("LDAP://" + userOU);

            

            if (firstNameTB.Text != hiddenFNTB.Text)
            {
                editUserProfile.Properties["givenName"].Clear();
                editUserProfile.Properties["givenName"].Value = firstNameTB.Text;
            }
            else
            {
            }

            if (lastNameTB.Text != hiddenLNTB.Text)
            {
                editUserProfile.Properties["sn"].Value = lastNameTB.Text;

                if (hiddenUNTB.Text != firstNameTB.Text.Trim().ToLower())
                {
                    string newUserName = firstNameTB.Text[0] + lastNameTB.Text;
                    string newUserNameLowerCase = newUserName.ToLower();

                    editUserProfile.Properties["sAMAccountName"].Value = newUserNameLowerCase;
                    editUserProfile.Properties["userPrincipalName"].Value = newUserNameLowerCase + "@EnvironmentalWorks.local";
                }
            }
            else
            {
            }


            if (firstNameTB.Text + " " + lastNameTB.Text != hiddenFNTB.Text + " " + hiddenLNTB.Text)
            {
                editUserProfile.Properties["displayName"].Clear();
                editUserProfile.Properties["displayName"].Value = firstNameTB.Text + " " + lastNameTB.Text;
            }
            else
            {
            }

            if (officeLocationCB.SelectedIndex != -1)
            {
                editUserProfile.Properties["physicalDeliveryOfficeName"].Value = officeLocationCB.SelectedItem.ToString();
            }
            else
            {
                editUserProfile.Properties["physicalDeliveryOfficeName"].Clear();
            }


            if (managerCB.SelectedItem.ToString() == "No Manager")
            {
                editUserProfile.Properties["manager"].Clear();
            }
            else
            {
                editUserProfile.Properties["manager"].Value = managerOUInfo.DistinguishedName.ToString();
            }



            if (departmentCB.SelectedIndex != -1 && departmentCB.SelectedItem.ToString() != hiddenDeptTB.Text)
            {
                editUserProfile.Properties["department"].Value = departmentCB.SelectedItem.ToString();
            }
            else
            {
                editUserProfile.Properties["department"].Clear();
            }

            if (jobTitleTB.Text.Trim().Length > 0)
            {
                editUserProfile.Properties["title"].Clear();
                editUserProfile.Properties["title"].Value = jobTitleTB.Text;
            }
            else
            {
                editUserProfile.Properties["title"].Clear();
            }

            if (descriptionEditBox.Text.Trim().Length > 0)
            {
                editUserProfile.Properties["description"].Value = descriptionEditBox.Text;
            }
            else
            {
                editUserProfile.Properties["description"].Clear();
            }


            try
            {
                editUserProfile.CommitChanges();
            }
            catch (Exception f)
            {
                System.Windows.Forms.MessageBox.Show("Failed to save changes   " + f);
                return;
            }


            editUserProfile.Rename("CN=" + firstNameTB.Text.Trim() + " " + lastNameTB.Text.Trim());


            nameChange();



            if (hiddenUNTB.Text != firstNameTB.Text)
            {
                appendMailBox();
            }

            

            if (officeLocationCB.SelectedIndex != -1 && officeLocationCB.SelectedItem.ToString() != hiddenOLTB.Text)
            {
                string newName = firstNameTB.Text + " " + lastNameTB.Text;

                PrincipalContext ctx2 = new PrincipalContext(ContextType.Domain);

                UserPrincipal userOUInfo2 = UserPrincipal.FindByIdentity(ctx, newName);

                string userOU2 = userOUInfo2.DistinguishedName.ToString();

                if (officeLocationCB.SelectedItem.ToString() == "[Office Location]" && departmentCB.SelectedItem.ToString() == "[Department]")
                {
                    DirectoryEntry currentLocation = new DirectoryEntry("LDAP://" + userOU2);
                    DirectoryEntry newLocation = new DirectoryEntry("[Active Directory Location]");
                    currentLocation.MoveTo(newLocation);

                    ///////////////////////////////////////////////////////////////////////////////////////////////////
                    ////  Change group memberships based on office location, if changed.  Code omitted in sample.  ////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////
                }
            }
            else
            {
                return;
            }

            System.Windows.Forms.MessageBox.Show("Change Successful");

            this.Hide();
        }



        private void nameChange()
        {

            string lowerCaseName = firstNameTB.Text.Trim().ToLower();

            if (lowerCaseName != hiddenUNTB.Text)
            {
                string currentFileName = hiddenFNTB.Text[0] + hiddenLNTB.Text;
                string currentFileNameLowerCase = currentFileName.ToLower();
                string newFileName = firstNameTB.Text[0] + lastNameTB.Text;
                string newFileNameLowerCase = newFileName.ToLower();

                if (newFileNameLowerCase != currentFileNameLowerCase)
                {
                    System.IO.Directory.Move("[Home Directory]" + currentFileNameLowerCase, "[Home Directory]" + newFileNameLowerCase);
                }
            }
            else
            {
            }
        }





        /// <summary>
        ///   Make changes to Exchange mailbox.
        /// </summary>
        private void appendMailBox()
        {
            // Change and add additional e-mail address
            SecureString password = new SecureString();
            string str_password = "[Admin Password]";
            string username = "[Username]";

            string liveIdconnectionUri = "[Powershell uri]";

            foreach (char x in str_password)
            {
                password.AppendChar(x);
            }

            PSCredential credential = new PSCredential(username, password);

            // Create connection info
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(liveIdconnectionUri)), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);

            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default;

            // Create a runspace on a remote path
            Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);

            // Create Exhange Server Mailbox
            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();

            string lowerCaseName = firstNameTB.Text.Trim().ToLower()[0] + lastNameTB.Text.Trim().ToLower();

            
            command.AddScript("set-mailbox " + hiddenUNTB.Text + " -alias \"" + lowerCaseName + "\"");


            powershell.Commands = command;
            try
            {
                // open the remote runspace
                runspace.Open();
                // associate the runspace with powershell
                powershell.Runspace = runspace;
                // invoke the powershell to obtain the results
                powershell.Invoke();
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            finally
            {


                // dispose the runspace and enable garbage collection
                runspace.Dispose();
                runspace = null;


                // Finally dispose the powershell and set all variables to null to free
                // up any resources.
                powershell.Dispose();
                powershell = null;
            }
        }
    }
}
