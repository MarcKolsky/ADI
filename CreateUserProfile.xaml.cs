using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.PowerShell;


namespace Active_Directory_Interface
{
    /// <summary>
    /// Interaction logic for CreateUserProfile.xaml
    /// </summary>
    public partial class CreateUserProfile : Window
    {

        

        public CreateUserProfile()
        {
            PrincipalContext insPrincipalContext = new PrincipalContext(ContextType.Domain, "EnvWorks", "OU=Users, OU=EnvWorks, DC=EnvironmentalWorks, DC=local");   // This is the main connection that points to User's folder in Active Directory.  This connection only works within this method.

            InitializeComponent();
            
            ////////////////////////////////////
            ////  Code Omitted for sample.  ////
            ////////////////////////////////////

            // This initializes the process that grabs all user's from active directory for use in the Manager's combobox.
            UserPrincipal insUserPrincipal = new UserPrincipal(insPrincipalContext);
            insUserPrincipal.Name = "*";
            SearchUsers(insUserPrincipal);

            AddUserSubmitButton.Click += AddUserSubmitButton_Click;  // When the user clicks the submit button, it initializes the process of adding the new employee to Active Directory.

        }




        /// <summary>
        /// This method is linked to the "Check Username" button, and prevents duplicate usernames in Active Directory.
        /// </summary>
        private void UserNameCheck(object sender, RoutedEventArgs e)
        {
            PrincipalContext principalUserNameCheck = new PrincipalContext(ContextType.Domain, "[Domain]", "[AD Path]");     //// Initializes connection to Active Directory.

            var userLogonNameText = firstNameTB.Text.Trim().ToLower()[0] + lastNameTB.Text.Trim().ToLower();          ////  Transforms user's name to the appropriate format.

            ////  Checks to see if the username for the new employee already exists in Active Directory.  If the username is already being used by another employee, it will prompt for an alternate username to be entered.
            UserPrincipal checkUserName = UserPrincipal.FindByIdentity(principalUserNameCheck, userLogonNameText);
            if (checkUserName != null)
            {
                userNameTB.Clear();
                userNameTB.Background = new SolidColorBrush(Colors.Red);
                System.Windows.Forms.MessageBox.Show(userLogonNameText + " already exists. Please enter alternate Username.");
            }
            else
            {
                userNameTB.Text = userLogonNameText;              ////  This username passes the checks and is displayed in the username textbox.
            }
        }




        /// <summary>
        ///  This method controls the the office combobox.  This method is initialized when a selection is made in the "Department" combobox.
        /// </summary>
        private void ComboBoxSelectedItemException(object sender, SelectionChangedEventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ////  When a department is selected, this method filters out offices where that department does not exist.  Code Omitted  ////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }





        /// <summary>
        ///  Grabs all active users from Active Directory.
        /// </summary>
        public void SearchUsers(UserPrincipal parUserPrincipal)
        {
            ////  Initializes connection to Active Directory.
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            List<string> names = new List<string>();        ////  Empty string to store all employee names from Active Directory


            ////  Establishes parameters for array search
            ManagerComboBoxDisplay.Items.Clear();
            PrincipalSearcher insPrincipalSearcher = new PrincipalSearcher();
            insPrincipalSearcher.QueryFilter = parUserPrincipal;
            PrincipalSearchResult<Principal> results = insPrincipalSearcher.FindAll();

            ////  Cycles through Active Directory to find all users.
            foreach (Principal p in results)
            {
                string name = p.ToString();     ////  Sets variable as a string

                UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, p.ToString());
                string userOU = userOUInfo.DistinguishedName.ToString();
                DirectoryEntry checkUser = new DirectoryEntry("LDAP://" + userOU);

                string discardOU = checkUser.Properties["distinguishedName"].Value.ToString();      ////  Parses distinguished name, and sets it as a string

                if (!discardOU.Contains("SBSUsers"))    //// Filters out any server profiles    
                {
                    names.Add(name);    ////  Adds name to the users list created at the beginning of this method.
                }

            }

            ////  Sorts all names in the list alphabetically
            names.Sort();


            ////  Cycles through the list and adds the names to the combobox in order.
            names.ForEach(delegate(String name)
            {
                ManagerComboBoxDisplay.Items.Add(name);
            });


            ManagerComboBoxDisplay.Items.Insert(0, "No Manager");       ////  Sets "No Manager" as the default display in the combobox

        }






        #region Add User

        /// <summary>
        ///   Initiates the process of adding the new employee to Active Directory
        /// </summary>
        internal void AddUserSubmitButton_Click(object sender, RoutedEventArgs e)
        {

            //// Verify all required fields have been filled out.
            if (firstNameTB.Text.Trim().Length != 0 && lastNameTB.Text.Trim().Length != 0 && jobTitleTB.Text.Trim().Length != 0 && DepartmentComboBox.SelectedItem != null && ManagerComboBoxDisplay.SelectedItem != null && OfficeComboBox.SelectedItem != null)
            {

                ////////////////////////////////////////////////////////////////////////
                ////  Set all textboxes and combo boxes to strings.  Code Omitted.  ////
                ////////////////////////////////////////////////////////////////////////


                ////  Formats message to be displayed once the user has been added to Active Directory
                string successMessage = "You have successfully added " + fullName + " to the system. \n\nDepartment: " + departmentName + "\n\nOffice: " + officeLocation + "\n\nManager: " + managerName;  // this is the message displayed when the active directory profile has been successfully created.



                /////////////////////////////////////////////////
                ///////  This is where the action begins  ///////
                /////////////////////////////////////////////////

                string addUserToADFolder = ActiveDirectoryUserLocation();       ////  Initializes ActiveDirectoryUserLocation() method, and sets distinguished Active Directory location of user to a string.


                PrincipalContext addUserConnection = new PrincipalContext(ContextType.Domain, "[Domain]", addUserToADFolder);   //// Initializes connection to the Active Directory

                ////  Secondary check to make sure user doesn't exist in Active Directory
                UserPrincipal usr = UserPrincipal.FindByIdentity(addUserConnection, fullName);    //// Output should be null
                if (usr != null)
                {
                    System.Windows.Forms.MessageBox.Show(fullName + " already exists. Please check employees name.");
                }


                ////  Second check to make sure the the username textbox was filled in,
                ////    and the username doesn't exist in Active Directory
                UserPrincipal secondUserNameCheck = UserPrincipal.FindByIdentity(addUserConnection, userLogonName);     ////  Output should be null
                if (secondUserNameCheck != null || userLogonName == "Validate User  --->")
                {
                    userNameTB.Clear();
                    userNameTB.Background = new SolidColorBrush(Colors.Red);
                    System.Windows.Forms.MessageBox.Show(userLogonName + " already exists. Please enter in alternate Username.");
                    return;
                }


                ////  Creates the instance for the creation of the new user profile.
                UserPrincipal addUserToAD = new UserPrincipal(addUserConnection);



                //////////////////////////////////////
                ////  Populate the user's profile  ///
                //////////////////////////////////////


                ////  Set the username
                if (userNameTB.Text == userLogonName)
                {
                    addUserToAD.SamAccountName = userLogonName;
                    addUserToAD.UserPrincipalName = userLogonName + "[@Domain address]";
                }
                else
                {
                    addUserToAD.SamAccountName = userNameTB.Text;
                    addUserToAD.UserPrincipalName = userNameTB.Text + "[@Domain address]";
                }


                addUserToAD.GivenName = firstNameTB.Text;       ////  Set user's first name

                addUserToAD.Surname = lastNameTB.Text;          ////  Set uer's last name

                addUserToAD.DisplayName = fullName;             ////  Set user's full name to be displayed in the Active Directory list

                addUserToAD.Name = fullName;                    ////  Set user's full name


                addUserToAD.SetPassword(initialPassword);       ////  Set user's initial password

                addUserToAD.Enabled = true;                     ////  Sets the user's account to active

                try
                {
                    addUserToAD.Save();     ///  Add user's profile to Active Directory
                }
                catch (Exception f)
                {
                    System.Windows.Forms.MessageBox.Show("Exception creating user object." + f);
                }


                ////////////////////////////////////////////////////////////////////////
                ////  Add additional information to user's Active Directory profile ////
                ////////////////////////////////////////////////////////////////////////


                ////  Establish connection to existing user profile
                if (addUserToAD.GetUnderlyingObjectType() == typeof(DirectoryEntry))
                {
                    DirectoryEntry entry = (DirectoryEntry)addUserToAD.GetUnderlyingObject();

                    entry.Properties["physicalDeliveryOfficeName"].Value = officeLocation;          ////  Sets user's office location


                    ////  Sets user's manager name, unless they have no manager
                    if (ManagerComboBoxDisplay.SelectedItem != null && ManagerComboBoxDisplay.SelectedItem.ToString() != "No Manager")
                    {
                        PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                        UserPrincipal managerOUInfo = UserPrincipal.FindByIdentity(ctx, ManagerComboBoxDisplay.SelectedItem.ToString());

                        string managerOU = managerOUInfo.DistinguishedName.ToString();

                        entry.Properties["manager"].Clear();
                        entry.Properties["manager"].Value = managerOU;
                    }
                    else
                    {
                        entry.Properties["manager"].Clear();
                    }


                    entry.Properties["department"].Value = departmentName;                  ////  Sets user's department

                    entry.Properties["title"].Value = jobTitleTB.Text;                      ////  Sets user's job title

                    entry.Properties["company"].Value = "Environmental Works, Inc.";        ////  Sets the company name

                    entry.Properties["pwdLastSet"].Value = 0;                               ////  Requires user to change password on first log on to the system

                    entry.Properties["homeDrive"].Value = "[Drive Letter]";                             ////  Sets user's home drive

                    entry.Properties["homeDirectory"].Value = "[Home Drive Address]";        ////  Creates the user's MyDocuments folder on the server


                    ////  Adds any text in the comments box
                    if (descriptionTextBox.Text.Trim().Length > 0)
                    {
                        entry.Properties["description"].Clear();
                        entry.Properties["description"].Value = descriptionTextBox.Text;
                    }

                    try
                    {
                        ////  Save changes to the Active Directory profile
                        entry.CommitChanges();
                        System.Windows.Forms.MessageBox.Show(successMessage);

                    }
                    catch (Exception g)
                    {
                        System.Windows.Forms.MessageBox.Show("Failed to add user to Active Directory:  " + g);
                    }


                }
        



                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////  Defines default groups dependent upon department and office location.  Code Omitted for sample.  ////
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////



                //////////////////////////////////////////////////////////////////////////////
                ////  Create user's mailbox on Exchange server using Exchange Powershell  ////
                //////////////////////////////////////////////////////////////////////////////

                #region  Create Mailbox

                //// Set secure password to connect to Exchange server
                SecureString password = new SecureString();
                string str_password = "[Administrative Password]";
                string username = "[Username]";

                string liveIdconnectionUri = "[Powershell uri]";      ////  Set connection uri as a string

                ////  This formats the password into a secure password
                foreach (char x in str_password)
                {
                    password.AppendChar(x);
                }

                PSCredential credential = new PSCredential(username, password);         ////  Sets administrator credentials for connection

                ////  Creates connection to Exchange Powershell
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(liveIdconnectionUri)), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);

                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default;

                ////  Creates a runspace on a remote path
                Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);

                ////  Creates the user's Exhange Server Mailbox
                PowerShell powershell = PowerShell.Create();

                ////  Launch Exchange Powershell
                PSCommand command = new PSCommand();

                ////  Powershell commands that create the user's mailbox
                command.AddCommand("Enable-Mailbox");
                command.AddParameter("Identity", fullName);
                command.AddParameter("Alias", userLogonName);
                command.AddParameter("Database", "[Mailbox Database]");
                
                ////  Initializes Powershell script
                powershell.Commands = command;


                try
                {
                    ////  Open the remote runspace
                    runspace.Open();

                    ////  Associate the runspace with powershell
                    powershell.Runspace = runspace;

                    ////  Invoke the Powershell command -> user's mailbox is now created
                    powershell.Invoke();
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    ////  Dispose of the runspace and enable garbage collection
                    runspace.Dispose();
                    runspace = null;


                    //// Finally dispose the powershell and set all variables to null to free up any resources.
                    powershell.Dispose();
                    powershell = null;
                }
                #endregion


            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please fill out this form completely");
                return;
            }

            this.Hide();

        }
        #endregion





        private string fullNameText { get { return firstNameTB.Text + " " + lastNameTB.Text; } }            ////  Makes user's full name global



        
        ///////////////////////////////////////////////////////////////////////////////////////////////
        ////  Add the new employee to assigned user groups named above.  Code omitted for sample.  ////
        ///////////////////////////////////////////////////////////////////////////////////////////////
        



        private string departmentNameText { get { return DepartmentComboBox.SelectedItem.ToString(); } }        ////  Makes department name global
        private string officeLocationText { get { return OfficeComboBox.SelectedItem.ToString(); } }            ////  Makes office location name global




        /// <summary>
        ///   Formats the Distinguished name for the employee's profile location in Active Directory
        /// </summary>
        private string ActiveDirectoryUserLocation()
        {
            string adUserLocation = "[AD profile location]";

            if (departmentNameText == "[Department]" && officeLocationText == "[Office location]")
            {
                adUserLocation = "[AD profile location]";
            }

            else if (departmentNameText != "[Department]" && officeLocationText != "[Office location]" && officeLocationText != "[Office location]")
            {
                adUserLocation = "[AD profile location]";
            }

            else if (departmentNameText == "[Department]" && officeLocationText == "[Office location]")
            {
                adUserLocation = "[AD profile location]";
            }

            return adUserLocation;
        }

    }

}
