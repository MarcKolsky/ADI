using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using Microsoft.PowerShell;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.NetworkInformation;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Management;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Active_Directory_Interface
{
    /// <summary>
    /// Interaction logic for DisableUserProfile.xaml
    /// </summary>
    public partial class DisableUserProfile : Window
    {
        string computerName;
        private readonly BackgroundWorker _bw = new BackgroundWorker();     //// Initializes background worker

        public DisableUserProfile()
        {
            ////  Esctablish connection to Active Directory
            PrincipalContext insPrincipalContext = new PrincipalContext(ContextType.Domain, "[Domain]", "[Active Directory Address]");

            InitializeComponent();


            ////  Gets all users in Active Directory
            UserPrincipal insUserPrincipal = new UserPrincipal(insPrincipalContext);
            insUserPrincipal.Name = "*";
            GetUsers(insUserPrincipal);

            _bw.DoWork += RemoveMailBoxFeatures;                ////  Backgroundworker start
            _bw.RunWorkerCompleted += BwRunWorkerCompleted;     ////  Backgroundworker completed

        }



        /// <summary>
        ///   Adds the name of all active users in Active Directory to the combobox.
        /// </summary>
        public void GetUsers(UserPrincipal parUserPrincipal)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            List<string> names = new List<string>();

            selectUserCB.Items.Clear();
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

            names.Sort();       ////  Sorts all names in the list alphabetically

            ////  Adds names to combo boxes
            names.ForEach(delegate(String name)
            {
                selectUserCB.Items.Add(name);
                emailForwardUserCB.Items.Add(name);
            });

        }
        

        /// <summary>
        ///   After the user's account has been disabled, this method searches the network to find what computer the user is logged into and pushes it into hibernation.
        /// </summary>
        private void HibernateUserComputer()
        {

            ////  Connection credentials to access the computer remotely
            ConnectionOptions disableComputer = new ConnectionOptions();
            disableComputer.Username = "[Credentials]";
            disableComputer.Password = "[Password]";
            disableComputer.Authority = "[Network uri]" + computerName;
            disableComputer.Impersonation = ImpersonationLevel.Impersonate;
            disableComputer.EnablePrivileges = true;


            ////  Establishes the connection to the user's computer
            ManagementScope scope = new ManagementScope("\\\\" + computerName + "\\root\\cimv2", disableComputer);
            scope.Connect();


            ////  Launches CMD to intiate the hibernation sequence
            ObjectGetOptions objectGetOptions = new ObjectGetOptions(null, System.TimeSpan.MaxValue, true);
            ManagementPath managementPath = new ManagementPath("Win32_Process");
            ManagementClass processClass = new ManagementClass(scope, managementPath, objectGetOptions);
            ManagementBaseObject inParameters = processClass.GetMethodParameters("Create");
            inParameters["CommandLine"] = "shutdown /h";
            ManagementBaseObject outParameters = processClass.InvokeMethod("Create", inParameters, null);

        }



        /// <summary>
        ///   Searches network to find our which computer the user is logged into.
        /// </summary>
        private void GetADComputers()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {

                ////////////////////////////////////////////////////////////////////////////////////////////
                ////  Creates the folder location on the user's local drive.  Code omitted for sample.  ////
                ////////////////////////////////////////////////////////////////////////////////////////////


                // Sets the selected user name to a string
                string userName = selectUserCB.SelectedItem.ToString();


                //  Retrieve employee's username from AD
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);


                UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, userName);
                string userOU = userOUInfo.DistinguishedName.ToString();
                DirectoryEntry getUserDetails = new DirectoryEntry("LDAP://" + userOU);
                string disableUserName = getUserDetails.Properties["sAMAccountName"].Value.ToString();

                /////////////////////////////////////////////////////////////////////////////////////////
                ////  Retrieves all computer names from Active Directory.  Code omitted for sample.  ////
                /////////////////////////////////////////////////////////////////////////////////////////


                ////  Loops through each computer name in Active Directory
                foreach (SearchResult resEnt in mySearcher.FindAll())
                {

                    string Names = resEnt.GetDirectoryEntry().Name.ToString();

                    ////  Excludes all servers
                    if (Names.StartsWith("CN="))
                    {
                        Names = Names.Remove(0, "CN=".Length);
                    }

                    ////  Ensures any inactive computers are not pinged
                    if ("[Exceptions added here]")
                    {
                        bool pingable = false;
                        Ping ping = new Ping();

                        ///////////////////////////////////////////////////////////////////////////////
                        ////  Pings each computer to see if it is turned on to reduce wasted time  ////
                        ///////////////////////////////////////////////////////////////////////////////

                        try
                        {
                            PingReply reply = ping.Send(Names);

                            if (reply.Status == IPStatus.Success)
                                pingable = true;
                        }
                        catch (PingException)
                        {
                            pingable = false;
                        }

                        string pingResult = pingable.ToString();        ////  Formats ping result as a string

                        ////  If computer is on, now we can launch CMD remotely to find out who is logged in.
                        if (pingResult == "True")
                        {
                            System.Diagnostics.Process cmdStartInfo = new System.Diagnostics.Process();
                            cmdStartInfo.StartInfo.FileName = System.IO.Path.Combine(Environment.SystemDirectory, "cmd.exe");
                            cmdStartInfo.StartInfo.RedirectStandardInput = true;
                            cmdStartInfo.StartInfo.UseShellExecute = false;
                            cmdStartInfo.StartInfo.RedirectStandardOutput = true;
                            cmdStartInfo.StartInfo.CreateNoWindow = true;               ////  Keeps a window from appearing and disrupting users.
                            cmdStartInfo.Start();

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            ////  Runs CMD command line script to identify who is logged in, and writes the name to a local file.  Code omitted for sample.  ////
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            cmdStartInfo.WaitForExit();         ////  Waits for CMD to close


                            ////  Compares the name in the temp file to the user we are looking for.
                            if (File.ReadAllText(tempFile).Contains(disableUserName))
                            {
                                ////  If we find the user we are looking for, this stores the name as a variable and initiates the next phase.
                                computerName = Names;
                                HibernateUserComputer();
                                break;
                            }

                            // Delete temporary files
                            File.Delete(tempFile);
                        }

                    }

                }

            }));

        }




        /// <summary>
        ///   Activated the process of disabling an employee in Active Directory.
        /// </summary>
        private void disableUserButton_Click(object sender, RoutedEventArgs e)
        {

            if (selectUserCB.SelectedItem != null)
            {
                System.Windows.Forms.Application.DoEvents();

                _bw.RunWorkerAsync();

                ////  A funny little swf file asking you to "Please wait"
                PleaseWait showWindow = new PleaseWait();
                showWindow.ShowDialog();
            }
            else
            {
                return;         ////  Process will not run until they select a user's account to be disabled.
            }
        }







        /// <summary>
        ///  This method disables the user's Active Directory profile, changes password, removes user's group memberships, and re-routes user's e-mail to the selected recipient.
        /// </summary>
        private void RemoveADProperties()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {

                //////////////////////////////////////////////////////////////////////////////////////////////////////
                ////  Removes all properties from employee's Active Directory profile.  Code omitted for sample.  ////
                //////////////////////////////////////////////////////////////////////////////////////////////////////

            }));

            // Throw user's computer into hibernate
            GetADComputers();
        }




        /// <summary>
        ///  Removes email rights from the user's Exchange Server mailbox, and re-routes their email to the selected email recipient.
        /// </summary>
        private void RemoveMailBoxFeatures(object sender, DoWorkEventArgs doWorkEventArgs)
        {
            ////  Pauses the process while the "Please Wait" animation finishes
            Thread.Sleep(5500);



            this.Dispatcher.Invoke((Action)(() =>
            {
                ////  Sets user's name to a string
                string userName = selectUserCB.SelectedItem.ToString();

                ////  Sets new email recipient's name to a string
                string emailForwardName = emailForwardUserCB.SelectedItem.ToString();

                
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                ////  Get employee's username from AD
                UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, userName);
                string userOU = userOUInfo.DistinguishedName.ToString();
                DirectoryEntry getUserDetails = new DirectoryEntry("LDAP://" + userOU);

                ////  Get email recipient's username from AD
                UserPrincipal forwardEmailUserOUInfo = UserPrincipal.FindByIdentity(ctx, emailForwardName);
                string forwardEmailUserOU = forwardEmailUserOUInfo.DistinguishedName.ToString();
                DirectoryEntry getDetails = new DirectoryEntry("LDAP://" + forwardEmailUserOU);



                ////////////////////////////////////////////////////////////////////////////////////////
                ////  Connection information to the Exchange powershell.  Code omitted for sample.  ////
                ////////////////////////////////////////////////////////////////////////////////////////


                string emailAddress = getUserDetails.Properties["mail"].Value.ToString();       //// Retrieves email address of disabled user and sets it to a string.
                string forwardEmailAddress = getDetails.Properties["mail"].Value.ToString();    //// Retrieves email address of the new email recipient and sets it to a string.


                ////  Adds two commands to the Exchange Powershell session, and removes the user's mailbox permission and forwards their email to the new recipient.
                command.AddScript("Set-CASMailbox -Identity " + emailAddress + " -OWAEnabled:$false  -ActiveSyncEnabled:$false  -MAPIEnabled:$false  -PopEnabled:$false  -ImapEnabled:$false");
                command.AddScript("Set-Mailbox -Identity " + emailAddress + " -ForwardingAddress " + forwardEmailAddress + "");

                
                powershell.Commands = command;      ////  Sets commands to be passed to powershell


                try
                {
                    //// Open the remote runspace
                    runspace.Open();

                    //// Associate the runspace with powershell
                    powershell.Runspace = runspace;

                    //// Run powershell commands
                    powershell.Invoke();
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    //// dispose the runspace and enable garbage collection
                    runspace.Dispose();
                    runspace = null;


                    //// Finally dispose the powershell and set all variables to null to free up any resources.
                    powershell.Dispose();
                    powershell = null;
                }
            }));

            //// Remove AD Information & Memberships
            RemoveADProperties();
            
        }






        /// <summary>
        ///  Cycles through each group in Active Directory, and removes the user if a member.
        /// </summary>
        private void RemoveADMemberships()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////  Cycles through all user groups, and identifies whether the employee is a member or not.  If they are a member, they are removed.  Code omitted for sample.  ////
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            }));

        }







        /// <summary>
        ///  Process Complete!!!
        /// </summary>
        private void BwRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ////  Closes all windows back to the main window upon completion.
            for (int intCounter = App.Current.Windows.Count - 1; intCounter >= 1; intCounter--)
                App.Current.Windows[intCounter].Close();

            //// Display completed message
            System.Windows.Forms.MessageBox.Show("Completed");
        }

    }
}
