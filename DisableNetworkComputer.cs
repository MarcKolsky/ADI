using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;

namespace Active_Directory_Interface
{
    public class DisableNetworkComputer
    {


        private string computerName;



        /// <summary>
        ///   Searches network to find our which computer the user is logged into.
        /// </summary>
        private void GetADComputers()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {

                // Specifies the location of the current users %AppData roaming folder. 
                string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                // Combine the base folder with your specific folder....
                string specificFolder = System.IO.Path.Combine(folder, "Active Directory Interface");

                // Check if %AppData roaming directory exists, if not, then create it
                if (!Directory.Exists(specificFolder))
                    Directory.CreateDirectory(specificFolder);

                // Creates a temp file to save the output from wmic
                string tempFile = specificFolder + "\\Computer_Name.txt";


                // Sets the selected user name [from a WPF interface] to a string
                string userName = selectUserCB.SelectedItem.ToString();


                //  Retrieve employee's username from AD
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);


                UserPrincipal userOUInfo = UserPrincipal.FindByIdentity(ctx, userName);
                string userOU = userOUInfo.DistinguishedName.ToString();
                DirectoryEntry getUserDetails = new DirectoryEntry("LDAP://" + userOU);
                string disableUserName = getUserDetails.Properties["sAMAccountName"].Value.ToString();

                ////  Connects to computer folder in Active Directory
                DirectoryEntry entry = new DirectoryEntry("LDAP://[ Organizational User Location ]");
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(name=*)");

                ////  Loops through each computer name in Active Directory
                foreach (SearchResult resEnt in mySearcher.FindAll())
                {

                    string Names = resEnt.GetDirectoryEntry().Name.ToString();

                    ////  Excludes all servers
                    if (Names.StartsWith("CN="))
                    {
                        Names = Names.Remove(0, "CN=".Length);          ////
                    }

                    ////  Ensures any inactive computers are not pinged
                    if (Names.StartsWith("ENV") && Names != "ENVLTAB")
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

                            cmdStartInfo.StandardInput.WriteLine("wmic.exe /node:" + Names + " /user:[ Domain Admin Username ] /password:[ Admin Password ] ComputerSystem Get UserName > " + tempFile);        ////  Initiates command to see who is logged in and writes the output to a temp file.
                            cmdStartInfo.StandardInput.WriteLine("exit");       //// Exit CMD

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
        ///   After the user's account has been disabled, this method searches the network to find what computer the user is logged into and pushes it into hibernation.
        /// </summary>
        private void HibernateUserComputer()
        {

            ////  Connection credentials to access the computer remotely
            ConnectionOptions disableComputer = new ConnectionOptions();
            disableComputer.Username = "[ Domain Admin ]";
            disableComputer.Password = "[ Admin Password ]";
            disableComputer.Authority = "Kerberos:[ Network Domain Name ]\\" + computerName;
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
    }
}
