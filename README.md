ADI
===

Active Directory Interface

This is an application I developed for a company that needed an easier way Human Resources to add new employees to the company's network.  The process of adding, editing, or deleting an employee from a company's network is virtually the same at every company.  Human Resources will submit a request to the IT department, and then have to wait for the request to be fulfilled.

By developing this application, the company was able to give a little freedom to the HR department by enabling them to add and edit employees without having to rely solely on the IT department.  The application was built with a very strict set of rules to manage the application, and guard against potentially harmful mistake to Active Directory and Exchange server.

With this application, a user could easily add, edit, disable, and re-enable an employee’s network account.  The application was built with presets for the various offices and departments at the company, so depending on what position an employee had at the company and which office they were located in, the application would know what security groups to add to the employees new Active Directory profile.  After the AD Profile was created, the employees email account was created using Exchange PowerShell and added desired distribution groups.  Finally a network folder on the company's shared drive was created.  A process that could one take days to create could now be completed in 2 minutes.

Editing an employee’s AD account was just as easy.  Should an employee’s name be spelled wrong, get married, or transfer departments, HR could easily update a user’s network account.  This part of the application is very restrictive to protect against inadvertent errors.  This is also where a former employee’s account could be re-enabled.

Finally, if an employee was to be let go it is imperative that their network access be disable.  With this application the user could select an active employee’s name, and then click a single button to disable an employee’s network access, email, and computer.  As a safety feature, I added the ability for the application to search the network, find which computer a specific employee was logged into, and force the computer into hibernation mode.

Because this application was built for another company, I cannot provide the code.  You can however find screenshots of the application's interface, and a demonstration of code to remotely disable a computer on a network.  The provided code helps demonstrate my thought process for having an end goal, and developing a process to achieve it.
