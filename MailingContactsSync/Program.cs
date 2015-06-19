using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Reflection;
using System.DirectoryServices;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;

namespace MailingContactsSync
{
    static class config
    {
        public static string GetConfig(string name)
        {
            string value = null;
            try
            {
                // Узнаем где расположена текущая сборка
                String assempbyLocation = Assembly.GetExecutingAssembly().Location;

                // Открываем наш конфигурационный файл
                Configuration config = ConfigurationManager.OpenExeConfiguration(assempbyLocation);

                // Get the configuration file.
                // System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                //Get the appSettings section.
                AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("appSettings");
                if (appSettings != null)
                {
                    foreach (string key in appSettings.Settings.AllKeys)
                    {
                        if (name == key)
                        {
                            value = appSettings.Settings[key].Value;
                            return value;
                        }
                    }
                }
            }
            catch (Exception EXr)
            {
                value = null;
                Console.WriteLine(EXr.Message);
                Console.WriteLine(EXr.ToString());
            }
            return value;
        }
    }
    static class PS
    {
       
        public static bool CreateContact(string name, string OU, string Alias, string externalsmtp, string DName, string usr, System.Security.SecureString pwd)
        {
            bool create = false;
            try            
            {               
               
                PSCredential ExchangeCredential = new PSCredential(usr, pwd);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri(config.GetConfig("pspath")), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", ExchangeCredential);

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                connectionInfo.SkipCACheck = true;
                connectionInfo.SkipCNCheck = true;
                connectionInfo.SkipRevocationCheck = true;
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);                
                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                
                //PipelineReader<object> errors;

                Command enableMailCmd = new Command("New-MailContact");
                enableMailCmd.Parameters.Add("Name", name);
                enableMailCmd.Parameters.Add("ExternalEmailAddress", externalsmtp);
                enableMailCmd.Parameters.Add("Alias", Alias);
                enableMailCmd.Parameters.Add("OrganizationalUnit", OU);
                enableMailCmd.Parameters.Add("DisplayName", DName);


                command.AddCommand(enableMailCmd);
                powershell.Commands = command;

                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();


                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.AppendLine(String.Empty);
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString());
                }


                if (results.Count > 0)
                {
                    create = true;
                    // log success code goes here                        
                    Console.WriteLine(stringBuilder.ToString());
                    System.Threading.Thread.Sleep(5000);
                }
                else
                {
                    create = false;
                }


                runspace.Close();
                runspace = null;

               
            }
            catch (Exception ex)
            {
                create = false;
                Console.WriteLine(ex.Message);
            }
            return create;
        }
        public static bool CreateGroup(string name, string OU, string usr, System.Security.SecureString pwd)
        {
            bool create = false;
            try
            {
                PSCredential ExchangeCredential = new PSCredential(usr, pwd);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri(config.GetConfig("pspath")), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", ExchangeCredential);

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                connectionInfo.SkipCACheck = true;
                connectionInfo.SkipCNCheck = true;
                connectionInfo.SkipRevocationCheck = true;
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);              
                PowerShell powershell = PowerShell.Create();                
                PSCommand command = new PSCommand();

                //PipelineReader<object> errors;

                Command enableMailCmd = new Command("New-DistributionGroup");
                enableMailCmd.Parameters.Add("SamAccountName", name);
                enableMailCmd.Parameters.Add("Alias", name);
                enableMailCmd.Parameters.Add("Name", name);
                enableMailCmd.Parameters.Add("OrganizationalUnit", OU);
                enableMailCmd.Parameters.Add("Type", "Distribution");

                command.AddCommand(enableMailCmd);
                powershell.Commands = command;

                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSObject> results = powershell.Invoke();


                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.AppendLine(String.Empty);
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString());
                }


                if (results.Count > 0)
                {
                    create = true;
                    // log success code goes here                        
                    Console.WriteLine(stringBuilder.ToString());
                    System.Threading.Thread.Sleep(5000);
                }
                else
                {
                    create = false;
                }


                runspace.Close();
                runspace = null;
                /*
                                Runspace myRunspace = RunspaceFactory.CreateRunspace();
                                myRunspace.Open();

                                RunspaceConfiguration rsConfig = RunspaceConfiguration.Create();
                                PSSnapInException snapInException = null;


                                PSSnapInInfo info = rsConfig.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out snapInException);

                                Runspace myRunSpace = RunspaceFactory.CreateRunspace(rsConfig);
                                myRunSpace.Open();
                                Pipeline pipeline = myRunSpace.CreatePipeline();

                                using (pipeline)
                                {
                                    // New-MailContact -Name <String> -ExternalEmailAddress "" -Alias "" -OrganizationalUnit config.GetConfig("oupath")
                                    //New-DistributionGroup
                                    // -Name
                                    // -Type Distribution
                                    //-OrganizationalUnit
                                    Command enableMailCmd = new Command("New-DistributionGroup");
                                    enableMailCmd.Parameters.Add("SamAccountName", name);
                                    enableMailCmd.Parameters.Add("Alias", name);
                                    enableMailCmd.Parameters.Add("Name", name);
                                    enableMailCmd.Parameters.Add("OrganizationalUnit", OU);
                                    enableMailCmd.Parameters.Add("Type", "Distribution");

                                    pipeline.Commands.Add(enableMailCmd);

                                    Collection<PSObject> cmdResults = pipeline.Invoke();

                                    StringBuilder stringBuilder = new StringBuilder();
                                    stringBuilder.AppendLine(String.Empty);
                                    foreach (PSObject obj in cmdResults)
                                    {
                                        stringBuilder.AppendLine(obj.ToString());
                                    }


                                    if (cmdResults.Count > 0)
                                    {
                                        create = true;
                                        // log success code goes here                        
                                        Console.WriteLine(stringBuilder.ToString());
                                        System.Threading.Thread.Sleep(5000);
                                    }
                                    else
                                    {
                                        create = false;
                                    }


                                    myRunspace.Close();
                                    myRunspace = null;

                                }
                                */
            }//end  try

            catch (Exception ex)
            {
                create = false;
                Console.WriteLine(ex.Message);
            }
            return create;
        }
    }
    static class AD
    {
        public static string TranslatePath(string path)
        {
            string ret = path;
            if (!string.IsNullOrEmpty(ret))
            {
                if (ret.ToLower().StartsWith(@"ldap://") && !ret.ToLower().StartsWith(config.GetConfig("ldappath")))
                {
                    ret = config.GetConfig("ldappath") + ret.Substring(7);
                }
                else if (ret.ToLower().StartsWith(@"CN="))
                {
                    ret = config.GetConfig("ldapcn") + ret.Substring(7);
                }
            }
            return ret;
        }
        public static string TranslateName(string name)
        {
            string ret = "";
            string MessageIDPattern = @"[а-яА-Яa-zA-Z0-9\,\{\}\(\)\-_ ]{1}";
            foreach (char l in name)
            {
                    if (Regex.IsMatch(l.ToString(), MessageIDPattern))
                    {
                        ret += l;
                    }                
            }
            return ret.Replace("  ", " ").Replace("( ", "(").Replace(") ", ")").Replace("  ", " ").Trim();
        }
        public static string GetPath(string Name, string usrname, string usrpass, string objectClass, string objectCategory)
        {
            try
            {
                string ADPath = config.GetConfig("ADPath");

                string filter = @"(&(objectCategory=" + objectCategory + ")(objectClass=" + objectClass + ")(name=" + Name + "))";
                int rescount = 0;
                string path = null;

                using (DirectoryEntry de = new DirectoryEntry(ADPath, usrname, usrpass))
                {
                    DirectorySearcher ds = new DirectorySearcher(de, filter);
                    ds.SearchScope = SearchScope.Subtree;
                    path = null;
                    rescount = 0;
                    foreach (SearchResult results in ds.FindAll())
                    {
                        if (results != null)
                        {
                            rescount++;
                            path = results.Path;
                            path = AD.TranslatePath(path);
                        }
                    }

                }
                if (rescount == 1)
                {
                    return path;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return null;
        }
        public static string GetPathWithSMTP(string smtp, string usrname, string usrpass)
        {
            try
            {
                string ADPath = config.GetConfig("ADPath");

                string filter = @"(&(proxyAddresses=*:" + smtp + "*))";
                int rescount = 0;
                string path = null;
                
                using (DirectoryEntry de = new DirectoryEntry(ADPath, usrname, usrpass))
                {
                    DirectorySearcher ds = new DirectorySearcher(de, filter);
                    ds.SearchScope = SearchScope.Subtree;
                    path = null;
                    rescount = 0;
                    foreach (SearchResult results in ds.FindAll())
                    {
                        if (results != null)
                        {
                            rescount++;
                            path = results.Path;
                         //   Console.WriteLine("1Path:"+path);
                            path = AD.TranslatePath(path);
                         //   Console.WriteLine("2Path:" + path);
                        }
                    }

                }
               // Console.WriteLine("1rescount:" + rescount);
                if (rescount == 1)
                {
                    return path;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return null;
        }
        public static string GetAttribute(string path, string name, string usrname, string usrpass)
        {
            string ret = null;
            path = AD.TranslatePath(path);
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                ret = de.Properties[name].Value.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return ret;
        }
        public static ArrayList GetAttributeArray(string path, string name, string usrname, string usrpass)
        {
            path = AD.TranslatePath(path);
            ArrayList ret = new ArrayList();
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                if (de != null)
                {
                    if (de.Properties[name] != null)
                    {
                        if (de.Properties[name].Count > 0)
                        {
                            foreach (string vl in de.Properties[name])
                            {
                                ret.Add(vl);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = new ArrayList();
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return ret;
        }
        public static void ClearAttribute(string path, string name, string usrname, string usrpass)
        {
            path = AD.TranslatePath(path);
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                de.Properties[name].Clear();
                de.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        public static void SetAttribute(string path, string name, string value, string usrname, string usrpass)
        {
            path = AD.TranslatePath(path);
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                de.Properties[name].Value = value;
                de.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        public static void SetAttribute(string path, string name, bool value, string usrname, string usrpass)
        {
           // Console.WriteLine("SetAttribute: " + path + " | " + name + " | " + value);
            path = AD.TranslatePath(path);
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                de.Properties[name].Value = value;
                de.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        public static void AddAttribute(string path, string name, string value, string usrname, string usrpass)
        {
            //Console.WriteLine("AddAttribute: " + path +" | "+ name + " | " + value);
            path = AD.TranslatePath(path);
            try
            {
                
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                de.Properties[name].Add(value);
                de.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        public static void RemoveAttribute(string path, string name, string value, string usrname, string usrpass)
        {
            path = AD.TranslatePath(path);
            try
            {
                DirectoryEntry de = new DirectoryEntry(path, usrname, usrpass);
                de.Properties[name].Remove(value);
                de.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
        }
        public static bool CheckForUpdateContact(string Name, string Dname, ArrayList SMTPs, string usrname, string usrpass, string objectClass, string objectCategory)
        {
            bool needUpdate = false;
            try
            {
                string path = GetPath(Name, usrname, usrpass, objectClass, objectCategory);
                path = AD.TranslatePath(path);
                if (GetAttribute(path, "displayName", usrname, usrpass) != Dname)
                {
                    needUpdate = true;
                }
                ArrayList DNs = new ArrayList();
                foreach (string smtpadr in SMTPs)
                {
                    if (!String.IsNullOrEmpty(smtpadr))
                    {
                        if (!String.IsNullOrEmpty(smtpadr))
                        {
                            string paths = GetPathWithSMTP(smtpadr, usrname, usrpass);
                            if (!String.IsNullOrEmpty(paths))
                            {
                                string value = GetAttribute(paths, "distinguishedName", usrname, usrpass);
                                if (!String.IsNullOrEmpty(value))
                                {
                                    DNs.Add(value);
                                }
                            }
                        }
                    }
                }
                ArrayList DNCs = new ArrayList();
                DNCs = GetAttributeArray(path, "authOrig", usrname, usrpass);
                foreach (string DN in DNs)
                {
                    if (!DNCs.Contains(DN))
                    {
                        needUpdate = true;
                        return true;
                    }
                }
                foreach (string DN in DNCs)
                {
                    if (!DNs.Contains(DN))
                    {
                        needUpdate = true;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                needUpdate = false;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return needUpdate;
        }
        public static bool CheckForUpdateGroup(string Name, ArrayList SMTPs, string usrname, string usrpass, string objectClass, string objectCategory)
        {
            bool needUpdate = false;
            try
            {
                string path = GetPath(Name, usrname, usrpass, objectClass, objectCategory);
                path = AD.TranslatePath(path);
                ArrayList DNs = new ArrayList();
                foreach (string smtp in SMTPs)
                {
                    if (!String.IsNullOrEmpty(smtp))
                    {
                       // Console.WriteLine("1value:" + smtp);
                        string value = GetPathWithSMTP(smtp, usrname, usrpass);
                       // Console.WriteLine("2value:" + smtp + "[" + value + "]");
                        if (!String.IsNullOrEmpty(value))
                        {
                            value = GetAttribute(value, "distinguishedName", usrname, usrpass);
                            if (!String.IsNullOrEmpty(value))
                            {
                                DNs = GetAttributeArray(path, "member", usrname, usrpass);
                                if (!DNs.Contains(value))
                                {
                                    needUpdate = true;
                                    return needUpdate;
                                }
                                else
                                {
                                    //    Console.WriteLine("GetAttributeArray(" + path + ", member, usrname, usrpass).Contains(" + value + "))");
                                }
                            }
                        }
                    }
                }


                foreach (string dn in DNs)
                {
                    bool found = false;
                    ArrayList pa = new ArrayList();
                    pa = GetAttributeArray("LDAP://" + dn, "proxyAddresses", usrname, usrpass);
                    foreach (string sm in pa)
                    {
                        if (!String.IsNullOrEmpty(sm))
                        {
                            if (sm.ToLower().StartsWith("smtp:"))
                            {
                                string smtp1 = sm.Substring(5);
                                foreach (string smtp2 in SMTPs)
                                {
                                    if (smtp1.ToLower() == smtp2.ToLower())
                                    {
                                        Console.ForegroundColor = ConsoleColor.Cyan;
                                        Console.WriteLine(smtp1 + " " + smtp2);
                                        Console.ResetColor();
                                        found = true;
                                    }
                                }
                            }
                        }
                    }
                    if (!found)
                    {
                        needUpdate = true;
                        return needUpdate;
                    }
                }
                ArrayList Adressb = GetAttributeArray(path, "showInAddressBook", usrname, usrpass);
                if (Adressb != null)
                {
                    if (Adressb.Count > 0)
                    {
                        needUpdate = true;
                        return needUpdate;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                needUpdate = false;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
            return needUpdate;
        }
        public static int Found(string Name, string usrname, string usrpass, string objectClass, string objectCategory)
        {
            string ADPath = config.GetConfig("ADPath");

            string filter = @"(&(objectCategory=" + objectCategory + ")(objectClass=" + objectClass + ")(name=" + Name + "))";
            int rescount = 0;
            try
            {
                using (DirectoryEntry de = new DirectoryEntry(ADPath, usrname, usrpass))
                {
                    DirectorySearcher ds = new DirectorySearcher(de, filter);
                    ds.SearchScope = SearchScope.Subtree;

                    rescount = 0;
                    foreach (SearchResult results in ds.FindAll())
                    {
                        if (results != null)
                        {
                            rescount++;

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
            return rescount;
        }
        public static void ModifyContact(string Name, string Dname, ArrayList memsmtp, string alias, string exsmtp, string insmtp, string ADresslist, string extensionAttribute3Value, string objectClass, string objectCategory, string usrname, string usrpass)
        {
            try
            {
                string path = GetPath(Name, usrname, usrpass, objectClass, objectCategory);
                path = AD.TranslatePath(path);
                string attr = "showInAddressBook";
                ClearAttribute(path, attr, usrname, usrpass);
                string value = ADresslist;
                AddAttribute(path, attr, value, usrname, usrpass);
                attr = "extensionAttribute3";
                value = extensionAttribute3Value;
                SetAttribute(path, attr, value, usrname, usrpass);
                attr = "msExchPoliciesIncluded";
                ClearAttribute(path, attr, usrname, usrpass);
                attr = "proxyAddresses";
                ClearAttribute(path, attr, usrname, usrpass);
                value = "SMTP:" + exsmtp;
                AddAttribute(path, attr, value, usrname, usrpass);
                value = "smtp:" + insmtp;
                AddAttribute(path, attr, value, usrname, usrpass);
                attr = "displayName";
                value = Dname;
                SetAttribute(path, attr, value, usrname, usrpass);
                attr = "authOrig";
                ClearAttribute(path, attr, usrname, usrpass);

                foreach (string smtpadr in memsmtp)
                {
                    if (!String.IsNullOrEmpty(smtpadr))
                    {
                        if (!String.IsNullOrEmpty(smtpadr))
                        {
                            value = GetPathWithSMTP(smtpadr, usrname, usrpass);
                            if (!String.IsNullOrEmpty(value))
                            {
                                value = GetAttribute(value, "distinguishedName", usrname, usrpass);
                                if (!String.IsNullOrEmpty(value))
                                {
                                    AddAttribute(path, attr, value, usrname, usrpass);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
        }
        public static void ModifyGroup(string ID, string Name, ArrayList memsmtp, string smtp, string objectClass, string objectCategory, string usrname, string usrpass)
        {
            try
            {
                string path = GetPath(Name, usrname, usrpass, objectClass, objectCategory);
                path = AD.TranslatePath(path);

                //Console.WriteLine("ModifyGroup: " + path + " | " + Name + " | " + objectClass + " | " + objectCategory);

                string attr = "msExchHideFromAddressLists";
                SetAttribute(path, attr, true, usrname, usrpass);
                attr = "showInAddressBook";
                ClearAttribute(path, attr, usrname, usrpass);
                attr = "msExchPoliciesIncluded";
                ClearAttribute(path, attr, usrname, usrpass);
                attr = "proxyAddresses";
                ClearAttribute(path, attr, usrname, usrpass);
                string value = "SMTP:" + smtp;
                AddAttribute(path, attr, value, usrname, usrpass);
                attr = "member";
                ClearAttribute(path, attr, usrname, usrpass);
                foreach (string smtpadr in memsmtp)
                {
                    if (!String.IsNullOrEmpty(smtpadr))
                    {
                        if (!String.IsNullOrEmpty(smtpadr))
                        {                           
                            value = GetPathWithSMTP(smtpadr, usrname, usrpass);

                            if (!String.IsNullOrEmpty(value))
                            {
                                value = GetAttribute(value, "distinguishedName", usrname, usrpass);
                                if (!String.IsNullOrEmpty(value))
                                {
                                    AddAttribute(path, attr, value, usrname, usrpass);
                                }
                            }
                        }
                    }
                }
                attr = "authOrig";
                ClearAttribute(path, attr, usrname, usrpass);
                value = config.GetConfig("AdminPath");
                AddAttribute(path, attr, value, usrname, usrpass);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
        }
    }
    /*static class Fenix
    {
        public static Hashtable GetMailings()
        {
            Hashtable ht = new Hashtable();
            try
            {
                string connstring = config.GetConfig("Connstring");
                SqlConnection con = new SqlConnection(connstring);
                con.Open();
                SqlCommand cmd = new SqlCommand();

                cmd.CommandType = CommandType.Text;

                cmd.CommandText = @"DECLARE	@return_value int, " +
            @"@SName varchar(256), " +
            @"@Res varchar(max) " +
    @"EXEC	@return_value = [dbo].[sp_SubscrForExchange] " +
            @"@SID = 0, " +
            @"@Delim = N';', " +
            @"@SName = @SName OUTPUT, " +
            @"@Res = @Res OUTPUT " +
    @"SELECT @SName as N'@SName', " +
            @"@Res as N'@Res';";

                cmd.Connection = con;


                //Console.WriteLine(cmd.ExecuteScalar().ToString());

                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    try
                    {
                        // Console.WriteLine("1:" + reader[0].ToString() + ";2:" + reader[1].ToString() + ";3:" + reader[2].ToString() + ";4:" + reader[3].ToString() + ";5:" + reader[4].ToString());
                        Console.WriteLine(reader[0].ToString() + ": " + reader[1].ToString());
                        ht.Add(reader[0].ToString(), reader[1].ToString());
                    }
                    catch (Exception lex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(lex.Message);
                        Console.WriteLine(lex.ToString());
                        Console.ResetColor();
                    }
                }

                con.Close();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ht = null;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
            return ht;
        }
        public static ArrayList GetACL(string ID)
        {
            ArrayList ret = new ArrayList();
            try
            {
                string connstring = config.GetConfig("Connstring");
                SqlConnection con = new SqlConnection(connstring);
                con.Open();
                SqlCommand cmd = new SqlCommand();

                cmd.CommandType = CommandType.Text;

                cmd.CommandText = @"DECLARE	@return_value int, " +
            @"@SName varchar(256), " +
            @"@Res varchar(max) " +
    @"EXEC	@return_value = [dbo].[sp_SubscrForExchange] " +
            @"@SID = 0, " +
            @"@Delim = N';', " +
            @"@SName = @SName OUTPUT, " +
            @"@Res = @Res OUTPUT " +
    @"SELECT @SName as N'@SName', " +
            @"@Res as N'@Res';";

                cmd.Connection = con;


                //Console.WriteLine(cmd.ExecuteScalar().ToString());

                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    try
                    {
                        if (reader[0].ToString() == ID)
                        {
                            Console.WriteLine(reader[0].ToString() + ": " + reader[2].ToString());
                            string lret = reader[2].ToString();
                            if (!String.IsNullOrEmpty(lret))
                            {
                                if (lret.Contains(";"))
                                {
                                    foreach (string smtp in lret.Split(';'))
                                    {
                                        ret.Add(smtp.Trim());
                                    }
                                }
                                else
                                {
                                    ret.Add(lret);
                                }
                            }
                            return ret;
                        }
                    }
                    catch (Exception lex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(lex.Message);
                        Console.WriteLine(lex.ToString());
                        Console.ResetColor();
                    }
                }

                con.Close();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ret = new ArrayList();
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
            return ret;
        }
    }*/
    class Program
    {
        static void Main(string[] args)
        {


            try
            {
                
                string usr = null;
                //string usrn = null;
                string pass = null;
            //    System.Security.SecureString pwd = null;
                MailingSQLHelper.MLSQLHelper sql = new MailingSQLHelper.MLSQLHelper();

                Hashtable table = sql.GetMLRubricList(true);
                    // Fenix.GetMailings();
                Console.WriteLine(table.Count);
                
                ArrayList GlobalDNs = new ArrayList();
                Console.WriteLine("ENTER USERNAME");
                usr = config.GetConfig("ldappath");

                Console.WriteLine("ENTER PWD");
                Console.BackgroundColor = ConsoleColor.Black;
                Console.ForegroundColor = ConsoleColor.Black;
                pass = config.GetConfig("pwd");
                System.Security.SecureString pwd = new System.Security.SecureString();
                foreach (char pfw in pass)
                {
                    pwd.AppendChar(pfw);
                }
                Console.ResetColor();
                Console.WriteLine("Begin Foreach");
                foreach (DictionaryEntry entry in table)
                {
                    if (entry.Key != null && entry.Value != null)
                    {
                        if ((!String.IsNullOrEmpty(entry.Key.ToString())))
                        {//MailerAgentData.ListRubric
                            string rID = ((MailerAgentData.ListRubric)entry.Value).PrefID;
                            string rSMTP = entry.Key.ToString();
                            Console.WriteLine("--------" + rID + "---------------------");
                            Console.WriteLine("--------" + rSMTP + "---------------------");
                            string name = "ML-ID-" + rID;
                            string Dname = AD.TranslateName(((MailerAgentData.ListRubric)entry.Value).Name);                            
                            string oupath = config.GetConfig("Oupath");
                            string alias = "ML-ID-" + rID;
                            string adsuf = config.GetConfig("Adsuf");
                            string exsmtp = "ML-ID-" + rID + adsuf;
                            string insmtp = "ML-ID-" + rID + adsuf;
                            string addrlist = config.GetConfig("Addrlist");
                            string smtplist = config.GetConfig("Smtplist");
                            string Ext3Value = "Рассылки";
                            Console.WriteLine("name; "+name);
                           // Console.WriteLine("Dname; ["+Dname+"]");
                           // Console.WriteLine("Dname; [" +AD.TranslateName(Dname) + "]");
                            
                            ArrayList ACL = new ArrayList(); //Fenix.GetACL(ID.ToString());
                            if (!string.IsNullOrEmpty(smtplist))
                            {
                                if (smtplist.Contains(";"))
                                {
                                    foreach (string sl in smtplist.Split(';'))
                                    {
                                        if (!string.IsNullOrEmpty(sl))
                                        {
                                            ACL.Add(sl.Trim());
                                        }
                                    }
                                }
                                else
                                {
                                    ACL.Add(smtplist.Trim());
                                }
                            }
                            
                            if (AD.Found(name, usr, pass, "contact", "person") == 0)
                            {
                                if (PS.CreateContact(name, oupath, alias, exsmtp, Dname, usr, pwd))
                                {
                                    int iterfound = 0;
                                    while (AD.Found(name, usr, pass, "contact", "person") == 0)
                                    {
                                        Console.WriteLine("-");
                                        iterfound++;
                                        System.Threading.Thread.Sleep(5000);
                                        if (iterfound > 50)
                                        {
                                            Console.Write(".");
                                            break;
                                        }
                                    }
                                    if (AD.Found(name, usr, pass, "contact", "person") == 1)
                                    {
                                        AD.ModifyContact(name, Dname, ACL, alias, exsmtp, insmtp, addrlist, Ext3Value, "contact", "person", usr, pass);
                                        if (AD.Found(name + "-ACL", usr, pass, "group", "group") == 0)
                                        {
                                            if (PS.CreateGroup(name + "-ACL", oupath, usr, pwd))
                                            {
                                                int iterfoundgr = 0;
                                                while (AD.Found(name + "-ACL", usr, pass, "group", "group") == 0)
                                                {
                                                    Console.WriteLine("-");
                                                    iterfoundgr++;
                                                    System.Threading.Thread.Sleep(5000);
                                                    if (iterfoundgr > 50)
                                                    {
                                                        Console.Write(".");
                                                        break;
                                                    }
                                                }
                                                AD.ModifyGroup(rID, name + "-ACL", ACL, name + "-ACL" + adsuf, "group", "group", usr, pass);
                                            }
                                        }
                                        if (!AD.GetAttributeArray(AD.GetPath("ML-Mail-ID-X", usr, pass, "group", "group"), "member", usr, pass).Contains(AD.GetAttribute(AD.GetPath(name, usr, pass, "contact", "person"), "distinguishedName", usr, pass)))
                                        {
                                            AD.AddAttribute(AD.GetPath("ML-Mail-ID-X", usr, pass, "group", "group"), "member", AD.GetAttribute(AD.GetPath(name, usr, pass, "contact", "person"), "distinguishedName", usr, pass), usr, pass);
                                        }
                                    }
                                }
                                else
                                {
                                     Console.ForegroundColor = ConsoleColor.Red;
                                     Console.WriteLine("Not Create Contact: " + name);
                                     Console.ResetColor();
                                }
                            }
                            else if (AD.Found(name, usr, pass, "contact", "person") == 1)
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("Contact: " + name + " ID: " + rID + " Already Exists");
                                Console.ResetColor();
                                if (AD.CheckForUpdateContact(name, Dname, ACL, usr, pass, "contact", "person"))
                                {
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine("Contact: " + name + " ID: " + rID + " Update needed");
                                    Console.ResetColor();
                                    AD.ModifyContact(name, Dname, ACL, alias, exsmtp, insmtp, addrlist, Ext3Value, "contact", "person", usr, pass);
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("Contact: " + name + " ID: " + rID + " Update Not needed");
                                    Console.ResetColor();
                                }

                                if (AD.Found(name + "-ACL", usr, pass, "group", "group") == 0)
                                {
                                    if (PS.CreateGroup(name + "-ACL", oupath,usr,pwd))
                                    {
                                        AD.ModifyGroup(rID, name + "-ACL", ACL, name + "-ACL" + adsuf, "group", "group", usr, pass);
                                    }
                                }
                                else if (AD.Found(name + "-ACL", usr, pass, "group", "group") == 1)
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("group: " + name + "-ACL" + " Already Exists");
                                    Console.ResetColor();
                                    if (AD.CheckForUpdateGroup(name + "-ACL", ACL, usr, pass, "group", "group"))
                                    {
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.WriteLine("group: " + name + "-ACL" + " Update needed");
                                        Console.ResetColor();
                                        AD.ModifyGroup(rID, name + "-ACL", ACL, name + "-ACL" + adsuf, "group", "group", usr, pass);
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.Green;
                                        Console.WriteLine("group: " + name + "-ACL" + " Update Not needed");
                                        Console.ResetColor();
                                    }
                                }
                            }
                            //check for user have permission to view Address List
                            ArrayList SMTPs = new ArrayList();
                            SMTPs = ACL;
                            if (SMTPs.Count > 0)
                            {
                                foreach (string SMTP in SMTPs)
                                {

                                    if (!String.IsNullOrEmpty(SMTP))
                                    {
                                        string pathes = AD.GetPathWithSMTP(SMTP, usr, pass);

                                       // Console.WriteLine(pathes);

                                        if (!String.IsNullOrEmpty(pathes))
                                        {
                                            string dn1 = AD.GetAttribute(pathes, "distinguishedName", usr, pass);
                                       //     Console.WriteLine("dn1: "+dn1);
                                            GlobalDNs.Add(dn1);

                                            string mlspath =AD.GetPath("ML-Mail-ID-S", usr, pass, "group", "group");
                                            mlspath = AD.TranslatePath(mlspath);
                                            if (!AD.GetAttributeArray(mlspath, "member", usr, pass).Contains(dn1))
                                            {
                                                Console.ForegroundColor = ConsoleColor.Yellow;
                                                Console.WriteLine(dn1 + " Added to ML-Рассылки");
                                                Console.ResetColor();
                                                AD.AddAttribute(mlspath, "member", dn1, usr, pass);
                                            }
                                        }
                                    }
                                }
                            }
                   /**/
                        } 
                    }                    
                }
                
                string mlspath1 = AD.GetPath("ML-Mail-ID-S", usr, pass, "group", "group");
                mlspath1 = AD.TranslatePath(mlspath1);
                foreach (string dn2 in AD.GetAttributeArray(mlspath1, "member", usr, pass))
                {
                    //Console.WriteLine(AD.GetPath("ML-Mail-ID-S", usr, pass, "group", "group"));
                    if (!GlobalDNs.Contains(dn2))
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine(dn2 + " Removed from ML-Mail-ID-S");
                        Console.ResetColor();
                        AD.RemoveAttribute(mlspath1, "member", dn2, usr, pass);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
        }
    }
}
