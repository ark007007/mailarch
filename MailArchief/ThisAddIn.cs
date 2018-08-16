using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using IWshRuntimeLibrary;

namespace MailArchief
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        List<KeyValuePair<string, string>> maiList = new List<KeyValuePair<string, string>>();
        Serializer serial = new Serializer();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Environment.SetEnvironmentVariable("foo", "bar", EnvironmentVariableTarget.Machine);
            try
            {
                //System.Windows.Forms.MessageBox.Show(String.)
                AddACategory();
                Set_clientnrList();
                mappen = serial.Deserialize(mappen, @"D:\data\paths.txt");
                mappen2 = Get_mappen(@"D:\Data");

                outlookNameSpace = this.Application.GetNamespace("MAPI");
                inbox = outlookNameSpace.GetDefaultFolder(
                        Microsoft.Office.Interop.Outlook.
                        OlDefaultFolders.olFolderInbox);

                items = inbox.Items;
                items.ItemAdd +=
                    new Outlook.ItemsEvents_ItemAddEventHandler(New_Item_Handler);
            }
            catch (System.Exception error)
            {

                System.Windows.Forms.MessageBox.Show("error = " + error.Message);
            }
            
        }
        void New_Item_Handler(object Item)
        {
            try
            {
                if (Item is Outlook.MailItem)
                {
                    MailItem mail = (MailItem)Item;

                    Get_clientnr2(mail.SenderEmailAddress, mappen, mail, mappen2);
                }
            }
            catch (System.Exception error)
            {
                System.Windows.Forms.MessageBox.Show("error = " + error.Message);
            }

                


        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion



        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RIbbon();
        }
        public void MakePathFile()
        {
                mappen = Get_mappen(@"D:\clienten");
                mappen2 = Get_mappen(@"D:\data");

                Make_path_to_Text();


        }

        public void Make_path_to_Text()
        {
            //StreamWriter newFile = new StreamWriter(@"C:\data\paths.txt");
            serial.SerialiseList(mappen, @"D:\data\paths.txt");
            //newFile.WriteLine(txt);

            //newFile.Close();
        }
        #region global vars

        public bool use_inbox = true;
        public bool use_sent_items = true;
        List<string> mappen = new List<string>();
        List<string> mappen2 = new List<string>();
        public string Folder1 = "";
        public string Folder2 = "";
        public string Folder3 = "";
        public string Folder4 = "";
        public string Folder5 = "";

        public bool F1Inbox = true;
        public bool F2Inbox = true;
        public bool F3Inbox = true;
        public bool F4Inbox = true;
        public bool F5Inbox = true;

        public string SubF1 = "";
        public string SubF2 = "";
        public string SubF3 = "";
        public string SubF4 = "";
        public string SubF5 = "";
        const OlSaveAsType MAIL_SAVETYPE = OlSaveAsType.olMSG;
        const string MAIL_SAVEFILE_EXT = ".msg";
        const string PATHSEPARATOR = @"\";
        #endregion
        public string GetOfolder()
        {
            string folder_name = "folder";
            Outlook.MAPIFolder pf = Application.Session.PickFolder();
            folder_name = pf.Name;
            return folder_name;
        }

        public void Test_4()
        {

            try
            {

                if (Folder1 != "" && SubF1 == "")
                    Klanten_verzonden_items2(mappen, F1Inbox, Folder1, mappen2);
                if (Folder2 != "" && SubF2 == "")
                    Klanten_verzonden_items2(mappen, F2Inbox, Folder2, mappen2);
                if (Folder3 != "" && SubF3 == "")
                    Klanten_verzonden_items2(mappen, F3Inbox, Folder3, mappen2);
                if (Folder4 != "" && SubF4 == "")
                    Klanten_verzonden_items2(mappen, F4Inbox, Folder4, mappen2);
                if (Folder5 != "" && SubF5 == "")
                    Klanten_verzonden_items2(mappen, F5Inbox, Folder5, mappen2);

                if (Folder1 != "" && SubF1 != "")
                    Klanten_verzonden_items2(mappen, F1Inbox, Folder1, mappen2, SubF1);
                if (Folder2 != "" && SubF2 != "")
                    Klanten_verzonden_items2(mappen, F2Inbox, Folder2, mappen2, SubF2);
                if (Folder3 != "" && SubF3 != "")
                    Klanten_verzonden_items2(mappen, F3Inbox, Folder3, mappen2, SubF3);
                if (Folder4 != "" && SubF4 != "")
                    Klanten_verzonden_items2(mappen, F4Inbox, Folder4, mappen2, SubF4);
                if (Folder5 != "" && SubF5 != "")
                    Klanten_verzonden_items2(mappen, F5Inbox, Folder5, mappen2, SubF5);

                if (use_inbox)
                    Test_2_2(mappen, true, mappen2);
                if (use_sent_items)
                    Test_2_2(mappen, false, mappen2);
            }
            catch (System.Exception error)
            {

                System.Windows.Forms.MessageBox.Show("error = " + error.Message);
            }
        }

        public void Test_2_2(List<string> mappen, bool inbox, List<string> mappen2)
        {
            MAPIFolder ProcessFolder;
            Outlook.Items Items;
            if (inbox == false)
                ProcessFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            else
                ProcessFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            Items = ProcessFolder.Items;
            foreach (object itm in Items)
            {
                if (itm is Outlook.MailItem)
                {
                    MailItem mail = (MailItem)itm;
                    try
                    {

                        if (mail.Categories == null)
                            mail.Categories = "none";
                        if (!mail.Categories.Contains("Automatisch gearchiveerd"))
                        {
                            if (inbox == true)
                                Get_clientnr2(mail.SenderEmailAddress, mappen, mail, mappen2);
                            else
                                foreach (Recipient Recipient in mail.Recipients)
                                    Get_clientnr2(Recipient.Address, mappen, mail, mappen2);
                        }
                    }
                    catch
                    {

                    }

                }
            }
        }


        public void Klanten_verzonden_items2(List<string> mappen, bool inbox, string map_mail, List<string> mappen2, string sub_map = "")
        {
            MAPIFolder ProcessFolder;
            Outlook.Items Items;
            if (inbox == false)
                ProcessFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            else
                ProcessFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            ProcessFolder = ProcessFolder.Folders[map_mail];
            if (sub_map != "")
                ProcessFolder = ProcessFolder.Folders[sub_map];
            Items = ProcessFolder.Items;
            foreach (object itm in Items)
            {
                if (itm is Outlook.MailItem)
                {
                    MailItem mail = (MailItem)itm;
                    try
                    {

                        if (mail.Categories == null)
                            mail.Categories = "none";
                        if (!mail.Categories.Contains("Automatisch gearchiveerd"))
                        {
                            if (inbox == true)
                                Get_clientnr2(mail.SenderEmailAddress, mappen, mail, mappen2);
                            else
                                foreach (Recipient recipient in mail.Recipients)
                                    Get_clientnr2(recipient.Address, mappen, mail, mappen2);
                        }
                    }
                    catch { }
                    
                }

            }
        }
        //voeg nog toe
        /*
        public void map_in_ander_mail_account(List<string> mappen, List<string> mappen2, bool inbox, object mail_account, object map, string map2 = "")
        {
            object Recipient;
            Outlook.Folder ProcessFolder;

            Outlook.Items Items;
            Outlook.NameSpace myNamespace;
            myNamespace = Application.GetNamespace("MAPI");
            if (map2 != null)
                ProcessFolder = myNamespace.Folders(mail_account).Folders(map);
            else
                ProcessFolder = myNamespace.Folders(mail_account).Folders(map).Folders(map2);
            Items = ProcessFolder.Items;
            foreach (MailItem Itm in Items)
            {
                if (Itm is Outlook.MailItem)
                {
                    MailItem mailItem = (MailItem)Itm;
                    if (inbox)
                        get_clientnr(mailItem.SenderEmailAddress, mappen, mailItem, mappen2);
                    else
                        foreach (Recipient Recipiet in mailItem.Recipients)
                            get_clientnr(Recipiet.Address, mappen, mailItem, mappen2);
                }
            }
        }
        */
        public static string RemoveIllegalFileNameChars(string input, string replacement = "")
        {
            var regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            var r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            return r.Replace(input, replacement);
        }

        public List<string> Get_mappen(String sPath)
        {
            List<string> List = new List<string>();
            List = Directory.GetDirectories(sPath, "*", SearchOption.AllDirectories).ToList();
            return List;
        }
        private void UrlShortcut(string clientnr, int year, string Oldpath)
        {
            string sPathFile = @"D:\Data\Corespondentie map automatische mail\" + clientnr + @"\" + year + @"\";
            CreateDir(sPathFile);
            string shortcutLocation = System.IO.Path.Combine(Oldpath, "automatisch gearchiveerde mail" + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);
            shortcut.Description = "My shortcut description";
            shortcut.TargetPath = sPathFile;
            shortcut.Save();

        }

        public void CreateDir(string strPath)
        {
            string strCheckPath;
            strCheckPath = "";
            foreach (var elm in strPath.Split('\\'))
            {
                strCheckPath = strCheckPath + elm + @"\";
                char last = strCheckPath[strCheckPath.Length - 1];
                Directory.CreateDirectory(strCheckPath);
            }
        }
        public void Save_item(List<string> mappen, MailItem Item, string clientnr, string mail, List<string> mappen2)
        {
            string map_doel = "";
            string Item_naam;
            int year = Item.SentOn.Year;
            int month = Item.SentOn.Month;
            int day = Item.SentOn.Day;

            string date = year + "-" + month + "-" + day + " " + Item.SentOn.Hour + " " + Item.SentOn.Minute + " " + Item.SentOn.Second;

            string sPathFile = @"D:\Data\Corespondentie map automatische mail\" + clientnr + @"\" + year + @"\";
            CreateDir(sPathFile);

            foreach (string element in mappen)
            {
                if (element.EndsWith(clientnr + @"\Correspondentie"))
                {
                    CreateDir(element + @"\" + year);
                    map_doel = element + @"\" + year;

                    UrlShortcut(clientnr, year, map_doel);
                }
                if (element.Contains(clientnr + @"\Correspondentie\" + year))
                {
                    map_doel = element;

                    UrlShortcut(clientnr, year, map_doel);
                }
            }
            map_doel = sPathFile;
            Item_naam = Item.Subject;
            Item_naam = RemoveIllegalFileNameChars(Item_naam);
            if (Item_naam.Length > 100)
            {
                Item_naam = Item_naam.Substring(0, 100);

            }
            if (clientnr == "99999999999999999999999999")
            {
                Item_naam = "mail adress" + " " + RemoveIllegalFileNameChars(mail);
                if (Item_naam.Length > 100)
                {
                    Item_naam = Item_naam.Substring(0, 100);
                }
                date = "";
                Item.Categories = "Onbekend mail adress";
            }
            else
            {
                Item.Categories = "Automatisch gearchiveerd";
            }

            if (!System.IO.File.Exists(map_doel + PATHSEPARATOR + date + " " + Item_naam + MAIL_SAVEFILE_EXT))
                Item.SaveAs(map_doel + PATHSEPARATOR + date + " " + Item_naam + MAIL_SAVEFILE_EXT, MAIL_SAVETYPE);
            Item.Save();



        }
        public void Set_clientnrList()
        {
            StreamReader reader = new StreamReader(System.IO.File.OpenRead(@"D:\Data\MailList.txt"));
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (!String.IsNullOrWhiteSpace(line))
                {
                    string[] values = line.Split(';');
                    if (values.Length >= 1)
                    {
                        maiList.Add(new KeyValuePair<string, string>(values[1], values[0]));
                    }
                }
            }
        }
        public void Get_clientnr2(string mail, List<string> mappen, MailItem Item, List<string> mappen2)
        {
            bool found = false;
            foreach (KeyValuePair<string, string> item in maiList)
            {
                if (item.Key == mail)
                {
                    Save_item(mappen, Item, item.Value, mail, mappen2);
                    found = true;
                }
            }
            if (!found)
            {
                string clnr = "99999999999999999999999999";
                Save_item(mappen, Item, clnr, mail, mappen2);
            }
        }
        private void AddACategory()
        {
            Outlook.Categories categories =
                Application.Session.Categories;
            if (!CategoryExists("Onbekend mail adress"))
            {
                Outlook.Category category = categories.Add("Onbekend mail adress",
                    Outlook.OlCategoryColor.olCategoryColorDarkRed);
            }
            if (!CategoryExists("none"))
            {
                Outlook.Category category = categories.Add("none",
                    Outlook.OlCategoryColor.olCategoryColorNone);
            }
            if (!CategoryExists("Automatisch gearchiveerd"))
            {
                Outlook.Category category = categories.Add("Automatisch gearchiveerd",
                    Outlook.OlCategoryColor.olCategoryColorGreen);
            }
        }

        private bool CategoryExists(string categoryName)
        {
            try
            {
                Outlook.Category category =
                    Application.Session.Categories[categoryName];
                if (category != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
        }

    }
}
