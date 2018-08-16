using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailArchief
{
    public partial class settings_form : Form
    {

        public bool inbox = false;
        public bool sent_items = false;
        public string Folder1 = "";
        public string Folder2 = "";
        public string Folder3 = "";
        public string Folder4 = "";
        public string Folder5 = "";
        public string sF1 = "";
        public string sF2 = "";
        public string sF3 = "";
        public string sF4 = "";
        public string sF5 = "";

        public settings_form()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            inbox = Inbox.Checked;
            sent_items = Sent_Items.Checked;
            Folder1 = Fol1.Text;
            Folder2 = Fol2.Text;
            Folder3 = Fol3.Text;
            Folder4 = Fol4.Text;
            Folder5 = Fol5.Text;
            sF1 = Subf1.Text;
            sF2 = Subf2.Text;
            sF3 = Subf3.Text;
            sF4 = Subf4.Text;
            sF5 = Subf5.Text;
            Globals.ThisAddIn.Folder1 = Folder1;
            Globals.ThisAddIn.Folder2 = Folder2;
            Globals.ThisAddIn.Folder3 = Folder3;
            Globals.ThisAddIn.Folder4 = Folder4;
            Globals.ThisAddIn.Folder5 = Folder5;

            Globals.ThisAddIn.use_inbox = inbox;
            Globals.ThisAddIn.use_sent_items = sent_items;

            Globals.ThisAddIn.F1Inbox = F1Inbox.Checked;
            Globals.ThisAddIn.F2Inbox = F2Inbox.Checked;
            Globals.ThisAddIn.F3Inbox = F3Inbox.Checked;
            Globals.ThisAddIn.F4Inbox = F4Inbox.Checked;
            Globals.ThisAddIn.F5Inbox = F5Inbox.Checked;

            Globals.ThisAddIn.SubF1 = sF1;
            Globals.ThisAddIn.SubF2 = sF2;
            Globals.ThisAddIn.SubF3 = sF3;
            Globals.ThisAddIn.SubF4 = sF4;
            Globals.ThisAddIn.SubF5 = sF5;

            Globals.ThisAddIn.Test_4();

            this.Visible = false;
            this.Dispose();
        }

        private void btnXML_Click(object sender, EventArgs e)
        {
            inbox = Inbox.Checked;
            sent_items = Sent_Items.Checked;
            Folder1 = Fol1.Text;
            Folder2 = Fol2.Text;
            Folder3 = Fol3.Text;
            Folder4 = Fol4.Text;
            Folder5 = Fol5.Text;
            sF1 = Subf1.Text;
            sF2 = Subf2.Text;
            sF3 = Subf3.Text;
            sF4 = Subf4.Text;
            sF5 = Subf5.Text;
            Globals.ThisAddIn.Folder1 = Folder1;
            Globals.ThisAddIn.Folder2 = Folder2;
            Globals.ThisAddIn.Folder3 = Folder3;
            Globals.ThisAddIn.Folder4 = Folder4;
            Globals.ThisAddIn.Folder5 = Folder5;

            Globals.ThisAddIn.use_inbox = inbox;
            Globals.ThisAddIn.use_sent_items = sent_items;

            Globals.ThisAddIn.F1Inbox = F1Inbox.Checked;
            Globals.ThisAddIn.F2Inbox = F2Inbox.Checked;
            Globals.ThisAddIn.F3Inbox = F3Inbox.Checked;
            Globals.ThisAddIn.F4Inbox = F4Inbox.Checked;
            Globals.ThisAddIn.F5Inbox = F5Inbox.Checked;

            Globals.ThisAddIn.SubF1 = sF1;
            Globals.ThisAddIn.SubF2 = sF2;
            Globals.ThisAddIn.SubF3 = sF3;
            Globals.ThisAddIn.SubF4 = sF4;
            Globals.ThisAddIn.SubF5 = sF5;

            Globals.ThisAddIn.Test_4();

            this.Visible = false;
            this.Dispose();
        }

        private void btnXML_Click_1(object sender, EventArgs e)
        {
            inbox = Inbox.Checked;
            sent_items = Sent_Items.Checked;
            Folder1 = Fol1.Text;
            Folder2 = Fol2.Text;
            Folder3 = Fol3.Text;
            Folder4 = Fol4.Text;
            Folder5 = Fol5.Text;
            sF1 = Subf1.Text;
            sF2 = Subf2.Text;
            sF3 = Subf3.Text;
            sF4 = Subf4.Text;
            sF5 = Subf5.Text;
            Globals.ThisAddIn.Folder1 = Folder1;
            Globals.ThisAddIn.Folder2 = Folder2;
            Globals.ThisAddIn.Folder3 = Folder3;
            Globals.ThisAddIn.Folder4 = Folder4;
            Globals.ThisAddIn.Folder5 = Folder5;

            Globals.ThisAddIn.use_inbox = inbox;
            Globals.ThisAddIn.use_sent_items = sent_items;

            Globals.ThisAddIn.F1Inbox = F1Inbox.Checked;
            Globals.ThisAddIn.F2Inbox = F2Inbox.Checked;
            Globals.ThisAddIn.F3Inbox = F3Inbox.Checked;
            Globals.ThisAddIn.F4Inbox = F4Inbox.Checked;
            Globals.ThisAddIn.F5Inbox = F5Inbox.Checked;

            Globals.ThisAddIn.SubF1 = sF1;
            Globals.ThisAddIn.SubF2 = sF2;
            Globals.ThisAddIn.SubF3 = sF3;
            Globals.ThisAddIn.SubF4 = sF4;
            Globals.ThisAddIn.SubF5 = sF5;

            Globals.ThisAddIn.Test_4();

            this.Visible = false;
            this.Dispose();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            inbox = Inbox.Checked;
            sent_items = Sent_Items.Checked;
            Folder1 = Fol1.Text;
            Folder2 = Fol2.Text;
            Folder3 = Fol3.Text;
            Folder4 = Fol4.Text;
            Folder5 = Fol5.Text;
            sF1 = Subf1.Text;
            sF2 = Subf2.Text;
            sF3 = Subf3.Text;
            sF4 = Subf4.Text;
            sF5 = Subf5.Text;
            Globals.ThisAddIn.Folder1 = Folder1;
            Globals.ThisAddIn.Folder2 = Folder2;
            Globals.ThisAddIn.Folder3 = Folder3;
            Globals.ThisAddIn.Folder4 = Folder4;
            Globals.ThisAddIn.Folder5 = Folder5;

            Globals.ThisAddIn.use_inbox = inbox;
            Globals.ThisAddIn.use_sent_items = sent_items;

            Globals.ThisAddIn.F1Inbox = F1Inbox.Checked;
            Globals.ThisAddIn.F2Inbox = F2Inbox.Checked;
            Globals.ThisAddIn.F3Inbox = F3Inbox.Checked;
            Globals.ThisAddIn.F4Inbox = F4Inbox.Checked;
            Globals.ThisAddIn.F5Inbox = F5Inbox.Checked;

            Globals.ThisAddIn.SubF1 = sF1;
            Globals.ThisAddIn.SubF2 = sF2;
            Globals.ThisAddIn.SubF3 = sF3;
            Globals.ThisAddIn.SubF4 = sF4;
            Globals.ThisAddIn.SubF5 = sF5;

            Globals.ThisAddIn.Test_4();

            this.Visible = false;
            this.Dispose();
        }
    }
}
