using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Collections;
using System.Net.Mail;
using System.Net;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security.Permissions;
using Infobox = Microsoft.VisualBasic.Interaction;

namespace PDFCombiner
{
    public partial class Form1 : Form
    {
        private string saveLoc = "";
        private static List<Tuple<string, string>> pdfs = null;
        private string cwd = "C:\\";

        public Form1()
        {
            InitializeComponent();
        }

        static void OnProcessExit(object sender, EventArgs e)
        {
            Console.WriteLine("I'm out of here");
            System.Environment.Exit(1);
        }

        private void cwdBtn_Click(object sender, EventArgs e)
        {
            pdfs = new List<Tuple<string, string>>();
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = @"T:\New Plan Document Roll Out\School Districts";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                cwd = fbd.SelectedPath;
                cwdLbl.Text = cwd;
                IRSlbl.Text = FindPDFfiles("IRS", "");
                plan403Lbl.Text = FindPDFfiles("403*plan", "");
                aa403Lbl.Text = FindPDFfiles("403*AA", "");

                if (cb457.Checked == true)
                {
                    plan457Lbl.Text = FindPDFfiles("457*plan", "");
                    aa457Lbl.Text = FindPDFfiles("457*AA", "457*Adopt");

                    plan457Btn.Visible = true;
                    plan457Lbl.Visible = true;
                    aa457Btn.Visible = true;
                    aa457Lbl.Visible = true;
                }
                else
                {
                    plan457Lbl.Text = "No 457 Plan docs";
                    plan457Lbl.ForeColor = System.Drawing.Color.Red;
                    aa457Lbl.Text = "No 457 Plan docs";
                    aa457Lbl.ForeColor = System.Drawing.Color.Red;

                    plan457Lbl.Visible = true;
                    aa457Lbl.Visible = true;
                }

                paLbl.Text = FindPDFfiles("PA ", "");
                addALbl.Text = FindPDFfiles("ADDENDUM A", "");
                multiLbl.Text = FindPDFfiles("Multi", "");
                addBLbl.Text = FindPDFfiles("ADDENDUM B", "");
                addCLbl.Text = FindPDFfiles("Addendum C ", "");
                AddCALbl.Text = FindPDFfiles("EXIHIBIT", "");
                taLbl.Text = FindPDFfiles("TA ", "");
                xeLbl.Text = FindPDFfiles("XE100", "");

                plan403Btn.Visible = true;
                plan403Lbl.Visible = true;
                aa403Btn.Visible = true;
                aa403Lbl.Visible = true;
                IRSbtn.Visible = true;
                IRSlbl.Visible = true;
                paBtn.Visible = true;
                paLbl.Visible = true;
                addABtn.Visible = true;
                addALbl.Visible = true;
                multiLbl.Visible = true;
                multiBtn.Visible = true;
                addBBtn.Visible = true;
                addBLbl.Visible = true;
                addCBtn.Visible = true;
                addCLbl.Visible = true;
                addCABtn.Visible = true;
                AddCALbl.Visible = true;
                taBtn.Visible = true;
                taLbl.Visible = true;
                xeBtn.Visible = true;
                xeLbl.Visible = true;

                makeBtn.Visible = true;
                makeBtn.Text = "Press to build files in: " + cwd;
            }
        }

        private void HideButtons()
        {
            cwdBtn.Hide();
            plan403Btn.Hide();
            aa403Btn.Hide();
            IRSbtn.Hide();
            paBtn.Hide();
            addABtn.Hide();
            multiBtn.Hide();
            addBBtn.Hide();
            addCBtn.Hide();
            addCABtn.Hide();
            taBtn.Hide();
            xeBtn.Hide();
            plan457Btn.Hide();
            aa457Btn.Hide();
        }

        private void LoadFiles()
        {
            pdfs.Add(Tuple.Create("IRS Determination Letter", IRSlbl.Text));
            pdfs.Add(Tuple.Create("403b Plan Document", plan403Lbl.Text));
            pdfs.Add(Tuple.Create("403b Adoption Agreement", aa403Lbl.Text));

            if (cb457.Checked == true)
            {
                pdfs.Add(Tuple.Create("457 Plan Document", plan457Lbl.Text));
                pdfs.Add(Tuple.Create("457 Adoption Agreement", aa457Lbl.Text));
            }

            pdfs.Add(Tuple.Create("403_457 PA Agreement", paLbl.Text));
            pdfs.Add(Tuple.Create("Addendum A", addALbl.Text));
            pdfs.Add(Tuple.Create("Multipurpose Employer Agreement", multiLbl.Text));
            pdfs.Add(Tuple.Create("Addendum B", addBLbl.Text));
            pdfs.Add(Tuple.Create("Addendum C", addCLbl.Text));
            pdfs.Add(Tuple.Create("Addendum C_Exhibit A", AddCALbl.Text));
            pdfs.Add(Tuple.Create("TA Application", taLbl.Text));
            pdfs.Add(Tuple.Create("XE100100 - School Districts Endorsement", xeLbl.Text));
        }

        private void IRSbtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                IRSlbl.Text = fd.FileName;
            }
        }

        private string FindPDFfiles(string name, string altName)
        {
            string oldest = "";
            List<string> found = new List<string>();
            string[] dirs = Directory.GetDirectories(cwd);
            foreach (string dir in dirs)
            {
                string[] files = Directory.GetFiles(dir, "*" + name + "*.pdf");

                if (files.Length == 0 && altName != "")//if no files are found using primary, search for alt if not blank
                {
                    files = Directory.GetFiles(dir, "*" + altName + "*.pdf");
                }

                if (files.Length != 0)
                {
                    DateTime dt = File.GetLastWriteTime(files[0]);

                    for (int i = 0; i < files.Length; i++)
                    {
                        DateTime temp = File.GetLastWriteTime(files[i]);
                        if (temp >= dt)
                        {
                            dt = temp;
                            oldest = files[i];
                        }
                    }
                }
            }
            return oldest;
        }

        public void Merge(List<Tuple<string, string>> InFiles, String OutFile)
        {
            var bookmarks = new List<Dictionary<string, object>>();
            using (FileStream stream = new FileStream(OutFile, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();
                PdfStamper stamper = null;
                PdfReader reader = null;
                PdfImportedPage page = null;

                int pageCntr = 1;
                InFiles.ForEach(file =>
                {
                    string newName = "";
                    reader = new PdfReader(file.Item2);
                    stamper = new PdfStamper(reader, stream);
                    AcroFields acroFields = stamper.AcroFields;

                    if (acroFields != null && acroFields.GenerateAppearances != true)
                    {
                        acroFields.GenerateAppearances = true;
                    }

                    IDictionary<string, AcroFields.Item> map = acroFields.Fields;
                    Random rand = new Random();
                    List<string> oldNames = new List<string>();

                    foreach (KeyValuePair<string, AcroFields.Item> entry in map)
                    {
                        oldNames.Add(entry.Key);
                    }

                    foreach (string oldName in oldNames)
                    {
                        newName = oldName + "_" + rand.Next(10000, 100000);
                        acroFields.RenameField(oldName, newName);
                        Console.WriteLine(newName);
                    }

                    for (int i = 0; i < reader.NumberOfPages; i++)
                    {
                        page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                        var h = page.Height;
                        if (i == 0)
                        {
                            var mark = new Dictionary<string, object>();
                            mark.Add("Action", "GoTo");
                            mark.Add("Title", file.Item1);
                            mark.Add("Page", pageCntr + " XYZ 0 " + h + " 0");
                            bookmarks.Add(mark);
                        }
                        pageCntr++;
                    }

                    pdf.FreeReader(reader);
                    reader.Close();
                });
                pdf.Outlines = bookmarks;
            }
        }

        private void FixAndRename(string src)
        {
            PdfReader reader = new PdfReader(src);
            PdfDictionary root = reader.Catalog;
            PdfDictionary form = root.GetAsDict(PdfName.ACROFORM);
            PdfArray fields = form.GetAsArray(PdfName.FIELDS);
            PdfDictionary page;
            PdfArray annots;
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                page = reader.GetPageN(i);
                annots = page.GetAsArray(PdfName.ANNOTS);
                for (int j = 0; j < annots.Size; j++)
                {
                    fields.Add(annots.GetAsIndirectObject(j));
                }
            }
            PdfStamper stamper = new PdfStamper(reader, new FileStream(src, FileMode.Create));
            stamper.Close();
            reader.Close();
        }



        private void makeBtn_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Finding PDFs...";
            LoadFiles();
            lblStatus.Text = "Picking Directory...";
            saveLoc = SaveFile();
            if (saveLoc != "" && saveLoc != null)
            {
                lblStatus.Text = "Building PDF...";
                Merge(pdfs, saveLoc);
                lblStatus.Text = "Done building PDF...";
            } else
            {
                lblStatus.Text = "No save location has been selected";
                return;
            }

            HideButtons();
            DialogResult result = MessageBox.Show(this, 
                "Would you like to open and view the document?", "Open File?", MessageBoxButtons.YesNo);
            if(result == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(saveLoc);
            }
            
            result = MessageBox.Show(this,
                "Would you to email the document?", "EMail File?", MessageBoxButtons.YesNo);

            if(result == DialogResult.Yes)
            {
                lblStatus.Text = "Mailing PDF...";
                SendFile();
                lblStatus.Text = "PDF Mailed";
            }

            result = MessageBox.Show(this,
                "Start a new file", "New File?", MessageBoxButtons.YesNo);

            if(result == DialogResult.Yes)
            {
                this.Hide();
                var form2 = new Form1();
                form2.Closed += (s, args) => this.Close();
                form2.Show();
            }
        }

        private string SaveFile()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = cwd;
            sfd.FileName = "Combined.pdf";
            sfd.Title = "Save the combined documents as...";
            sfd.Filter = "PDF files | *.pdf";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine(sfd.FileName);
                return sfd.FileName;
            }
            else return "";

        }

        private void plan403Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                plan403Lbl.Text = fd.FileName;
            }
        }

        private void aa403Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                aa403Lbl.Text = fd.FileName;
            }
        }

        private void plan457Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                plan457Lbl.Text = fd.FileName;
            }
        }

        private void aa457Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                aa457Lbl.Text = fd.FileName;
            }
        }

        private void paBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                paLbl.Text = fd.FileName;
            }
        }

        private void addABtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addALbl.Text = fd.FileName;
            }
        }

        private void multiBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                multiLbl.Text = fd.FileName;
            }
        }

        private void addBBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addBLbl.Text = fd.FileName;
            }
        }

        private void addCBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addCLbl.Text = fd.FileName;
            }
        }

        private void addCABtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                AddCALbl.Text = fd.FileName;
            }
        }

        private void taBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                taLbl.Text = fd.FileName;
            }
        }

        private void xeBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                xeLbl.Text = fd.FileName;
            }
        }

         private void SendFile()
        {

            string toList = "cgoldman@tdsgroup.org ; nbillings@tdsgroup.org";
            string ccList = "ccolton@tdsgroup.org; jtimmerman@tdsgroup.org; rhofhine@tdsgroup.org";
            string testList = "jchavis@tdsgroup.org ; jchavis@ralotter.com";

            string fileName = Path.GetFileName(saveLoc);
            Outlook.Application outlookApp = new Outlook.Application();
            string school = Infobox.InputBox("What is the school name", "Enter School Name", "", 100, 100);
            Outlook.MailItem mail = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = "Sending combined documents, " + fileName + " " + school;
            Outlook.AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;

            if(currentUser.Type == "EX")
            {
                Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
                mail.Body = "Please see the attached document for " + school;
                mail.To = toList;
                mail.BCC = ccList;
                mail.Attachments.Add(saveLoc, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                mail.Send();

            }
        }

    }
}
