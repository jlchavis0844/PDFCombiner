using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Infobox = Microsoft.VisualBasic.Interaction;

namespace PDFCombiner
{
    /// <summary>
    /// Form window class
    /// </summary>
    public partial class Form1 : Form
    {
        private string saveLoc = "";//save location of the output file
        private static List<Tuple<string, string>> pdfs = null;//list of different pdf files
        private string cwd = "C:\\";//default directory

        /// <summary>
        /// Build the window
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// shutdown behavior of the application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void OnProcessExit(object sender, EventArgs e)
        {
            Console.WriteLine("I'm out of here");
            System.Environment.Exit(1);
        }

        /// <summary>
        /// loads the directory to check for the  various types of pdf files to be combined
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cwdBtn_Click(object sender, EventArgs e)
        {
            pdfs = new List<Tuple<string, string>>();//holds tuple of "Name of Document", "C:\\file\location\here.pdf
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = @"T:\New Plan Document Roll Out\School Districts";//default path to the file storage
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                cwd = fbd.SelectedPath;
                cwdLbl.Text = cwd;

                //begin looking for the files
                IRSlbl.Text = FindPDFfiles("IRS", "");
                plan403Lbl.Text = FindPDFfiles("403*plan", "");
                aa403Lbl.Text = FindPDFfiles("403*AA", "");

                //if a 457 plan exists
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

                //write the label
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

                //turn on make button
                makeBtn.Visible = true;
                makeBtn.Text = "Press to build files in: " + cwd;
            }
        }

        /// <summary>
        /// Hides all the file location buttons for the various pdf files
        /// </summary>
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

        /// <summary>
        /// builds the tuples and puts them in the pdf's list. File locations taken from label text
        /// </summary>
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

        /// <summary>
        /// pick the IRS letter file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IRSbtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                IRSlbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Searches the current working diectory (cwd) for a pdf file with the given name
        /// </summary>
        /// <param name="name">The string to search the file name for</param>
        /// <param name="altName">The 2nd string to search the file for</param>
        /// <returns>string of the file location for the newest matching files</returns>
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

        /// <summary>
        /// Runs the process to merge all the pdf files.
        /// </summary>
        /// <param name="InFiles"></param>
        /// <param name="OutFile"></param>
        public void Merge(List<Tuple<string, string>> InFiles, String OutFile)
        {
            var bookmarks = new List<Dictionary<string, object>>();//holds the bookmars
            using (FileStream stream = new FileStream(OutFile, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();
                PdfStamper stamper = null;
                PdfReader reader = null;
                PdfImportedPage page = null;

                //merge the files
                int pageCntr = 1;
                int pageTotals = 0;

                pbFiles.Visible = true;
                pbFiles.Minimum = 0;
                pbFiles.Maximum = InFiles.Count;
                pbFiles.Value = 1;
                pbFiles.Step = 1;
                pbPages.Visible = true;
                pbPages.Minimum = 0;
                pbPages.Step = 1;

                InFiles.ForEach(file =>
                {
                    string newName = "";
                    reader = new PdfReader(file.Item2);
                    stamper = new PdfStamper(reader, stream);
                    AcroFields acroFields = stamper.AcroFields;

                    //turns on all form fields. Some are missing without the?
                    if (acroFields != null && acroFields.GenerateAppearances != true)
                    {
                        acroFields.GenerateAppearances = true;
                    }

                    IDictionary<string, AcroFields.Item> map = acroFields.Fields;
                    Random rand = new Random();
                    List<string> oldNames = new List<string>();

                    //load names for fields
                    foreach (KeyValuePair<string, AcroFields.Item> entry in map)
                    {
                        oldNames.Add(entry.Key);
                    }

                    //rename all fields to unique names
                    foreach (string oldName in oldNames)
                    {
                        newName = oldName + "_" + rand.Next(10000, 100000);
                        acroFields.RenameField(oldName, newName);
                        Console.WriteLine(newName);
                    }
                    pbPages.Maximum = reader.NumberOfPages;
                    pbPages.Value = 1;
                    //insert page by page
                    for (int i = 0; i < reader.NumberOfPages; i++)
                    {
                        page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                        var h = page.Height;

                        //on the first page, add a bookmark
                        if (i == 0)
                        {
                            var mark = new Dictionary<string, object>();
                            mark.Add("Action", "GoTo");
                            mark.Add("Title", file.Item1);
                            mark.Add("Page", pageCntr + " XYZ 0 " + h + " 0");
                            bookmarks.Add(mark);
                        }
                        pageCntr++;
                        pbPages.PerformStep();
                        pbPages.Refresh();
                    }
                    pbFiles.PerformStep();
                    pbFiles.Refresh();
                    pdf.FreeReader(reader);
                    reader.Close();
                });
                pdf.Outlines = bookmarks;//assign bookmarks
            }
        }

        /// <summary>
        /// drives the creation of the PDF files. Save location, Save, Open, Email, Restart
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void makeBtn_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Finding PDFs...";
            LoadFiles();
            lblStatus.Text = "Picking Directory...";
            saveLoc = SaveFile(); //get save location

            if (saveLoc != "" && saveLoc != null)
            {
                lblStatus.Text = "Building PDF...";
                lblStatus.Refresh();
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
            //if not open, then continue on
            if (result == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(saveLoc);
            }
            
            result = MessageBox.Show(this,
                "Would you to email the document?", "EMail File?", MessageBoxButtons.YesNo);
            //if not email, continue
            if(result == DialogResult.Yes)
            {
                lblStatus.Text = "Mailing PDF...";
                SendFile();
                lblStatus.Text = "PDF Mailed";
            }

            result = MessageBox.Show(this,
                "Start a new file", "New File?", MessageBoxButtons.YesNo);
            //if not start, do nothing
            ///TODO: quit the program if they don't refresh?
            if(result == DialogResult.Yes)
            {
                this.Hide();
                var form2 = new Form1();
                form2.Closed += (s, args) => this.Close();//subscribe to exit command listener
                form2.Show();
            }
        }

        /// <summary>
        /// Called to get the save location of the combined PDF and returns a path to that file
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void plan403Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                plan403Lbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aa403Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                aa403Lbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void plan457Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                plan457Lbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aa457Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                aa457Lbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void paBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                paLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addABtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addALbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void multiBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                multiLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addBBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addBLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addCBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                addCLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addCABtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                AddCALbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void taBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                taLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// Locates the PDF file to match the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void xeBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                xeLbl.Text = fd.FileName;
            }
        }

        /// <summary>
        /// emails the combined PDF file to other. Default to and cc lists are set but may be changed at mailing.
        /// </summary>
        private void SendFile()
        {
            //default lists
            string toList = "cgoldman@tdsgroup.org ; nbillings@tdsgroup.org";
            string ccList = "ccolton@tdsgroup.org; jtimmerman@tdsgroup.org; rhofhine@tdsgroup.org";
            //string testList = "jchavis@tdsgroup.org ; jchavis@ralotter.com";

            string fileName = Path.GetFileName(saveLoc);
            Outlook.Application outlookApp = new Outlook.Application();
            string school = Infobox.InputBox("What is the school name", "Enter School Name", "", 100, 100);
            toList = Infobox.InputBox("Email To (Seperate with ;):", "To Addresses", toList, 100, 100);
            ccList = Infobox.InputBox("Email CC (Seperate with ;):", "CC Addresses", ccList, 100, 100);

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
                try {
                    mail.Send();
                } catch(Exception e) {
                    MessageBox.Show("Something went wrong trying to email\n" + e.ToString());
                }
            }
        }
    }
}
