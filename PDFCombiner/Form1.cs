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
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

namespace PDFCombiner
{
    public partial class Form1 : Form
    {
        private bool has457 = true;
        private bool cwdLoaded = false;
        private bool IRSLoaded = false;
        private bool Plan403Loaded = false;
        private bool AA403Loaded = false;
        private bool Plan457Loaded = false;
        private bool AA457Loaded = false;
        private bool PALoaded = false;
        private bool AddALoaded = false;
        private bool MultiLoaded = false;
        private bool AddBLoaded = false;
        private bool AddCLoaded = false;
        private bool AddCALoaded = false;
        private bool TALoaded = false;
        private bool X1000Loaded = false;
        private string cwd = "C:\\";

        public Form1()
        {
            InitializeComponent();
        }


        private void cwdBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = "T:\\New Plan Document Roll Out\\Plan Document roll out\\School Districts";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                cwd = fbd.SelectedPath;
                cwdLbl.Text = cwd;
                IRSlbl.Text =FindPDFfiles("IRS", "");
                plan403Lbl.Text = FindPDFfiles("403*plan", "");
                aa403Lbl.Text = FindPDFfiles("403*AA", "");

                if(cb457.Checked == true)
                {
                    plan457Lbl.Text = FindPDFfiles("457*plan", "");
                    aa457Lbl.Text = FindPDFfiles("457*AA", "457*Adopt");
                    plan457Btn.Visible = true;
                    plan457Lbl.Visible = true;
                    aa457Btn.Visible = true;
                    aa457Lbl.Visible = true;
                } else
                {
                    plan457Lbl.Text = "No 457 Plan docs";
                    plan457Lbl.ForeColor = System.Drawing.Color.Red;
                    aa457Lbl.Text = "No 457 Plan docs";
                    aa457Lbl.ForeColor = System.Drawing.Color.Red;

                    plan457Lbl.Visible = true;
                    aa457Lbl.Visible = true;
                }

                paLbl.Text =  FindPDFfiles("PA ","");
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

        private void IRSbtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult result = fd.ShowDialog();
            if(result == DialogResult.OK)
            {
                IRSlbl.Text = "IRS " + fd.FileName;
            }
        }

        private string FindPDFfiles(string name, string altName)
        {
            string oldest = "";
            List<string> found = new List<string>();
            string[] dirs = Directory.GetDirectories(cwd);
            foreach (string dir in dirs)
            {
                string[] files = Directory.GetFiles(dir,"*" + name + "*.pdf");

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

        private void makeBtn_Click(object sender, EventArgs e)
        {
            using (PdfDocument irsDoc = PdfReader.Open(IRSlbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument plan403doc = PdfReader.Open(plan403Lbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument aa403doc = PdfReader.Open(aa403Lbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument paDoc = PdfReader.Open(paLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument addADoc = PdfReader.Open(addALbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument multiDoc = PdfReader.Open(multiLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument addBDoc = PdfReader.Open(addBLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument addCDoc = PdfReader.Open(addCLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument addCADoc = PdfReader.Open(AddCALbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument taDoc = PdfReader.Open(taLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument xeDoc = PdfReader.Open(xeLbl.Text, PdfDocumentOpenMode.Import))
            using (PdfDocument comDoc = new PdfDocument())
            {
                CopyPages(irsDoc, comDoc, "IRS Determination Letter");
                CopyPages(plan403doc, comDoc, "403b Plan Document");
                CopyPages(aa403doc, comDoc, "403b Adoption Agreement");
                CopyPages(paDoc, comDoc, "403_457 PA Agreement");
                CopyPages(addADoc, comDoc, "Addendum A");
                CopyPages(multiDoc, comDoc, "Multipurpose Employer Agreement");
                CopyPages(addBDoc, comDoc, "Addendum B");
                CopyPages(addCDoc, comDoc, "Addendum C");
                CopyPages(addCADoc, comDoc, "Addendum C_Exhibit A");
                CopyPages(taDoc, comDoc, "TA Application");
                CopyPages(xeDoc, comDoc, "XE100100 - School Districts Endorsement");

                comDoc.Save(cwd + "\\Combined.pdf");
                comDoc.Close();
            }
            openBtn.Visible = true;
            openBtn.Text = "Click here to open file";
        }

        private void CopyPages(PdfDocument from, PdfDocument to, String name)
        {
            Console.Error.Write("test");
            int pdfSIZE = from.PageCount;
            PdfPage page;
            page = from.Pages[0];
            to.AddPage(page);
            to.Outlines.Add(name, to.Pages[to.PageCount - 1], true, PdfOutlineStyle.Regular, XColors.Black);
            for (int i = 1; i < pdfSIZE; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }

        private void openBtn_Click(object sender, EventArgs e)
        {
            //File.Open(cwd + "\\Combined.pdf", FileMode.Open);
            System.Diagnostics.Process.Start(cwd + "\\Combined.pdf");
        }
    }
}
