using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Nawras_Pages_Merge
{
    public partial class Form1 : Form
    {
        private ResourceManager nRM;   
        private CultureInfo nCI;

        private static string FRENCH = "fr";
        private static string ENGLISH = "en";
        private static string ARABIC = "ar";

        private string e401 = "Merci de choisir les pages impaires, les pages paires ainsi que le fichier résultat !";
        private string e404 = "Impossible d'accéder aux pages impaires et paires choisies !";
        private string e300 = "Le nombre des pages impaires et paires ne correspondent pas !";
        private string e200 = "Fusion terminée avec succès !";

        private int rotateLeft = 0;
        private int rotateRight = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            nRM = new ResourceManager("Nawras_Pages_Merge.Lang.Res", typeof(Form1).Assembly);
            françaisfrToolStripMenuItem.Checked = true;
            pbStatusBar.Width = statusBar.Width - 20;
            pbStatusBar.Visible = false;
        }

        private void imgI_Click(object sender, EventArgs e)
        {
            nOFD.Title = "Nawras Pages Merge";
            nOFD.Filter = "PDF Files|*.pdf";
            DialogResult nResult = nOFD.ShowDialog();
            if (nResult == DialogResult.OK)
            {
                string nPDF = nOFD.FileName;
                txtSourceI.Text = nPDF;
            }
        }

        private void imgP_Click(object sender, EventArgs e)
        {
            nOFD.Title = "Nawras Pages Merge";
            nOFD.Filter = "PDF Files|*.pdf";
            DialogResult nResult = nOFD.ShowDialog();
            if (nResult == DialogResult.OK)
            {
                string nPDF = nOFD.FileName;
                txtSourceP.Text = nPDF;
            }
        }

        private void imgSave_Click(object sender, EventArgs e)
        {
            nSFD.Title = "Nawras Pages Merge";
            nSFD.Filter = "PDF Files|*.pdf";
            DialogResult nResult = nSFD.ShowDialog();
            if (nResult == DialogResult.OK)
            {
                string nPDF = nSFD.FileName;
                txtOutput.Text = nPDF;
            }
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            pbStatusBar.Visible = true;
            allStatus(false);
            string nSourceI = txtSourceI.Text;
            string nSourceP = txtSourceP.Text;
            string nOutput = txtOutput.Text;
            if (string.IsNullOrWhiteSpace(nSourceI) || string.IsNullOrWhiteSpace(nSourceP) || string.IsNullOrWhiteSpace(nOutput))
            {
                MessageBox.Show(e401, "Nawras Pages Merge");
            }
            else
            {
                if (File.Exists(nSourceI) && File.Exists(nSourceP))
                {
                    try
                    {
                        PdfReader nI = new PdfReader(nSourceI);
                        PdfReader nP = new PdfReader(nSourceP);
                        
                        if (nI.NumberOfPages != nP.NumberOfPages)
                        {
                            MessageBox.Show(e300, "Nawras Pages Merge");
                        }
                        else
                        {
                            pbStatusBar.Maximum = nI.NumberOfPages * 2;
                            Document nSave = new Document(nI.GetPageSizeWithRotation(nI.NumberOfPages));
                            PdfCopy nCopy = new PdfCopy(nSave, new FileStream(nOutput, FileMode.Create));

                            //Add PDF metadata
                            nSave.AddAuthor(txtAuthor.Text);
                            nSave.AddCreator("Nawras Pages Merge");                            
                            nSave.AddKeywords(txtKeywords.Text);
                            nSave.AddSubject(txtSubject.Text);
                            nSave.AddTitle(txtTitle.Text);
                            PdfImportedPage nPageI, nPageP;
                            PdfDictionary nPageIp, nPagePp;
                            nSave.Open();
                            
                            //Apply pages rotation
                            for (int i = 1; i <= nI.NumberOfPages; i++)
                            {
                                nPageIp = nI.GetPageN(i);
                                nPagePp = nP.GetPageN(i);
                                nPageIp.Put(PdfName.ROTATE, new PdfNumber(rotateLeft));
                                nPagePp.Put(PdfName.ROTATE, new PdfNumber(rotateRight));
                            }
                            //Apply pages merging
                            for (int i = 1; i <= nI.NumberOfPages; i++)
                            {
                                nPageI = nCopy.GetImportedPage(nI, i);
                                pbStatusBar.Value += 1;
                                nPageP = nCopy.GetImportedPage(nP, i);
                                pbStatusBar.Value += 1;
                                nCopy.AddPage(nPageI);
                                nCopy.AddPage(nPageP);
                            }

                            nSave.Close();
                            nI.Close();
                            nP.Close();
                            MessageBox.Show(e200, "Nawras Pages Merge");
                            pbStatusBar.Value = 0;
                        }
                    }
                    catch (Exception nErr)
                    {
                        MessageBox.Show("Erreur : " + nErr.Message);
                    }
                }
                else
                {
                    MessageBox.Show(e404, "Nawras Pages Merge");
                }
            }
            allStatus(true);
            pbStatusBar.Visible = false;
            Cursor.Current = Cursors.Default;
        }

        private void allStatus(Boolean status)
        {
            imgI.Enabled = status;
            imgP.Enabled = status;
            imgSave.Enabled = status;
            txtSourceI.Enabled = status;
            txtSourceP.Enabled = status;
            txtOutput.Enabled = status;
            btnGo.Enabled = status;
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            allClear();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.nawrasg.fr/html/app/pages_merge");
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            new Form2().ShowDialog();
        }

        private void françaisfrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setLangage(FRENCH);
            englishenToolStripMenuItem.Checked = false;
            françaisfrToolStripMenuItem.Checked = true;
            arabicarToolStripMenuItem.Checked = false;
        }

        private void setLangage(string lang)
        {
            nCI = CultureInfo.CreateSpecificCulture(lang);
            tabMain.Text = nRM.GetString("home", nCI);
            tabRotation.Text = nRM.GetString("rotation", nCI);
            tabProperties.Text = nRM.GetString("description", nCI);
            groupBox1.Text = nRM.GetString("pi", nCI);
            groupBox2.Text = nRM.GetString("pp", nCI);
            lblI.Text = nRM.GetString("pi", nCI) + " :";
            lblP.Text = nRM.GetString("pp", nCI) + " :";
            lblOutput.Text = nRM.GetString("output", nCI) + " :";
            lblTitle.Text = nRM.GetString("title", nCI);
            lblAuthor.Text = nRM.GetString("author", nCI);
            lblSubject.Text = nRM.GetString("subject", nCI);
            lblKeywords.Text = nRM.GetString("keywords", nCI);
            e401 = nRM.GetString("e401", nCI);
            e404 = nRM.GetString("e404", nCI);
            e300 = nRM.GetString("e300", nCI);
            e200 = nRM.GetString("e200", nCI);
        }

        private void englishenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setLangage(ENGLISH);
            englishenToolStripMenuItem.Checked = true;
            françaisfrToolStripMenuItem.Checked = false;
            arabicarToolStripMenuItem.Checked = false;
        }

        private void allClear()
        {
            txtOutput.Clear();
            txtSourceI.Clear();
            txtSourceP.Clear();
            txtTitle.Clear();
            txtAuthor.Clear();
            txtSubject.Clear();
            txtKeywords.Clear();
            rotateLeft = 0;
            rotateRight = 0;
            imgLeft.Image = Nawras_Pages_Merge.Properties.Resources.left128;
            imgRight.Image = Nawras_Pages_Merge.Properties.Resources.right128;
        }

        private void btnLeftI_Click(object sender, EventArgs e)
        {
            System.Drawing.Image nImage = imgLeft.Image;
            nImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
            imgLeft.Image = nImage;
            rotateLeft -= 90;
            if (Math.Abs(rotateLeft) == 360) rotateLeft = 0;
        }

        private void btnRightI_Click(object sender, EventArgs e)
        {
            System.Drawing.Image nImage = imgLeft.Image;
            nImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
            imgLeft.Image = nImage;
            rotateLeft += 90;
            if (Math.Abs(rotateLeft) == 360) rotateLeft = 0;
        }

        private void btnLeftP_Click(object sender, EventArgs e)
        {
            System.Drawing.Image nImage = imgRight.Image;
            nImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
            imgRight.Image = nImage;
            rotateRight -= 90;
            if (Math.Abs(rotateRight) == 360) rotateRight = 0;
        }

        private void btnRightP_Click(object sender, EventArgs e)
        {
            System.Drawing.Image nImage = imgRight.Image;
            nImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
            imgRight.Image = nImage;
            rotateRight += 90;
            if (Math.Abs(rotateRight) == 360) rotateRight = 0;
        }

        private void arabicarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setLangage(ARABIC);
            englishenToolStripMenuItem.Checked = false;
            françaisfrToolStripMenuItem.Checked = false;
            arabicarToolStripMenuItem.Checked = true;
        }

        private void txtSourceI_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] nFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (Path.GetExtension(nFiles[0]) == ".pdf")
                {
                    txtSourceI.Text = nFiles[0];
                }
            }
        }

        private void txtSourceI_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void txtSourceP_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] nFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (Path.GetExtension(nFiles[0]) == ".pdf")
                {
                    txtSourceP.Text = nFiles[0];
                }
            }
        }

    }
}
