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
using ScintillaNET;
using System.Xml;

namespace EmailViewer
{
    public partial class FrmMain2 : Form
    {
        private string MsgFilePath = null;
        private string MailBody = null;
        private List<string> Attachments = new List<string>();

        private QuoteproLexer qpLexer = new QuoteproLexer("Type Agent Code State Quoted Insurer Web Login Password ProducerCode Comment has been added as Autoquote Homequote");

        public FrmMain2(string path)
        {
            this.MsgFilePath = path;
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            DialogResult msgFileSelectResult = this.openFileDialog1.ShowDialog();
            if (msgFileSelectResult == DialogResult.OK)
            {
                string msgfile = this.openFileDialog1.FileName;
                ReadMsgFile(msgfile);
            }
        }

        private void ApplyStyling()
        {
            //txtMail.StartStyling(0);
            //txtMail.SetStyling(5, 1);
            //txtMail.SetStyling(2, 2);
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            txtMail.StyleResetDefault();
            txtMail.Styles[Style.Default].Font = "Consolas";
            txtMail.Styles[Style.Default].Size = 10;
            txtMail.StyleClearAll();

            txtMail.Styles[QuoteproLexer.StyleDefault].ForeColor = Color.Black;
            txtMail.Styles[QuoteproLexer.StyleKeyword].ForeColor = Color.Blue;
            txtMail.Styles[QuoteproLexer.StyleIdentifier].ForeColor = Color.Teal;
            txtMail.Styles[QuoteproLexer.StyleNumber].ForeColor = Color.Purple;
            txtMail.Styles[QuoteproLexer.StyleString].ForeColor = Color.Red;

            txtMail.Styles[QuoteproLexer.StyleAutoOrHome].ForeColor = Color.Black;
            txtMail.Styles[QuoteproLexer.StyleAutoOrHome].BackColor = Color.Yellow;

            txtMail.Lexer = Lexer.Container;

            //txtMail.Lexer = Lexer.Null;

            //txtMail.Styles[1].ForeColor = Color.Red;
            //txtMail.Styles[2].ForeColor = Color.Blue;
            //txtMail.SetKeywords(0, "Insurer Web Login:");

            if (!string.IsNullOrWhiteSpace(MsgFilePath))
            {
                if (System.IO.File.Exists(MsgFilePath))
                {
                    ReadMsgFile(MsgFilePath);
                }
                else
                {
                    MessageBox.Show("文件["+ MsgFilePath +"]不存在","Oops", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ReadMsgFile(string msgfile)
        {
            //清空附件内容列表，删除附件菜单项
            MailBody = string.Empty;
            Attachments.Clear();
            btnFiles.DropDownItems.Clear();
            lblStatus.Text = "-";

            FileInfo fi = new FileInfo(msgfile);
            this.Text = string.Format("EmailViewer of {0}", fi.Name);
            using (Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read))
            {
                //fileName = string.Empty;
                OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                messageStream.Close();
                try
                {
                    //获取comment#
                    //Comment has been added as: 34227
                    //string comment = null;
                    MailBody = message.BodyText;
                    this.txtMail.Text = MailBody;

                    //附件
                    if (message.Attachments.Count > 0)
                    {
                        int i = 0;
                        foreach (OutlookStorage.Attachment item in message.Attachments)
                        {
                            string fileContent = System.Text.Encoding.UTF8.GetString(item.Data);
                            Attachments.Add(fileContent);
                            ToolStripMenuItem ddItem = new ToolStripMenuItem(item.Filename);
                            ddItem.Name = "btnFile" + i.ToString();
                            ddItem.Width = 300;
                            ddItem.Click += ddItem_Click;
                            btnFiles.DropDownItems.Add(ddItem);
                        }
                        
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "不好,出了点儿问题", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    message.Dispose();
                }
            }
        }

        void ddItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ddItem = sender as ToolStripMenuItem;
            if (ddItem == null)
            {
                return;
            }
            string btnName = ddItem.Name;
            int idx = ddItem.Name.IndexOf("btnFile");
            if (idx==-1)
            {
                lblStatus.Text = "Error";
                return;
            }
            string AttachIndex = ddItem.Name.Substring(idx+7);
            int? index = ToInt(AttachIndex);
            if (index.HasValue)
            {
                this.txtMail.Text = Attachments[index.Value];
            }
            else
            {
                lblStatus.Text = "Error";
            }
        }

        private void txtMail_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void txtMail_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                string file = files[0];
                if (System.IO.File.Exists(file))
                {
                    MsgFilePath = file;
                    //目前仅支持 msg 文件
                    ReadMsgFile(MsgFilePath);
                }
            }
        }

        private void btnMailBody_Click(object sender, EventArgs e)
        {
            this.txtMail.Text = MailBody;
        }

        private int? ToInt(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                return null;
            }
            int i = -1;
            if (int.TryParse(str, out i))
            {
                return i;
            }
            return null;
        }
        //int needed = 1;
        private void txtMail_StyleNeeded(object sender, StyleNeededEventArgs e)
        {
            var startPos = txtMail.GetEndStyled();
            var endPos = e.Position;

            qpLexer.Style(txtMail, startPos, endPos);
            //lblStatus.Text = "needed" + needed.ToString();
            //needed++;
        }

        private void btnDump_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Multiselect = true;
            DialogResult msgFileSelectResult = this.openFileDialog1.ShowDialog();
            if (msgFileSelectResult == DialogResult.OK)
            {
                string targetFolder = this.openFileDialog1.FileNames[0];
                targetFolder = System.IO.Directory.GetParent(targetFolder).FullName;
                foreach (string msgfile in this.openFileDialog1.FileNames)
                {
                    bool attachGenerated = false;
                    string fileName = null;
                    using (Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read))
                    {
                        fileName = string.Empty;
                        OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                        messageStream.Close();
                        try
                        {
                            //获取comment#
                            //Comment has been added as: 34227
                            string comment = null;
                            string body = message.BodyText;
                            comment = body.Substring(body.IndexOf("Comment has been added as:") + 26).Trim();
                            //Insurer: Aggressive Insurance Texas Elite
                            string insurer = null;
                            string state = null;
                            string title = message.Subject;
                            if (title.Contains("Reminder Email for debugging purpose"))
                            {
                                //CarrierName: Progressive Auto
                                insurer = ExtractHtml(body, "CarrierName:", "\r\n");
                                if (string.IsNullOrEmpty(insurer))
                                {
                                    throw new Exception("Failed to extract CarrierName");
                                }
                                insurer = insurer.Trim().Split(' ')[0];
                                //Quoting State: VT
                                state = body.Substring(body.IndexOf("Quoting State:") + 14).Trim();
                                state = state.Substring(0, state.IndexOf("\r\n"));
                            }
                            else
                            {
                                //insurer = body.Substring(body.IndexOf("Insurer:")+8).Trim();
                                //insurer = insurer.Substring(0, insurer.IndexOf("\r\n"));
                                //insurer = insurer.Substring(0, insurer.IndexOf(" "));
                                //qpquote\stdInsurers\HartfordScrape.vb
                                insurer = body.Substring(body.LastIndexOf("qpquote\\stdInsurers\\") + 20);
                                if (insurer.Contains(".vb"))
                                {
                                    insurer = insurer.Substring(0, insurer.IndexOf(".vb"));
                                    if (insurer.Contains("Scrape"))
                                    {
                                        //insurer = insurer.Replace("Scrape", "");
                                        insurer = insurer.Substring(0, insurer.IndexOf("Scrape"));
                                    }
                                }
                                else
                                {
                                    //Insurer: Progressive Insurance Company
                                    insurer = ExtractHtml(insurer, "Insurer:", "Insurer Web Login:");
                                    if (string.IsNullOrEmpty(insurer))
                                    {
                                        throw new Exception("Failed to extract Insurer Company Name");
                                    }
                                    insurer = insurer.Trim().Split(' ')[0];
                                }
                                //State Quoted: MA

                                state = body.Substring(body.IndexOf("State Quoted:") + 13).Trim();
                                state = state.Substring(0, state.IndexOf("\r\n"));
                            }
                            //2017-08-10,iimax:state rule changed
                            if (state.StartsWith("State"))
                            {
                                state = state.Substring(5);
                            }
                            fileName = string.Format("#{0}-{1} ({2}).txt", comment, insurer, state);
                            //附件
                            if (message.Attachments.Count > 0)
                            {
                                OutlookStorage.Attachment file = message.Attachments[0];
                                string fileContent = System.Text.Encoding.UTF8.GetString(file.Data);
                                //2016-06-29：修改last name为 testing，修改 effdate
                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.PreserveWhitespace = true;
                                doc.LoadXml(fileContent);
                                if (fileContent.Contains("<AutoQuote>"))
                                {
                                    var CusLastName = doc.DocumentElement.SelectSingleNode("/AutoQuote/Quotation/CusLastName");
                                    if (CusLastName != null)
                                    {
                                        //Create a comment.
                                        XmlComment newComment = doc.CreateComment(CusLastName.OuterXml);
                                        CusLastName.ParentNode.InsertAfter(newComment, CusLastName);
                                        CusLastName.InnerText = "Testing";
                                    }
                                    var EffDate = doc.DocumentElement.SelectSingleNode("/AutoQuote/Quotation/EffDate");
                                    if (EffDate != null)
                                    {
                                        string eff = EffDate.InnerText;
                                        DateTime dt = DateTime.Today;
                                        DateTime.TryParse(eff, out dt);
                                        if (dt < DateTime.Today)
                                        {
                                            XmlComment newComment = doc.CreateComment(EffDate.OuterXml);
                                            EffDate.ParentNode.InsertAfter(newComment, EffDate);
                                            EffDate.InnerText = DateTime.Today.ToString("MM/dd/yyyy");
                                        }
                                    }
                                    var drivers = doc.DocumentElement.SelectNodes("/AutoQuote/Quotation/Drivers/Driver");
                                    if (drivers != null)
                                    {
                                        foreach (XmlNode item in drivers)
                                        {
                                            var LastName = item.SelectSingleNode("LastName");
                                            if (LastName != null)
                                            {
                                                XmlComment newComment = doc.CreateComment(LastName.OuterXml);
                                                LastName.ParentNode.InsertAfter(newComment, LastName);
                                                LastName.InnerText = "Testing";
                                            }
                                        }
                                    }
                                }
                                else if (fileContent.Contains("<HomeQuote>"))
                                {
                                    var CusLastName = doc.DocumentElement.SelectSingleNode("/HomeQuote/InsuredInfo/CusLastName");
                                    if (CusLastName != null)
                                    {
                                        //Create a comment.
                                        XmlComment newComment = doc.CreateComment(CusLastName.OuterXml);
                                        CusLastName.ParentNode.InsertAfter(newComment, CusLastName);
                                        CusLastName.InnerText = "Testing";
                                    }
                                    var hEffDate = doc.DocumentElement.SelectSingleNode("/HomeQuote/InsuredInfo/hEffDate");
                                    if (hEffDate != null)
                                    {
                                        string eff = hEffDate.InnerText;
                                        DateTime dt = DateTime.Today;
                                        if (DateTime.TryParse(eff, out dt))
                                        {

                                        }
                                        if (dt < DateTime.Today)
                                        {
                                            XmlComment newComment = doc.CreateComment(hEffDate.OuterXml);
                                            hEffDate.ParentNode.InsertAfter(newComment, hEffDate);
                                            hEffDate.InnerText = DateTime.Today.ToString("MM/dd/yyyy");

                                            var hExpDate = doc.DocumentElement.SelectSingleNode("/HomeQuote/InsuredInfo/hExpDate");
                                            if (hExpDate != null)
                                            {
                                                XmlComment hExpDateComment = doc.CreateComment(hExpDate.OuterXml);
                                                hExpDate.ParentNode.InsertAfter(hExpDateComment, hExpDate);
                                                hExpDate.InnerText = DateTime.Today.AddYears(1).ToString("MM/dd/yyyy");
                                            }
                                            //修改 RiskInfo
                                            var EffDate = doc.DocumentElement.SelectSingleNode("/HomeQuote/RiskInfo/ThisRisk/EffDate");
                                            var ExpDate = doc.DocumentElement.SelectSingleNode("/HomeQuote/RiskInfo/ThisRisk/ExpDate");
                                            if (EffDate != null)
                                            {
                                                XmlComment xComment = doc.CreateComment(EffDate.OuterXml);
                                                EffDate.ParentNode.InsertAfter(xComment, EffDate);
                                                EffDate.InnerText = DateTime.Today.ToString("MM/dd/yyyy");
                                            }

                                            if (ExpDate != null)
                                            {
                                                XmlComment xComment = doc.CreateComment(ExpDate.OuterXml);
                                                ExpDate.ParentNode.InsertAfter(xComment, ExpDate);
                                                ExpDate.InnerText = DateTime.Today.AddYears(1).ToString("MM/dd/yyyy");
                                            }
                                        }

                                    }
                                }
                                fileContent = doc.OuterXml;
                                System.IO.File.WriteAllText(Path.Combine(targetFolder, fileName), fileContent);
                            }
                            attachGenerated = true;
                            //MessageBox.Show(fileName);
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        finally
                        {
                            message.Dispose();
                        }

                        if (attachGenerated)
                        {
                            //重命名msg
                            FileInfo fi = new FileInfo(msgfile);
                            //目标文件名
                            string msgfileName = fi.Name;
                            string ourfileName = fi.Name.Replace(".msg", string.Format("{0}.msg", fileName));
                            ourfileName = ourfileName.Replace("comment#", "#").Replace(".txt.", ".");
                            System.IO.File.Move(msgfile, Path.Combine(fi.DirectoryName, ourfileName));
                        }
                    }


                }
            }
        }

        private string ExtractHtml(string rs, string prefix, string suffix)
        {
            if (!rs.Contains(prefix))
            {
                return string.Empty;
            }
            if (string.IsNullOrWhiteSpace(prefix))
            {
                throw new ArgumentNullException("prefix is required");
            }

            int start = rs.IndexOf(prefix);
            rs = rs.Substring(start + prefix.Length);
            //if (string.IsNullOrWhiteSpace(suffix))
            //{
            //    return rs;
            //}
            //if (!rs.Contains(suffix))
            //{
            //    return string.Empty;
            //}

            return rs.Substring(0, rs.IndexOf(suffix));
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            txtMail.InsertText(0, "//mailto:iimax@outlook.com" + Environment.NewLine);
        }
    }
}
