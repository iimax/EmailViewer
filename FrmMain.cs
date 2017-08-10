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

namespace EmailViewer
{
    public partial class FrmMain : Form
    {
        private string MsgFilePath = null;
        private string MailBody = null;
        private List<string> Attachments = new List<string>();
        public FrmMain(string path)
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

        private void FrmMain_Load(object sender, EventArgs e)
        {
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
    }
}
