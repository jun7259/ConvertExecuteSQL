using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertExecuteSQL
{
    public partial class Form1 : Form
    {
        [DllImport("User32.dll")]
        protected static extern int
               SetClipboardViewer(int hWndNewViewer);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool
               ChangeClipboardChain(IntPtr hWndRemove,
                                    IntPtr hWndNewNext);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg,
                                             IntPtr wParam,
                                             IntPtr lParam);

        IntPtr nextClipboardViewer;
        public Form1()
        {
            InitializeComponent();
            nextClipboardViewer = (IntPtr)SetClipboardViewer((int)
                             this.Handle);

        }

        protected override void
          WndProc(ref System.Windows.Forms.Message m)
        {
            // defined in winuser.h
            const int WM_DRAWCLIPBOARD = 0x308;
            const int WM_CHANGECBCHAIN = 0x030D;

            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    DisplayClipboardData();
                    SendMessage(nextClipboardViewer, m.Msg, m.WParam,
                                m.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        SendMessage(nextClipboardViewer, m.Msg, m.WParam,
                                    m.LParam);
                    break;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        protected override void Dispose(bool disposing)
        {
            ChangeClipboardChain(this.Handle, nextClipboardViewer);

            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        void DisplayClipboardData()
        {
            try
            {
                IDataObject iData = new DataObject();
                iData = Clipboard.GetDataObject();

                if (iData.GetDataPresent(DataFormats.Text))
                {
                    var strSQL = (string)iData.GetData(DataFormats.Text);
                    if (!strSQL.Contains("exec sp_executesql") || strSQL.Trim() == "exec sp_executesql") return;
                    var newSql = ConvertSql(strSQL);
                    Clipboard.SetText(newSql);
                }
            }
            catch (Exception)
            {
            }
        }

        private static string ConvertSql(string origSql)
        {
            string tmp = origSql.Replace("''", "~~");       // Temporary replacement to simplify matching isolated single quotes
            string baseSql;
            string paramTypes = "";
            string paramData = "";
            int i0 = tmp.IndexOf("'") + 1;
            int i1 = tmp.IndexOf("'", i0);
            if (i1 > 0)
            {
                baseSql = tmp.Substring(i0, i1 - i0);       // Main SQL statement is first parameter in single quotes
                i0 = tmp.IndexOf("'", i1 + 1);
                i1 = tmp.IndexOf("'", i0 + 1);
                if (i0 > 0 && i1 > 0)
                {
                    paramTypes = tmp.Substring(i0 + 1, i1 - i0 - 1);     // This is not required, but retained for reference
                    paramData = tmp.Substring(i1 + 1);
                }
            }
            else
                throw new Exception("Cannot identify SQL statement in first parameter");

            baseSql = baseSql.Replace("~~", "'");           // Undo initial temp replacement, and convert to single instance of single quote
            if (!String.IsNullOrEmpty(paramData))           // Check if there are any parameters to replace
            {
                string[] paramList = paramData.Split(",".ToCharArray());
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("declare " + paramTypes);
                foreach (string paramValue in paramList)
                {
                    if (!paramValue.Contains("@")) continue;
                    sb.AppendLine("set " + paramValue.Replace("~~", "''"));
                }
                sb.AppendLine(baseSql);
                return sb.ToString();
            }
            return baseSql;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.notifyIcon1.Visible = true;
                this.Hide();
            }
            else
            {
                this.notifyIcon1.Visible = false;
            }
        }
    }
}
