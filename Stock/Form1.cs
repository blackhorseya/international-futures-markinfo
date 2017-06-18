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

using HtmlAgilityPack;
using System.Net;
using System.IO;
using System.Globalization;

namespace Stock
{
    public partial class Form1 : Form
    {
        #region Variable
        //item
        int ItemCount = 0;                                  //項目數量
        List<Label> lb_ItemName = new List<Label>();        //label項目名稱
        List<ComboBox> cb_ItemName = new List<ComboBox>();  //combobox項目名稱
        List<DataTable> dt_ItemName = new List<DataTable>();//datatable項目名稱
        List<Button> btn_FilePath = new List<Button>();     //button檔案路徑
        List<TextBox> tb_FilePath = new List<TextBox>();    //textbox檔案路徑
        List<Label> lb_Field = new List<Label>();           //label欄位
        List<TextBox> tb_Field = new List<TextBox>();       //textbox欄位
        List<Button> btn_ItemDel = new List<Button>();      //button項目刪除
        List<Panel> panel_Item = new List<Panel>();         //panel項目

        //Stock Post
        HtmlAgilityPack.HtmlDocument htmldoc = null;
        HtmlNodeCollection nodename, nodeID, _P, _T;
        List<string> FilePath = new List<string>();

        //Text
        public enum MsgType { System, User, Normal, Warning, Error };
        private Color[] MsgTypeColor = { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red };

        //Time
        public int startHour = 8, startMin = 45, endHour = 13, endMin = 50;
        public List<string> timeRules = new List<string>();

        //Class
        StockPost sp = new StockPost();
        ExportExcel excel = new ExportExcel();
        setting setting;
        #endregion

        public Form1()
        {
            InitializeComponent();
            axWindowsMediaPlayer1.settings.volume = 100;
            this.panel1.MouseWheel += new MouseEventHandler(this.panel1_MouseWheel);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            setting = new setting(this);
            setting.button1_Click(null, null);
            while (htmldoc == null)
                htmldoc = sp.Posthtmldoc();
            nodename = htmldoc.DocumentNode.SelectNodes("//a");
            nodeID = htmldoc.DocumentNode.SelectNodes("//li");
        }

        #region panel wheel
        private void panel1_MouseWheel(object sender, EventArgs e)
        {
            if (panel1.AutoScrollPosition.Y > 0)
            {
                panel1.AutoScrollPosition.ToString();
                panel1.VerticalScroll.Value += 10;
                panel1.Invalidate();
            }
        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            this.panel1.Focus();
        }

        private void panel_item_MouseClick(object sender, MouseEventArgs e)
        {
            this.panel1.Focus();
        }
        #endregion 

        #region notifyIcon
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.notifyIcon1.Visible = true;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }
        #endregion

        private void btn_AddItem_Click(object sender, EventArgs e)
        {
            #region new Controls
            dt_ItemName.Add(new DataTable()
            {
                
            });

            dt_ItemName[ItemCount].Columns.Add(new DataColumn("ItemName", typeof(string)));
            dt_ItemName[ItemCount].Columns.Add(new DataColumn("Id", typeof(string)));
            int j = 0;
            foreach (HtmlNode name in nodename)
            {
                dt_ItemName[ItemCount].Rows.Add(dt_ItemName[ItemCount].NewRow());
                dt_ItemName[ItemCount].Rows[j][0] = name.InnerText.Trim();
                j++;
            }
            j = 0;
            foreach (HtmlNode id in nodeID)
            {
                dt_ItemName[ItemCount].Rows[j][1] = id.Id.Trim();
                j++;
            }

            lb_ItemName.Add(new Label()
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(Font.FontFamily, 16.0f),
                Location = new Point(3, 3),
                Text = "項目名稱：",
            });

            cb_ItemName.Add(new ComboBox()
            {
                Location = new Point(lb_ItemName[ItemCount].Width + 25, 0),
                Font = new Font(Font.FontFamily, 16.0f),
                DataSource = dt_ItemName[ItemCount],
                DisplayMember = "ItemName",                                     //combobox.SelectedText
                ValueMember = "Id",                                             //combobox.SelectedValue
            });

            btn_FilePath.Add(new Button()
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(Font.FontFamily, 16.0f),
                Location = new Point(3, 33),
                Text = "選擇路徑",
                Parent = this,
                BackColor = Color.Transparent,
            });

            tb_FilePath.Add(new TextBox()
            {
                Size = new System.Drawing.Size(500, 22),
                Location = new Point(btn_FilePath[ItemCount].Width + 15, 39),
            });

            lb_Field.Add(new Label()
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(Font.FontFamily, 16.0f),
                Location = new Point(lb_ItemName[ItemCount].Width + cb_ItemName[ItemCount].Width + 40, 3),
                Text = "欄位：",
            });

            tb_Field.Add(new TextBox()
            {
                Location = new Point(lb_ItemName[ItemCount].Width + cb_ItemName[ItemCount].Width + lb_Field[ItemCount].Width + 20, 3),
            });

            btn_ItemDel.Add(new Button()
            {
                Size = new System.Drawing.Size(32, 32),
                Image = Image.FromFile(Application.StartupPath + @"\red-cross-icon.png"),
                Location = new Point(lb_ItemName[ItemCount].Width + cb_ItemName[ItemCount].Width + lb_Field[ItemCount].Width + tb_Field[ItemCount].Width + 120, 0),
                FlatStyle = FlatStyle.Flat,
            });
            btn_ItemDel[ItemCount].FlatAppearance.BorderSize = 0;

            panel_Item.Add(new Panel()
            {
                Location = new Point(0, ItemCount * 2 * 40),
                Size = new Size(panel1.Width - 24, 73),
                BackColor = GetRandomColor(),
            });
            #endregion

            for (int i = 0; i < ItemCount + 1; i++)
            {
                btn_ItemDel[i].Tag = i;
                btn_FilePath[i].Tag = i;
                cb_ItemName[i].Tag = i;
            }

            ControlsAdd(ItemCount);

            //自動置底
            Point newPoint = new Point(0, this.panel1.Height - panel1.AutoScrollPosition.Y);
            panel1.AutoScrollPosition = newPoint;

            btn_FilePath[ItemCount].Click += FilePath_Select_Click;
            btn_ItemDel[ItemCount].Click += ItemDel_Click;
            panel_Item[ItemCount].MouseClick += panel_item_MouseClick;

            ItemCount++;
        }

        private void FilePath_Select_Click(object sender, EventArgs e)
        {
            int tag = Convert.ToInt32((sender as Button).Tag.ToString());
            //MessageBox.Show("btn_filepath_tag : " + btn_FilePath[tag].Tag.ToString() + "\n" +
            //                "cb_itemname_text : " + cb_ItemName[tag].Text.ToString() + "\n" +
            //                "cb_itemname_value : " + cb_ItemName[tag].SelectedValue.ToString());

            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == DialogResult.OK)
            {
                tb_FilePath[tag].Text = of.FileName;

                excel.WriteExcel(tb_FilePath[tag].Text, tb_Field[tag].Text);
            }

            //HtmlAgilityPack.HtmlDocument html = null;
            //while (html == null)
            //{ html = sp.PostStockItem(cb_ItemName[tag].SelectedValue.ToString()); }
            //HtmlNodeCollection hn = html.DocumentNode.SelectNodes("//span");
            //foreach (HtmlNode node in hn)
            //{
            //    AddText(MsgType.Warning, node.InnerText.Trim() + "\n");
            //}
        }

        private void ItemDel_Click(object sender, EventArgs e)
        {
            int tag = Convert.ToInt32((sender as Button).Tag.ToString());
            DialogResult result = MessageBox.Show("確定刪除?", "提醒", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                lb_ItemName.RemoveAt(tag);
                dt_ItemName.RemoveAt(tag);
                cb_ItemName.RemoveAt(tag);
                btn_FilePath.RemoveAt(tag);
                tb_FilePath.RemoveAt(tag);
                lb_Field.RemoveAt(tag);
                tb_Field.RemoveAt(tag);
                btn_ItemDel.RemoveAt(tag);
                panel_Item.RemoveAt(tag);
                ItemCount = panel_Item.Count;

                panel1.Controls.Clear();
                for (int i = 0; i < ItemCount; i++)
                {
                    btn_ItemDel[i].Tag = i;
                    btn_FilePath[i].Tag = i;
                    cb_ItemName[i].Tag = i;
                    ControlsAdd(i);
                }
            }
        }

        //Add Text
        public void AddText(MsgType msgtype, string msg)
        {
            richTextBox1.Invoke(new EventHandler(delegate
            {
                richTextBox1.SelectedText = string.Empty;
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, FontStyle.Bold);
                richTextBox1.SelectionColor = MsgTypeColor[(int)msgtype];
                richTextBox1.AppendText(msg);
                richTextBox1.ScrollToCaret();
            }));
        }

        //Add Controls
        private void ControlsAdd(int i)
        {
            panel1.Controls.Add(panel_Item[i]);
            panel_Item[i].Controls.Add(lb_ItemName[i]);
            panel_Item[i].Controls.Add(cb_ItemName[i]);
            panel_Item[i].Controls.Add(btn_FilePath[i]);
            panel_Item[i].Controls.Add(tb_FilePath[i]);
            panel_Item[i].Controls.Add(lb_Field[i]);
            panel_Item[i].Controls.Add(tb_Field[i]);
            panel_Item[i].Controls.Add(btn_ItemDel[i]);
        }

        //Random Color
        public Color GetRandomColor()
        {
            Random rand = new Random((int)DateTime.Now.Ticks);
            Thread.Sleep(rand.Next(50));
            Random rand2 = new Random((int)DateTime.Now.Ticks);

            int red = rand.Next(256);
            int green = rand2.Next(256);
            int blue = (red + green > 400) ? 0 : 400 - red - green;

            if (blue > 255)
                blue = 0;

            return Color.FromArgb(red, green, blue);
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerReportsProgress != true)
            {
                backgroundWorker1.WorkerReportsProgress = true;
                AddText(MsgType.User, "開始下載！\n");
                notifyIcon1.Text = "下載中";
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void btn_Stop_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerReportsProgress == true)
            {
                backgroundWorker1.WorkerReportsProgress = false;
                AddText(MsgType.User, "停止下載！\n");
                notifyIcon1.Text = "即時報價";
                backgroundWorker1.CancelAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (; ; )
            {
                if (backgroundWorker1.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    try
                    {
                        backgroundWorker1.ReportProgress(0);
                        System.Threading.Thread.Sleep(1000);
                    }
                    catch (Exception)
                    {
                        e.Cancel = true;
                        break;
                    }
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            DateTime dt = DateTime.Now;
            DateTime begin = new DateTime(dt.Year, dt.Month, Convert.ToInt32(dt.Date.ToString("dd")), startHour, startMin - 5, 00);
            DateTime end = new DateTime(dt.Year, dt.Month, Convert.ToInt32(dt.Date.ToString("dd")), endHour, endMin + 5, 00);
            bool a = dt.CompareTo(begin) >= 0 ? true : false;
            bool b = dt.CompareTo(end) <= 0 ? true : false;
            //台電金期貨指數即時報價
            if (timeRules.Contains(dt.ToString()) && a && b)
            {
                //計時開始
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Reset(); sw.Start();
                //AddText(MsgType.System, "開始時間：" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "\n");
                AddText(MsgType.System, "開始時間：" + dt.ToString("yyyy/MM/dd HH:mm:ss") + "\n");
                HtmlNodeCollection nodes = null;
                while (nodes == null)
                {
                    nodes = sp.PostTaiwanFutures();
                }

                foreach (HtmlNode node in nodes)
                {
                    HtmlNode tx1 = node.SelectSingleNode("//tr[3]//td[@class='rt r']");
                    HtmlNode txel1 = node.SelectSingleNode("//tr[19]//td[@class='rt r']");
                    HtmlNode txfi1 = node.SelectSingleNode("//tr[22]//td[@class='rt r']");

                    //AddText(MsgType.Error, "台指期：" + tx1.InnerText.Trim(',').Replace(",", "") + "\n"
                    //                      + "電子期：" + txel1.InnerText.Trim(',').Replace(",", "") + "\n"
                    //                      + "金融期：" + txfi1.InnerText.Trim(',').Replace(",", "") + "\n");
                    AddText(MsgType.Error, tx1.InnerText.Trim(',').Replace(",", "") + "\t"
                                          + txel1.InnerText.Trim(',').Replace(",", "") + "\t"
                                          + txfi1.InnerText.Trim(',').Replace(",", "") + "\n");
                    //複製到剪貼簿
                    Clipboard.SetData(DataFormats.Text, tx1.InnerText.Trim(',').Replace(",", "") + "\t"
                                          + txel1.InnerText.Trim(',').Replace(",", "") + "\t"
                                          + txfi1.InnerText.Trim(',').Replace(",", ""));

                    excel.FuturesWriteExcel(tx1.InnerText.Trim().Replace(",", ""),
                                                  txel1.InnerText.Trim().Replace(",", ""),
                                                  txfi1.InnerText.Trim().Replace(",", ""),
                                                  dt.ToString("yyyy/MM/dd HH:mm"));
                }

                //計時結束
                sw.Stop();
                AddText(MsgType.System, "總耗時：" + sw.Elapsed.TotalSeconds.ToString() + " 秒\n");
                axWindowsMediaPlayer1.Ctlcontrols.play();
            }

            //國際期貨指數即時報價
            if (dt.ToString("mm") == "00" && dt.ToString("ss") == "00")
            {
                //計時開始
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Reset(); sw.Start();
                AddText(MsgType.System, "開始時間：" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "\n");

                htmldoc = null;
                List<string> header = new List<string>();
                List<string> point = new List<string>();
                while (htmldoc == null)
                    htmldoc = sp.Posthtmldoc();
                for (int i = 0; i < panel_Item.Count; i++)
                {
                    _P = htmldoc.DocumentNode.SelectNodes("//span[@id='" + cb_ItemName[i].SelectedValue.ToString() + "_P']");
                    _T = htmldoc.DocumentNode.SelectNodes("//span[@id='" + cb_ItemName[i].SelectedValue.ToString() + "_T']");
                    foreach (HtmlNode node in _P)
                    {
                        header.Add(cb_ItemName[i].Text);
                        point.Add(node.InnerText.Trim());
                        AddText(MsgType.Error, cb_ItemName[i].Text + "：" + node.InnerText.Trim() + "    ");
                    }
                    foreach (HtmlNode node in _T)
                    {
                        AddText(MsgType.Error, "時間：" + node.InnerText.Trim() + "\n");
                    }
                }

                excel.FuturesWriteExcel(header, point, dt.ToString("yyyy-MM-dd HH:mm"));

                //計時結束
                sw.Stop();
                AddText(MsgType.System, "總耗時：" + sw.Elapsed.TotalSeconds.ToString() + " 秒\n");
            }
        }

        private void btn_music_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == DialogResult.OK)
            {
                lb_music.Text = of.FileName.ToString();
                axWindowsMediaPlayer1.URL = lb_music.Text;
                axWindowsMediaPlayer1.Ctlcontrols.play();
            }
        }

        private void 設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setting.Show();
        }
    }
}
