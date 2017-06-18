using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Stock
{
    public partial class setting : Form
    {
        Form1 Parent;

        public setting(Form1 _Parent)
        {
            InitializeComponent();
            Parent = _Parent;
        }

        public void button1_Click(object sender, EventArgs e)
        {
            Parent.startHour = (int)nud_sh.Value;
            Parent.startMin = (int)nud_sm.Value;
            Parent.endHour = (int)nud_eh.Value;
            Parent.endMin = (int)nud_em.Value;
            Parent.timeRules.Clear();
            DateTime dt = DateTime.Now;
            DateTime begin = new DateTime(dt.Year, dt.Month, Convert.ToInt32(dt.Date.ToString("dd")), (int)nud_sh.Value, (int)nud_sm.Value - 5, 00);
            DateTime end = new DateTime(dt.Year, dt.Month, Convert.ToInt32(dt.Date.ToString("dd")), (int)nud_eh.Value, (int)nud_em.Value, 00);
            DateTime tick = begin.AddMinutes((double)nud_interval_min.Value).AddSeconds((double)nud_interval_sec.Value);
            bool a = tick.CompareTo(begin) >= 0 ? true : false;
            bool b = tick.CompareTo(end) <= 0 ? true : false;
            while (a && b)
            {
                Parent.timeRules.Add(tick.ToString());
                tick = tick.AddMinutes((double)nud_interval_min.Value).AddSeconds((double)nud_interval_sec.Value);
                a = tick.CompareTo(begin) >= 0 ? true : false;
                b = tick.CompareTo(end) <= 0 ? true : false;
            }
            this.Hide();
        }
    }
}
