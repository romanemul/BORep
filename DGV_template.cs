using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Linq;
using System.Threading;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Back_Order_Report.Properties;
using System.Diagnostics;
using Oracle.ManagedDataAccess.Client;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
using ExcelDataReader;
using System.Reflection;
using EXCEL = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using RRD;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OPENXMLSPREADSHEET = DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Back_Order_Report
{
    public partial class DGV_template : UserControl
    {
        public string filterSTR;
        public string tipSTR;

        private Button c_btn;
        private Point P_C_btn_piont;

        private bool Double_click;
        public bool Load_template_completed;

        private Dictionary<String, ToolTip> tipDCT = new Dictionary<string, ToolTip>();
        private Dictionary<String, Size> Size_txt_filter = new Dictionary<string, Size>();

        private EXCEL.Application xlApp;
        private EXCEL.Workbook xlWb;
        private EXCEL.Worksheet xlWs;
        private EXCEL.Worksheet xlWsBO;
        private EXCEL.Range xlRx;
        private EXCEL.Range xlRy;
        private EXCEL.Range xlRfr;

        private string BAck_Order_File_Name;

        BackgroundWorker b1;

        Setting_DataTable_Variables SDV;
        selected_items si;
        //#####################################################################################################################################################
        public DGV_template(Color color)
        {
            InitializeComponent();

            sum_qty_by_filters_lbl.Visible = false;
            Setting_panel.BackColor = color;

            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.DoubleBuffer, true);

            typeof(DataGridView).InvokeMember("DoubleBuffered",
            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
            null, this.dataGridView_Data, new object[] { true });

            typeof(Panel).InvokeMember("DoubleBuffered",
            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
            null, this.Setting_panel, new object[] { true });

            typeof(DGV_template).InvokeMember("DoubleBuffered",
            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
            null, this, new object[] { true });

            SDV = new Setting_DataTable_Variables();

            tipDCT.Add(Font_btn.Name, SetTooltip(Font_btn, "Nastavit Písmo pro označený sloupec"));
            tipDCT.Add(BackRound_color_btn.Name, SetTooltip(BackRound_color_btn, "Nastavit Barvu pozadí pro označený sloupec"));
            tipDCT.Add(Fore_color_btn.Name, SetTooltip(Fore_color_btn, "Nastavit Barvu písma pro označený sloupec"));
            tipDCT.Add(Show_Hide_Column_btn.Name, SetTooltip(Show_Hide_Column_btn, "Nastavit Zobrazení sloupců"));
            tipDCT.Add(Reset_btn.Name, SetTooltip(Reset_btn, "Smaže hodnoty ve všech Filtrech"));
            tipDCT.Add(conditional_formatting_btn.Name, SetTooltip(conditional_formatting_btn, "Podmíněné Formátování"));

            Load_template_completed = false;
        }
        //#####################################################################################################################################################
        private ToolTip SetTooltip(Control control, string ToolTipTitle, string caption = " ")
        {
            ToolTip tip = new ToolTip();
            tip.ToolTipIcon = ToolTipIcon.Info;
            tip.ToolTipTitle = ToolTipTitle;
            tip.IsBalloon = true;
            tip.ShowAlways = true;
            tip.AutoPopDelay = 30000;
            tip.SetToolTip(control, caption);

            return tip;
        }
        private void DGV_template_Load(object sender, EventArgs e)
        {
            dataGridView_Data.ReadOnly = false;

            dataGridView_Data.TopLeftHeaderCell.Value = dataGridView_Data.RowCount.ToString();
            dataGridView_Data.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView_Data.DefaultCellStyle.Font = new Font("Arial Narrow", 6.75F);
            dataGridView_Data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView_Data.ColumnHeadersDefaultCellStyle.Font = new Font("Arial Narrow", 7.75F);
            dataGridView_Data.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void dataGridView_Data_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (e.ColumnIndex < 0) return;
                if (e.RowIndex < 0) return;

                if (dataGridView_Data.SelectionMode == DataGridViewSelectionMode.ColumnHeaderSelect) dataGridView_Data.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;

                for (int i = 0; i < dataGridView_Data.Columns.Count; i++)
                    dataGridView_Data.Columns[i].SortMode = DataGridViewColumnSortMode.Automatic;

                dataGridView_Data.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;

                dataGridView_Data.InvalidateCell(e.ColumnIndex, e.RowIndex);

                //if (dataGridView_Data.Columns[e.ColumnIndex].Name == "Comments")
                //{
                //    Set_Final_Back_Order_Comment();
                //}

                //set_conditional_formatting(dataGridView_Data);
            }
        }
        private void dataGridView_Data_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Double_click = true;

                dataGridView_Data.ClearSelection();

                for (int i = 0; i < dataGridView_Data.Columns.Count; i++)
                    dataGridView_Data.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                dataGridView_Data.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect;
                dataGridView_Data.Columns[e.ColumnIndex].Selected = true;
            }
        }
        private void dataGridView_Data_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView_Data.SelectionMode == DataGridViewSelectionMode.FullRowSelect)
            {
                dataGridView_Data.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
                if (e.ColumnIndex >= 0) if (e.RowIndex >= 0) dataGridView_Data.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
            }
            else
            {
                dataGridView_Data.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                if (e.RowIndex >= 0) dataGridView_Data.Rows[e.RowIndex].Selected = true;
            }
        }
        private void dataGridView_Data_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) return;

            TextBox[] textboxAll = Setting_panel.Controls.OfType<TextBox>().ToArray();

            foreach (TextBox t in textboxAll)
            {
                t.Width = dataGridView_Data.Columns[t.Name].Width;
                t.Location = new Point(t.Location.X - e.NewValue + e.OldValue, t.Location.Y);

                typeof(TextBox).InvokeMember("DoubleBuffered",
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                null, t, new object[] { true });
            }
        }
        private void dataGridView_Data_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            set_conditional_formatting(dataGridView_Data);
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Font_btn_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table = XRW.Read_xml("FI_setting_column_" + Name);

            foreach (DataGridViewColumn dgvCol in dataGridView_Data.SelectedColumns.Cast<DataGridViewColumn>().ToArray())
            {
                if (table != null)
                {
                    fd.Font = dataGridView_Data.Columns[dgvCol.Index].DefaultCellStyle.Font;

                    if (fd.ShowDialog() == DialogResult.OK)
                    {
                        dataGridView_Data.Columns[dgvCol.Index].DefaultCellStyle.Font = fd.Font;

                        DataRow[] RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dgvCol.Name).ToArray();

                        RowTable[0]["Fnt_name"] = fd.Font.Name;
                        RowTable[0]["Fnt_style"] = fd.Font.Style;
                        RowTable[0]["Fnt_underline"] = fd.Font.Underline;
                        RowTable[0]["Fnt_strikeout"] = fd.Font.Strikeout;
                        RowTable[0]["Fnt_size"] = fd.Font.Size;

                        table.AcceptChanges();
                        XRW.Write_xml(table, "FI_setting_column_" + Name);
                    }
                }
            }
        }
        private void BackRound_color_btn_Click(object sender, EventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            cd.AllowFullOpen = true;
            cd.ShowHelp = true;
            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table = XRW.Read_xml("FI_setting_column_" + Name);

            foreach (DataGridViewColumn dgvCol in dataGridView_Data.SelectedColumns.Cast<DataGridViewColumn>().ToArray())
            {
                cd.Color = dgvCol.DefaultCellStyle.BackColor;
                if (cd.ShowDialog() == DialogResult.OK)
                {
                    if (table != null)
                    {
                        DataRow[] RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dgvCol.Name).ToArray();

                        RowTable[0]["BC"] = ColorTranslator.ToHtml(cd.Color);

                        table.AcceptChanges();
                        XRW.Write_xml(table, "FI_setting_column_" + Name);
                    }
                    dataGridView_Data.Columns[dgvCol.Index].DefaultCellStyle.BackColor = cd.Color;
                }
            }
        }
        private void Fore_color_btn_Click(object sender, EventArgs e)
        {
            ColorDialog cd = new ColorDialog();
            cd.AllowFullOpen = true;
            cd.ShowHelp = true;
            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table = XRW.Read_xml("FI_setting_column_" + Name);

            foreach (DataGridViewColumn dgvCol in dataGridView_Data.SelectedColumns.Cast<DataGridViewColumn>().ToArray())
            {
                cd.Color = dgvCol.DefaultCellStyle.ForeColor;
                if (cd.ShowDialog() == DialogResult.OK)
                {
                    if (table != null)
                    {
                        DataRow[] RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dgvCol.Name).ToArray();

                        RowTable[0]["FC"] = ColorTranslator.ToHtml(cd.Color);

                        table.AcceptChanges();
                        XRW.Write_xml(table, "FI_setting_column_" + Name);
                    }
                    dataGridView_Data.Columns[dgvCol.Index].DefaultCellStyle.ForeColor = cd.Color;
                }
            }
        }
        private void Show_Hide_Column_btn_Click(object sender, EventArgs e)
        {
            if (GridView_panel.Visible)
            {
                GridView_panel.Visible = false;
                GridView_panel.Controls.Clear();
                return;
            }

            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table;

            table = XRW.Read_xml("FI_setting_column_" + Name);

            CheckBox ChB;
            int cyklus = 0;
            int cCyklus = 0;
            int Number_of_lines_in_the_column = 14;

            foreach (DataGridViewColumn dgvCol in dataGridView_Data.Columns.Cast<DataGridViewColumn>().OrderBy(ob => ob.Index))
            {
                bool ColVisible = false;

                if (table != null)
                {
                    DataRow[] dr = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dgvCol.Name).ToArray();

                    ColVisible = (dr.Length == 0 ? false : table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dgvCol.Name).Select(s => Convert.ToBoolean(s["Vissible"])).First());
                }

                dgvCol.Visible = ColVisible;
                cyklus++;
                ChB = new CheckBox();
                ChB.BackColor = Color.Transparent;
                ChB.Width = 300;

                if (cyklus > Number_of_lines_in_the_column)
                {
                    cCyklus++;
                    ChB.Location = new Point(1 + (ChB.Width * cCyklus), 40 + (dgvCol.Index - (Number_of_lines_in_the_column * cCyklus)) * ChB.Height);
                    cyklus = 1;
                }
                else
                {
                    ChB.Location = new Point(
                        1 + (ChB.Width * cCyklus),
                        ((dgvCol.Index - (Number_of_lines_in_the_column * cCyklus) * 1)) + 40 + (dgvCol.Index - (Number_of_lines_in_the_column * cCyklus)) * ChB.Height);
                }

                ChB.CheckedChanged += new EventHandler(ChB_CheckedChanged);
                ChB.Checked = ColVisible;
                ChB.Text = dgvCol.Name;
                GridView_panel.Controls.Add(ChB);
            }


            Button Save_and_Exit_btn = new Button();
            Save_and_Exit_btn.Text = "Zavřít a uložit...";
            Save_and_Exit_btn.BackColor = Color.Transparent;
            Save_and_Exit_btn.BackgroundImage = Properties.Resources.silber_blue;
            Save_and_Exit_btn.BackgroundImageLayout = ImageLayout.Stretch;
            Save_and_Exit_btn.Font = new Font("Microsoft Sans Serif", 12.75F, FontStyle.Bold | FontStyle.Italic);
            Save_and_Exit_btn.Size = new Size(300, 30);
            GridView_panel.Visible = true;

            Save_and_Exit_btn.Location = new Point(GridView_panel.Width - (Save_and_Exit_btn.Width + 5), 5);
            Save_and_Exit_btn.Click += new EventHandler(Save_and_Exit_btn_Click);

            GridView_panel.Controls.Add(Save_and_Exit_btn);
            GridView_panel.Update();
        }
        public void Save_and_Exit_btn_Click(object sender, EventArgs e)
        {
            CheckBox[] chALL = GridView_panel.Controls.OfType<CheckBox>().ToArray();

            DataTable table = new DataTable(Name);
            table.Columns.Add("Column_Name", typeof(String));
            table.Columns.Add("Column_Width", typeof(int));
            table.Columns.Add("Header_Name", typeof(String));
            table.Columns.Add("BC", typeof(String));
            table.Columns.Add("FC", typeof(String));
            table.Columns.Add("Vissible", typeof(Boolean));
            table.Columns.Add("DI", typeof(Byte));
            //-----------------------------------------------
            table.Columns.Add("Fnt_name", typeof(string));
            table.Columns.Add("Fnt_style", typeof(string));
            table.Columns.Add("Fnt_underline", typeof(bool));
            table.Columns.Add("Fnt_strikeout", typeof(bool));
            table.Columns.Add("Fnt_size", typeof(float));
            //-----------------------------------------------
            DataRow row;
            foreach (CheckBox ch in chALL)
            {
                Color BC = dataGridView_Data.Columns[ch.Text].DefaultCellStyle.BackColor;
                Color FC = dataGridView_Data.Columns[ch.Text].DefaultCellStyle.ForeColor;

                dataGridView_Data.Columns[ch.Text].Visible = ch.Checked;

                row = table.NewRow();
                row["Column_Name"] = ch.Text;
                row["Column_Width"] = dataGridView_Data.Columns[ch.Text].Width;
                row["Header_Name"] = (dataGridView_Data.Columns[ch.Text].HeaderText == string.Empty ? ch.Text : dataGridView_Data.Columns[ch.Text].HeaderText);
                row["BC"] = ColorTranslator.ToHtml(BC);
                row["FC"] = ColorTranslator.ToHtml(FC);
                row["Vissible"] = ch.Checked;
                row["DI"] = dataGridView_Data.Columns[ch.Text].DisplayIndex;
                //-------------------------------------------------------------------------------------------------
                bool fnt_exist = (dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font == null ? false : true);

                row["Fnt_name"] = (fnt_exist ? dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font.Name : "Microsoft Sans Serif");
                row["Fnt_style"] = (fnt_exist ? dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font.Style : FontStyle.Regular);
                row["Fnt_underline"] = (fnt_exist ? dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font.Underline : false);
                row["Fnt_strikeout"] = (fnt_exist ? dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font.Strikeout : false);
                row["Fnt_size"] = (fnt_exist ? dataGridView_Data.Columns[ch.Text].DefaultCellStyle.Font.Size : 7.25F);
                //-------------------------------------------------------------------------------------------------
                table.Rows.Add(row);
            }

            table.AcceptChanges();

            XML_Read_Write XRW = new XML_Read_Write();
            XRW.Write_xml(table, "FI_setting_column_" + Name);

            GridView_panel.Visible = false;
            GridView_panel.Controls.Clear();

            set_conditional_formatting(dataGridView_Data);
            Resizing_Filters_for_DGV_template(dataGridView_Data, Setting_panel);
        }
        private void ChB_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
                ((CheckBox)sender).BackColor = Color.GreenYellow;
            else
                ((CheckBox)sender).BackColor = Color.Transparent;
        }
        private void conditional_formatting_btn_Click(object sender, EventArgs e)
        {
            conditional_formatting_menu cfm = new conditional_formatting_menu(this);
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void dataGridView_Data_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (dataGridView_Data.SelectionMode == DataGridViewSelectionMode.ColumnHeaderSelect)
                {
                    if (c_btn != null)
                    {
                        c_btn.Dispose();
                        c_btn = null;
                    }
                }

                if (e.Location.X != P_C_btn_piont.X)
                {
                    if (c_btn != null)
                    {
                        c_btn.Visible = true;
                        c_btn.Location = new Point(e.X - (c_btn.Width / 2), e.Y + c_btn.Height);
                    }
                }
            }
        }
        private void dataGridView_Data_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                var myHitTest = dataGridView_Data.HitTest(e.X, e.Y);

                if (myHitTest.Type != DataGridViewHitTestType.ColumnHeader)
                {
                    if (c_btn != null)
                    {
                        c_btn.Dispose();
                        c_btn = null;
                    }

                    return;
                }

                if (c_btn != null)
                {
                    DataGridViewColumn c = dataGridView_Data.Columns[myHitTest.ColumnIndex];
                    int from_col = dataGridView_Data.Columns[c_btn.Text].DisplayIndex;
                    int to_col = dataGridView_Data.Columns[c.Name].DisplayIndex;

                    if (from_col == to_col || Double_click)
                    {
                        c_btn.Dispose();
                        c_btn = null;
                        return;
                    }

                    XML_Read_Write XRW = new XML_Read_Write();
                    DataTable table;
                    table = XRW.Read_xml("FI_setting_column_" + Name);

                    if (table != null)
                    {
                        DataRow[] RowTable;
                        RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c_btn.Text).ToArray();
                        RowTable[0]["DI"] = to_col;
                        RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).ToArray();
                        RowTable[0]["DI"] = from_col;

                        table.AcceptChanges();
                        XRW.Write_xml(table, "FI_setting_column_" + Name);
                    }

                    c_btn.Text += " jde na místo " + c.Name;
                    c_btn.Location = new Point((e.X - (c_btn.Width / 2) < 0 ? 0 : (e.X - (c_btn.Width / 2))), e.Y + c_btn.Height);

                    Update();

                    Thread.Sleep(500);

                    Update();

                    Thread.Sleep(500);

                    set_conditional_formatting(dataGridView_Data);
                    Resizing_Filters_for_DGV_template(dataGridView_Data, Setting_panel);

                    c_btn.Dispose();
                    c_btn = null;
                }
            }
        }
        private void dataGridView_Data_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Double_click = false;

                P_C_btn_piont = new Point(e.Location.X, e.Location.Y);
                var myHitTest = dataGridView_Data.HitTest(e.X, e.Y);

                if (myHitTest.Type == DataGridViewHitTestType.ColumnHeader)
                {
                    DataGridViewColumn c = dataGridView_Data.Columns[myHitTest.ColumnIndex];

                    c_btn = new Button();
                    c_btn.Name = "Column_button";
                    c_btn.Parent = dataGridView_Data;
                    c_btn.Text = c.Name;
                    c_btn.Font = new Font("Microsoft Sans Serif", 8.75F, FontStyle.Bold);
                    c_btn.AutoSize = true;
                    c_btn.BackColor = Color.AliceBlue;
                    c_btn.Location = new Point(e.X + 5, e.Y);
                    c_btn.Visible = false;
                }
            }

            if (e.Button == MouseButtons.Right)
            {
                P_C_btn_piont = new Point(e.X, e.Y);

                var myHitTest = dataGridView_Data.HitTest(e.X, e.Y);

                if (myHitTest.ColumnIndex < 0)
                {
                    dataGridView_Data.ContextMenuStrip = null;
                }
                else
                {
                    contextMenuStrip1.Items[0].Text = "Přejmenovat Sloupec :" + Environment.NewLine +
                        dataGridView_Data.Columns[myHitTest.ColumnIndex].HeaderText;

                    dataGridView_Data.ContextMenuStrip = contextMenuStrip1;
                }
            }
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Reset_btn_Click(object sender, EventArgs e)
        {
            foreach (TextBox t in Setting_panel.Controls.OfType<TextBox>())
            {
                t.Text = string.Empty;
            }

            Load_template_completed = false;
            txt_filter_KeyDown(this, new KeyEventArgs(Keys.Enter));
            Load_template_completed = true;
        }
        private void přejmenovatSloupecToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var myHitTest = dataGridView_Data.HitTest(P_C_btn_piont.X, P_C_btn_piont.Y);

            if (myHitTest.ColumnIndex < 0) return;

            dataGridView_Data.Columns[myHitTest.ColumnIndex].HeaderText = new Input_Box(dataGridView_Data.Columns[myHitTest.ColumnIndex].HeaderText).GS_New_Text;

            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table;
            table = XRW.Read_xml("FI_setting_column_" + Name);

            if (table != null)
            {
                DataRow[] RowTable;
                RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == dataGridView_Data.Columns[myHitTest.ColumnIndex].Name).ToArray();
                RowTable[0]["Header_Name"] = dataGridView_Data.Columns[myHitTest.ColumnIndex].HeaderText;

                table.AcceptChanges();
                XRW.Write_xml(table, "FI_setting_column_" + Name);
            }
        }
        private void hromadnéHledáníToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var myHitTest = dataGridView_Data.HitTest(P_C_btn_piont.X, P_C_btn_piont.Y);

            if (myHitTest.ColumnIndex < 0) return;

            string slave = Clipboard.GetText();

            DataGridView unitDGV = dataGridView_Data;
            unitDGV.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            DataObject d = unitDGV.GetClipboardContent();

            string master = (d == null ? string.Empty : d.GetText());

            si = new selected_items(slave, master);

            Update();

            TextBox t = Setting_panel.Controls.OfType<TextBox>().Where(w => w.Name == dataGridView_Data.Columns[myHitTest.ColumnIndex].Name).First();

            t.Text = new find_group().GS_New_Filter.Replace(Environment.NewLine, ";");

            txt_filter_KeyDown(t, new KeyEventArgs(Keys.Enter));
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        public void Resizing_Filters_for_DGV_template(DataGridView dataGridView_Data, Panel Setting_panel)
        {
            int count_textbox_filter = Setting_panel.Controls.OfType<TextBox>().Count();

            dataGridView_Data.Invalidate();
            dataGridView_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            GetSet_setting_DatagridView(Setting_panel.Parent.Name, dataGridView_Data.Columns);

            dataGridView_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;

            if (count_textbox_filter <= 0)
            {
                int SumWidth = 0;
                foreach (DataGridViewColumn c in dataGridView_Data.Columns.Cast<DataGridViewColumn>().OrderBy(ob => ob.DisplayIndex))
                {
                    TextBox txt = new TextBox();

                    txt.Name = c.Name;
                    txt.TextAlign = HorizontalAlignment.Center;
                    txt.Font = new Font("Arial Narrow", 9.75F, FontStyle.Bold);
                    txt.KeyDown += new KeyEventHandler(txt_filter_KeyDown);
                    txt.MouseDown += new MouseEventHandler(Txt_MouseDown);
                    txt.Width = c.Width;
                    txt.Location = new Point(dataGridView_Data.RowHeadersWidth + SumWidth, Setting_panel.Height - txt.Height);
                    txt.Visible = c.Visible;
                    txt.CharacterCasing = CharacterCasing.Upper;
                    SumWidth += (c.Visible ? txt.Width : 0);

                    Setting_panel.Controls.Add(txt);
                }
            }
            else
            {
                int SumWidth = 0;
                foreach (DataGridViewColumn c in dataGridView_Data.Columns.Cast<DataGridViewColumn>().OrderBy(ob => ob.DisplayIndex))
                {
                    TextBox[] txt_arr = Setting_panel.Controls.OfType<TextBox>().Where(w => w.Name == c.Name).ToArray();

                    if (txt_arr.Length == 0) return;

                    TextBox txt = txt_arr[0];

                    if (c.Visible)
                    {
                        txt.Visible = true;
                        txt.Width = dataGridView_Data.Columns[txt.Name].Width;
                        txt.Location = new Point(dataGridView_Data.RowHeadersWidth + SumWidth - dataGridView_Data.HorizontalScrollingOffset, Setting_panel.Height - txt.Height);
                        SumWidth += (c.Visible ? txt.Width : 0);

                        txt.Update();
                        Setting_panel.Update();
                    }
                    else
                    {
                        txt.Visible = false;
                    }
                }
            }

        }
        public void set_conditional_formatting(DataGridView DGV)
        {
            DataTable setting_conditional_formatting_DT = new XML_Read_Write().Read_xml("conditional_formatting_menu_" + this.Name);

            if (setting_conditional_formatting_DT == null) return;

            foreach (DataRow dr in setting_conditional_formatting_DT.AsEnumerable().
                Where(w => w["table_name"].ToString() == Name).
                OrderBy(ob => ob["column_name"]).
                OrderBy(ob => ob["condition_value"]).
                ToArray())
            {
                foreach (DataGridViewRow r in DGV.Rows)
                {
                    r.Cells[dr["column_name"].ToString()].Style.BackColor = Color.Empty;
                    r.Cells[dr["column_name"].ToString()].Style.ForeColor = Color.Empty;
                }
            }

            if (setting_conditional_formatting_DT != null)
            {
                foreach (DataRow dr in setting_conditional_formatting_DT.AsEnumerable().
                    Where(w => w["table_name"].ToString() == Name).
                    OrderBy(ob => ob["column_name"]).
                    OrderBy(ob => ob["condition_value"]).
                    ToArray())
                {
                    switch (dr["condition"].ToString())
                    {
                        case "Hodnota v buňce Rovná":
                            #region
                            {
                                DataGridViewRow[] dgvR = null;

                                if (dr["condition_value"].ToString().ToUpper().Contains("NULL"))
                                {
                                    dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                        Where(w => w.Cells[dr["column_name"].ToString()].Value == DBNull.Value).ToArray();

                                    if (dgvR.Length == 0) dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                         Where(w => w.Cells[dr["column_name"].ToString()].Value.ToString() == string.Empty).ToArray();
                                }
                                else
                                {
                                    dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                        Where(w => w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().Equals(dr["condition_value"].ToString().ToUpper())).ToArray();
                                }

                                if (dgvR.Length == 0) dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                     Where(w => w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().
                                     Equals((dr["new_cell_value"].ToString().ToUpper() == string.Empty ? "null" : dr["new_cell_value"].ToString().ToUpper()))).ToArray();

                                foreach (DataGridViewRow r in dgvR)
                                {
                                    r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                    r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                    if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                }

                                break;
                            }
                        #endregion
                        case "Hodnota v buňce Obsahuje":
                            #region
                            {
                                DataGridViewRow[] dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                    Where(w => w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().Contains(dr["condition_value"].ToString().ToUpper())).ToArray();

                                if (dgvR.Length == 0) dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                    Where(w => w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().
                                    Contains((dr["new_cell_value"].ToString().ToUpper() == string.Empty ? "null" : dr["new_cell_value"].ToString().ToUpper()))).ToArray();

                                foreach (DataGridViewRow r in dgvR)
                                {
                                    r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                    r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                    if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                }

                                break;
                            }
                        #endregion
                        case "Hodnota v buňce Ne-Rovná":
                            #region
                            {
                                DataGridViewRow[] dgvR = null;

                                if (dr["condition_value"].ToString().ToUpper().Contains("NULL"))
                                {
                                    dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                        Where(w => w.Cells[dr["column_name"].ToString()].Value != DBNull.Value).ToArray();
                                }
                                else
                                {
                                    dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                        Where(w => !w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().Equals(dr["condition_value"].ToString().ToUpper())).ToArray();
                                }

                                //if (dgvR.Length == 0) dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                //     Where(w => !w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().
                                //     Equals((dr["new_cell_value"].ToString().ToUpper() == string.Empty ? "null" : dr["new_cell_value"].ToString().ToUpper()))).ToArray();

                                foreach (DataGridViewRow r in dgvR)
                                {
                                    r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                    r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                    if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                }

                                break;
                            }
                        #endregion
                        case "Hodnota v buňce Ne-Obsahuje":
                            #region
                            {
                                DataGridViewRow[] dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                    Where(w => !w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().Contains(dr["condition_value"].ToString().ToUpper())).ToArray();

                                if (dgvR.Length == 0) dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                     Where(w => !w.Cells[dr["column_name"].ToString()].Value.ToString().ToUpper().Trim().
                                     Contains((dr["new_cell_value"].ToString().ToUpper() == string.Empty ? "null" : dr["new_cell_value"].ToString().ToUpper()))).ToArray();

                                foreach (DataGridViewRow r in dgvR)
                                {
                                    r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                    r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                    if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                }

                                break;
                            }
                        #endregion
                        case "Hodnota v buňce Je Větší":
                            #region
                            {
                                DataGridViewRow[] dgvR = new DataGridViewRow[DGV.Rows.Count];
                                DataGridViewColumn[] dgvC = dataGridView_Data.Columns.Cast<DataGridViewColumn>().Where(w => w.Name.ToString() == dr["condition_value"].ToString()).ToArray();

                                try
                                {
                                    if (dgvC.Length > 0)
                                    {
                                        dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                            Where(w => w.Cells[dr["column_name"].ToString()].Value != DBNull.Value).ToArray();

                                        Int64 i;
                                        DataGridViewRow[] dgvR_not_string_type = dgvR.AsEnumerable().
                                            Where(w => (Int64.TryParse(w.Cells[dr["column_name"].ToString()].Value.ToString(), out i)) == true).
                                            ToArray();

                                        dgvR = dgvR_not_string_type.AsEnumerable().
                                            Where(w => Convert.ToInt64(w.Cells[dr["column_name"].ToString()].Value) > Convert.ToInt64(w.Cells[dgvC[0].Name].Value)).ToArray();
                                    }
                                    else
                                    {
                                        Int64 i;
                                        DataGridViewRow[] dgvR_not_string_type = DGV.Rows.Cast<DataGridViewRow>().
                                            Where(w => (Int64.TryParse(w.Cells[dr["column_name"].ToString()].Value.ToString(), out i)) == true).
                                            ToArray();

                                        Int64 condition_value = Convert.ToInt64(dr["condition_value"].ToString());
                                        string column_name = dr["column_name"].ToString();

                                        dgvR = dgvR_not_string_type.AsEnumerable().
                                            Where(w => Convert.ToInt64(w.Cells[column_name].Value) > condition_value).ToArray();
                                    }

                                    foreach (DataGridViewRow r in dgvR)
                                    {
                                        r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                        r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                        if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                    }
                                }
                                catch (Exception ex)
                                {

                                }

                                break;
                            }
                        #endregion
                        case "Hodnota v buňce Je Menší":
                            #region
                            {
                                try
                                {
                                    DataGridViewRow[] dgvR = new DataGridViewRow[DGV.Rows.Count];
                                    DataGridViewColumn[] dgvC = dataGridView_Data.Columns.Cast<DataGridViewColumn>().Where(w => w.Name == dr["condition_value"].ToString()).ToArray();

                                    if (dgvC.Length > 0)
                                    {

                                        dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                            Where(w => w.Cells[dr["column_name"].ToString()].Value != DBNull.Value).ToArray();

                                        dgvR = dgvR.Cast<DataGridViewRow>().
                                            Where(w => Convert.ToInt64(w.Cells[dr["column_name"].ToString()].Value) < Convert.ToInt64(w.Cells[dgvC[0].Name].Value)).ToArray();
                                    }
                                    else
                                    {
                                        dgvR = DGV.Rows.Cast<DataGridViewRow>().
                                            Where(w => Convert.ToInt64(w.Cells[dr["column_name"].ToString()].Value) < Convert.ToInt64(dr["condition_value"])).ToArray();
                                    }

                                    foreach (DataGridViewRow r in dgvR)
                                    {
                                        r.Cells[dr["column_name"].ToString()].Style.BackColor = ColorTranslator.FromHtml(dr["BC"].ToString());
                                        r.Cells[dr["column_name"].ToString()].Style.ForeColor = ColorTranslator.FromHtml(dr["FC"].ToString());

                                        if (dr["new_cell_value"].ToString() != string.Empty) r.Cells[dr["column_name"].ToString()].Value = dr["new_cell_value"].ToString();
                                    }
                                }
                                catch { }

                                break;
                            }
                            #endregion
                    }
                }
            }
        }
        private DataGridViewColumnCollection GetSet_setting_DatagridView(string Name, DataGridViewColumnCollection DataGridViewColumnS)
        {
            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table;

            table = XRW.Read_xml("FI_setting_column_" + Name);

            foreach (DataGridViewColumn c in DataGridViewColumnS.Cast<DataGridViewColumn>().AsEnumerable().OrderBy(ob => ob.Index))
            {
                if (table != null)
                {
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                    try
                    {
                        SDV.GS_ColVisible = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => Convert.ToBoolean(s["Vissible"])).First();
                    }
                    catch
                    {
                        continue;
                    }

                    var res_test = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).ToArray();

                    c.Visible = SDV.GS_ColVisible;
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                    SDV.GS_Header_Name = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => s["Header_Name"].ToString()).First();
                    SDV.GS_Column_Width = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => Convert.ToInt32(s["Column_Width"].ToString())).First();
                    SDV.GS_BC = ColorTranslator.FromHtml(table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => s["BC"].ToString()).First());
                    SDV.GS_FC = ColorTranslator.FromHtml(table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => s["FC"].ToString()).First());
                    SDV.GS_DI = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => Convert.ToInt16(s["DI"])).First();
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                    SDV.GS_Fnt_name = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => s["Fnt_name"].ToString()).First();
                    SDV.GS_Fnt_style = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => s["Fnt_style"].ToString()).First();
                    SDV.GS_Fnt_underline = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => (bool)s["Fnt_underline"]).First();
                    SDV.GS_Fnt_strikeout = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => (bool)s["Fnt_strikeout"]).First();
                    SDV.GS_Fnt_size = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == c.Name).Select(s => (float)s["Fnt_size"]).First();
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                    c.HeaderText = SDV.GS_Header_Name;
                    c.Width = SDV.GS_Column_Width;
                    c.DefaultCellStyle.BackColor = SDV.GS_BC;
                    c.DefaultCellStyle.ForeColor = SDV.GS_FC;
                    c.DisplayIndex = SDV.GS_DI;
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                    FontStyle fs = new FontStyle();
                    if (SDV.GS_Fnt_style.IndexOf("Bold") >= 0) fs = fs | FontStyle.Bold;
                    if (SDV.GS_Fnt_style.IndexOf("Italic") >= 0) fs = fs | FontStyle.Italic;
                    if (SDV.GS_Fnt_underline) fs = fs | FontStyle.Underline;
                    if (SDV.GS_Fnt_strikeout) fs = fs | FontStyle.Strikeout;

                    c.DefaultCellStyle.Font = new Font(
                    SDV.GS_Fnt_name,
                    SDV.GS_Fnt_size, fs);
                    //-------------------------------------------------------------------------------------------------------------------------------------------
                }
                else
                {
                    c.Visible = false;
                    c.DisplayIndex = c.Index;
                }
            }

            return DataGridViewColumnS;
        }
        public void txt_filter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                Load_template_completed = false;

                DGV_template DGVT = null;

                if (Variables.DT == null) return;

                if ((((Control)sender).Parent is DGV_template)) Txt_MouseDown(sender, new MouseEventArgs(MouseButtons.Left, 2, 0, 0, 0));

                if (((Control)sender) is DGV_template)
                {
                    DGVT = ((Control)sender) as DGV_template;
                }

                if (((Control)sender).Parent is DGV_template)
                {
                    DGVT = ((Control)sender).Parent as DGV_template;
                }

                if (((Control)sender).Parent.Parent is DGV_template)
                {
                    DGVT = ((Control)sender).Parent.Parent as DGV_template;
                }

                BindingSource bs = new BindingSource();

                switch (DGVT.Name)
                {
                    case "BO_form":
                        {
                            bs.DataSource = Variables.DT.Copy();
                            break;
                        }
                    case "PROD_WORK_ORDER_SEQUENCE_TAB":
                        {
                            bs.DataSource = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.Copy();
                            break;
                        }
                    case "3PAR_Actual_Status":
                        {
                            bs.DataSource = Variables.PAR_Actual_Status.Copy();
                            break;
                        }
                    case "General_Item_Info":
                        {
                            bs.DataSource = Variables.General_Item_Info.Copy();
                            break;
                        }
                    case "SCM_STOCK_REPORT":
                        {
                            bs.DataSource = Variables.SCM_STOCK_REPORT.Copy();
                            break;
                        }
                    case "Rework_tracker_WO_2018_NEW_finance":
                        {
                            bs.DataSource = Variables.Rework_tracker_WO_2018_NEW_finance.Copy();
                            break;
                        }
                }

                StringBuilder filter = new StringBuilder();
                filterSTR = filter.ToString();

                DataTable setting_conditional_formatting_DT = new XML_Read_Write().Read_xml("conditional_formatting_menu_" + DGVT.Name);

                int count_column = DGVT.Setting_panel.Controls.OfType<TextBox>().Where(t => t.Visible).Where(t => t.Text != string.Empty).Count();
                int cyklus = 0;
                int i = 0;

                foreach (TextBox t in DGVT.Setting_panel.Controls.OfType<TextBox>())
                {
                    if (t.Visible == false) continue;
                    if (t.Text == string.Empty) continue;

                    if (setting_conditional_formatting_DT != null)
                    {
                        DataRow[] drArr_exist_column_in_CF = setting_conditional_formatting_DT.AsEnumerable().
                            Where(w => w["table_name"].ToString() == DGVT.Name).
                            Where(w => w["column_name"].ToString() == "`" + t.Name + "`").
                            ToArray();

                        if (drArr_exist_column_in_CF.Length > 0)
                        {
                            bool IsNot = false;
                            if (t.Text.Contains("!")) IsNot = true;
                            if (IsNot) t.Text = t.Text.Replace("!", string.Empty);

                            DataRow[] drArr = setting_conditional_formatting_DT.AsEnumerable().
                                Where(w => w["table_name"].ToString() == DGVT.Name).
                                Where(w => w["column_name"].ToString() == "`" + t.Name + "`").
                                Where(w => w["new_cell_value"].ToString().Contains(t.Text)).
                                ToArray();

                            if (drArr.Length > 0)
                            {
                                t.Text = (IsNot ? "!" : string.Empty) + "*" + t.Text + "*";
                            }
                            else
                            {
                                if (IsNot) t.Text = t.Text = "!" + t.Text;
                            }

                            foreach (DataRow dr in drArr)
                            {
                                if (t.Text == string.Empty)
                                {
                                    t.Text += (IsNot ? "!" : string.Empty) + "*" + dr["condition_value"].ToString() + "*";
                                }
                                else
                                {
                                    if (t.Text.Contains("*" + dr["condition_value"].ToString() + "*")) continue;

                                    t.Text += ";" + (IsNot ? "!" : string.Empty) + "*" + dr["condition_value"].ToString() + "*";
                                }
                            }
                        }
                    }

                    cyklus++;

                    try
                    {
                        #region
                        i = Convert.ToInt32(t.Text.
                            Replace("+", "").
                            Replace("-", "").
                            Replace(">", "").
                            Replace("<", "").
                            Replace("=", ""));

                        bool exist_char = false;

                        if (t.Text.IndexOf("+") >= 0) exist_char = true;
                        if (t.Text.IndexOf("-") >= 0) exist_char = true;
                        if (t.Text.IndexOf(">") >= 0) exist_char = true;
                        if (t.Text.IndexOf("<") >= 0) exist_char = true;
                        if (t.Text.IndexOf("=") >= 0) exist_char = true;
                        #endregion

                        bs.Filter = string.Format("CONVERT({0}, System.Int32) {1}", "`" + t.Name + "`", (exist_char ? t.Text : "=" + t.Text));

                        filter.Append(string.Format("{0} {1}", "`" + t.Name + "`", (exist_char ? t.Text : "=" + t.Text)));
                        if (cyklus != count_column) filter.Append(Environment.NewLine + " AND ");
                        if (cyklus != count_column) filter.Append(Environment.NewLine);

                        filterSTR = filter.ToString();
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            DateTime dt = Convert.ToDateTime(t.Text.
                                Replace(">", "").
                                Replace("<", "").
                                Replace("=", ""));

                            bool exist_char = false;
                            string ch = string.Empty;

                            // IndexOf <,<=,>,>=
                            #region 
                            if (t.Text.IndexOf(">") >= 0)
                            {
                                exist_char = true;
                                ch = ">";
                            }
                            if (t.Text.IndexOf(">=") >= 0)
                            {
                                exist_char = true;
                                ch = ">=";
                            }
                            if (t.Text.IndexOf("<") >= 0)
                            {
                                exist_char = true;
                                ch = "<";
                            }
                            if (t.Text.IndexOf("<=") >= 0)
                            {
                                exist_char = true;
                                ch = "<=";
                            }
                            #endregion

                            filter.Append(string.Format("{0} " + (exist_char ? ch : "=") + " '{1:yyyy-MM-dd}'", "`" + t.Name + "`", dt));
                            if (cyklus != count_column) filter.Append(Environment.NewLine + " AND ");
                            if (cyklus != count_column) filter.Append(Environment.NewLine);
                            filterSTR = filter.ToString();
                        }
                        catch
                        {
                            String[] sArr = t.Text.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

                            int cyklus_arr = 0;

                            t.Text = string.Empty;

                            foreach (string s in sArr)
                            {
                                int cyklus_arr_plus = 0;

                                cyklus_arr++;

                                string ss = null;

                                ss = (s.IndexOf("%") >= 0 ? s.Replace("%", "[%]") : s);
                                ss = ss.Replace("[[", "[").Replace("]]", "]");
                                //ss = (sArr.Length > 1 ? ss.Replace("/", string.Empty) : ss);
                                ss = (sArr.Length > 1 ? ss.Trim() : ss);
                                ss = (sArr.Length > 1 ? (ss.IndexOf("*") >= 0 ? string.Empty : "*") + ss.Trim() + (ss.IndexOf("*") >= 0 ? string.Empty : "*") : ss.Trim());

                                if (ss.IndexOf("+") >= 0)
                                {
                                    String[] sArr_plus = (ss.IndexOf("+") >= 0 ? ss.Split(new string[] { "+" }, StringSplitOptions.RemoveEmptyEntries) : new string[0]);

                                    foreach (string s_plus in sArr_plus)
                                    {
                                        cyklus_arr_plus++;

                                        if ((cyklus_arr == 1) && (cyklus_arr_plus == 1)) filter.Append("(");
                                        if (cyklus_arr_plus == 1) filter.Append("(");

                                        filter.Append(
                                            string.Format("CONVERT({0}, System.String)" + (s_plus.Contains("NULL") ? " IS" : string.Empty) +
                                            (s_plus.Contains("!") ? " NOT" : string.Empty) + (s_plus.Contains("NULL") ? " {1}" : " like '{1}'"),
                                            "`" + t.Name + "`", (s_plus == string.Empty ? "%%" : s_plus.Replace("*", "%").Replace("!", string.Empty))));

                                        filterSTR = filter.ToString();

                                        if (cyklus_arr_plus == sArr_plus.Length) filter.Append(")");
                                        if ((cyklus_arr == sArr.Length) && (cyklus_arr_plus == sArr_plus.Length)) filter.Append(")");

                                        if (cyklus_arr_plus != sArr_plus.Length) filter.Append(" AND ");
                                        if (cyklus_arr_plus != sArr_plus.Length) if ((cyklus_arr_plus % (t.Name.ToUpper().Contains("DESC") ? 5 : 10)) == 0) filter.Append(Environment.NewLine);

                                        if (cyklus_arr_plus != sArr_plus.Length)
                                        {
                                            t.Text += s_plus + "+";
                                        }
                                        else
                                        {
                                            t.Text += s_plus;
                                        }

                                        if ((cyklus_arr_plus == sArr_plus.Length) && (cyklus_arr != sArr.Length)) { filter.Append(" OR "); t.Text += ";"; }
                                        filterSTR = filter.ToString();
                                    }
                                }
                                else
                                {
                                    if (cyklus_arr == 1) filter.Append((ss.Contains("!") ? " NOT " : string.Empty) + "(");

                                    filter.Append(
                                        string.Format("CONVERT({0}, System.String)" + (ss.Contains("NULL") ? " IS" : string.Empty) +
                                        (ss.Contains("NULL") ? " {1}" : " like '{1}'"),
                                        "`" + t.Name + "`", (ss == string.Empty ? "%%" : ss.Replace("*", "%").Replace("!", string.Empty))));

                                    filterSTR = filter.ToString();

                                    if (cyklus_arr == sArr.Length) filter.Append(")");

                                    if (cyklus_arr != sArr.Length)
                                    {
                                        t.Text += ss + ";";
                                    }
                                    else
                                    {
                                        t.Text += ss;
                                    }

                                    if (cyklus_arr != sArr.Length) filter.Append(" OR ");
                                    if (cyklus_arr != sArr.Length) if ((cyklus_arr % (t.Name.ToUpper().Contains("DESC") ? 5 : 10)) == 0) filter.Append(Environment.NewLine);
                                    filterSTR = filter.ToString();
                                }
                            }

                            if (cyklus != count_column) filter.Append(Environment.NewLine + " AND ");
                            if (cyklus != count_column) filter.Append(Environment.NewLine);
                            filterSTR = filter.ToString();
                        }
                    }
                }

                try
                {
                    filterSTR = filter.ToString();

                    bs.Filter = filter.ToString();

                    String[] filrerArray = filter.ToString().
                             Replace("CONVERT(", string.Empty).
                             Replace(", System.String)", string.Empty).
                             Replace("%", "*").
                             Replace("[*]", "%").Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                    TextBox[] tArr = DGVT.Setting_panel.Controls.OfType<TextBox>().Where(w => w.Visible == true).Where(w => w.Text != string.Empty).ToArray();

                    tipSTR = "#######################################################################################################" +
                             "#######################################################################################################" +
                             Environment.NewLine;

                    foreach (String s in filrerArray)
                    {
                        if (s == " AND ")
                        {
                            tipSTR += s.Replace(" ", string.Empty);
                            tipSTR += " ---------------------------------------------------------------------------------------------------------------------------------------------" +
                                      "----------------------------------------------------------------------------------------------------------------------------------------------";
                            tipSTR += Environment.NewLine;
                            continue;
                        }

                        tipSTR += "        " + s;
                        tipSTR += Environment.NewLine;
                    }

                    tipSTR += "#######################################################################################################" +
                              "#######################################################################################################";

                    ToolTip tip = tipDCT[Reset_btn.Name];
                    tip.SetToolTip(Reset_btn, Environment.NewLine + "Detail Filter:" + Environment.NewLine + tipSTR);

                    DataView dv = new DataView();
                    dv = bs.List as DataView;

                    DataTable dt = new DataTable();
                    dt = dv.ToTable().Copy();

                    //DGVT.dataGridView_Data.DataSource = null;

                    DGVT.dataGridView_Data.DataSource = dt;
                    DGVT.dataGridView_Data.TopLeftHeaderCell.Value = DGVT.dataGridView_Data.Rows.Count.ToString();

                    switch (DGVT.Name)
                    {
                        case "BO_form":
                            {

                                break;
                            }
                        case "PROD_WORK_ORDER_SEQUENCE_TAB":
                            {

                                break;
                            }
                        case "3PAR_Actual_Status":
                            {

                                break;
                            }
                        case "General_Item_Info":
                            {

                                break;
                            }
                        case "SCM_STOCK_REPORT":
                            {
                                try
                                {
                                    sum_qty_by_filters_lbl.Text = "Summary QAVAILABLE: " +
                                        dt.AsEnumerable().Sum(sum => Convert.ToInt32(sum["QAVAILABLE"].ToString())).ToString();

                                    sum_qty_by_filters_lbl.Visible = true;
                                }
                                catch { }

                                break;
                            }
                        case "Rework_tracker_WO_2018_NEW_finance":
                            {
                                try
                                {
                                    sum_qty_by_filters_lbl.Text = "Summary Quantity: " +
                                        dt.AsEnumerable().Sum(sum => Convert.ToInt32(sum["Quantity"].ToString())).ToString();

                                    sum_qty_by_filters_lbl.Visible = true;
                                }
                                catch { sum_qty_by_filters_lbl.Visible = false; }

                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Ship-Log 2017", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                set_conditional_formatting(dataGridView_Data);
                Resizing_Filters_for_DGV_template(dataGridView_Data, Setting_panel);

                Load_template_completed = true;
            }
        }
        private void Txt_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (e.Clicks == 2)
                {
                    TextBox t = sender as TextBox;
                    Font f = t.Font;

                    if (!Size_txt_filter.ContainsKey(t.Name)) Size_txt_filter[t.Name] = t.Size;

                    if (t.Size == Size_txt_filter[t.Name])
                    {
                        t.Parent = this;
                        t.Multiline = true;
                        t.Size = new Size(1000, 150);
                        t.Font = new Font(t.Font.Name, 12F, FontStyle.Bold);
                    }
                    else
                    {
                        t.Parent = Setting_panel;
                        t.Multiline = false;
                        t.Size = Size_txt_filter[t.Name];
                        t.Font = new Font(t.Font.Name, 9.75F, FontStyle.Bold);

                        Size_txt_filter.Remove(t.Name);
                        dataGridView_Data.Update();
                    }

                    t.BringToFront();
                    t.Update();
                    t.Focus();
                }
            }
        }
        private void dataGridView_Data_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            XML_Read_Write XRW = new XML_Read_Write();
            DataTable table = XRW.Read_xml("FI_setting_column_" + Name);

            if (table != null)
            {
                DataRow[] RowTable = table.AsEnumerable().Where(w => w["Column_Name"].ToString() == e.Column.Name).ToArray();
                RowTable[0]["Column_Width"] = e.Column.Width;
                table.AcceptChanges();
                XRW.Write_xml(table, "FI_setting_column_" + Name);
            }

            if (Load_template_completed) Resizing_Filters_for_DGV_template(dataGridView_Data, Setting_panel);
        }
        //#####################################################################################################################################################
        private void load_Back_Order_from_excel_btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                load_BO_from_excel_btn.Visible = false;
                progressBar1.Visible = true;

                BAck_Order_File_Name = ofd.FileName;

                Variables.DT = Excel_Reader(ofd.FileName).Tables[1];

                b1 = new BackgroundWorker();
                b1.WorkerReportsProgress = true;
                b1.DoWork += new DoWorkEventHandler(B1_DoWork);
                b1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(B1_RunWorkerCompleted);
                b1.ProgressChanged += new ProgressChangedEventHandler(B1_ProgressChanged);
                b1.RunWorkerAsync();

                //B1_Dowork();
            }
        }

        //------------------------------------------------------------------------------
        private void B1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            progressBar1.Update();
        }
        private void B1_DoWork(object sender, DoWorkEventArgs e)
        {
            b1.ReportProgress(10);
            Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB();

            b1.ReportProgress(20);
            Add_column_from_3PAR_Actual_Status_2018();

            b1.ReportProgress(30);
            Add_column_from_General_Item_Info_AR_all_AB();
            Add_column_from_SCM_STOCK_REPORT_AR_all_AB();
            Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB_AR_all_AB();

            b1.ReportProgress(40);
            Add_column_from_General_Item_Info_AR_current_AB();
            Add_column_from_SCM_STOCK_REPORT_AR_current_AB();
            Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB_AR_current_AB();
            b1.ReportProgress(60);

            Add_column_from_3PAR_Actual_Status_2018_AR_all_AB();
            b1.ReportProgress(70);

            Add_column_from_Rework_tracker();
            b1.ReportProgress(80);

            Print_WO_Alternatives_To_Excel();
            Set_Final_Back_Order_Comment();
            Writing_to_Excel_Comment();
            b1.ReportProgress(100);
        }
        private void B1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dataGridView_Data.DataSource = Variables.DT;
            Load_template_completed = false;
            txt_filter_KeyDown(this, new KeyEventArgs(Keys.Enter));
            Load_template_completed = true;

            load_BO_from_excel_btn.Visible = true;
            progressBar1.Visible = false;

            b1.Dispose();
            b1 = null;

            //Writing_to_Excel_Details();
        }
        //#####################################################################################################################################################
        private void Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[2];
                dc_arr[0] = new DataColumn("sum_quantities_for_open_Work_orders", typeof(int));
                dc_arr[1] = new DataColumn("Data_for_open_Work_orders", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                string cmd =
                    @"SELECT
                    MAX(T1.WO_TYPE), T1.WO_NUMBER, MAX(T1.BRANCH), MAX(T1.ITEM_NUMBER), MAX(T1.DESCRIPTION), MAX(T1.STATUS), 
                    MAX(T1.CUSTOMER), MAX(T1.QTY_TOTAL), MAX(T1.QTY_BUILT), MAX(T1.QTY_OPEN), MAX(T2.TEST_YN), MAX(T2.TIME_END_PLANNED), MAX(T2.LINE_ID)
                    FROM
                    RPTDBLINK.WOPL_ALL_PL_PUBLIC T1, BRNODATA.PROD_WORK_ORDER_SEQUENCE_TAB T2
                    WHERE
                    (T1.LOTLPS = 'LPS' AND
                    T1.QTY_OPEN > 0 AND
                    T1.STATUS != '95' AND T1.STATUS != '96' AND T1.STATUS != '98' AND T1.STATUS != '99' AND
                    T1.BRANCH = '         728' AND
                    T1.WO_NUMBER = T2.WORK_ORDER AND
                    T2.WORK_ORDER > 0)
                    GROUP BY T1.WO_NUMBER";

                Variables.PROD_WORK_ORDER_SEQUENCE_TAB.Clear();

                using (OracleDataAdapter daOra = new OracleDataAdapter(cmd, Variables.OraConn_DBS2))
                {
                    daOra.Fill(Variables.PROD_WORK_ORDER_SEQUENCE_TAB);
                }

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    // secte QTY open
                    // secti vsechna otevrena WO pro PN 
                    //T1.ItemNumber == PN nebo T1.Description == PN a secti QTY OPEN
                    dr["sum_quantities_for_open_Work_orders"] =
                        Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                   w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                        Sum(sum => Convert.ToInt64(sum["MAX(T1.QTY_OPEN)"].ToString()));
                    //vypis vsech WO QTY open pcs + time end a test Y/N
                    string[] res_WO =
                        Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                   w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                        Select(s => s["MAX(T1.QTY_OPEN)"].ToString() + " pcs " + s["MAX(T2.TIME_END_PLANNED)"].ToString() + " Test=" + s["MAX(T2.TEST_YN)"].ToString()).
                        ToArray();

                    string Join_res_WO = string.Join("; ", res_WO);
                    //zapis do sloupce "Data_for_open_Work_orders"
                    dr["Data_for_open_Work_orders"] = (Join_res_WO == string.Empty ? null : Join_res_WO);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        private void bRNODATAPRODWORKORDERSEQUENCETABToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form();
            f.Icon = Resources.left;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Width = 1500;
            f.Height = 800;

            DGV_template dgvt = new DGV_template(Color.FromArgb(255, 224, 192));
            dgvt.load_BO_from_excel_btn.Visible = false;
            dgvt.Name = "PROD_WORK_ORDER_SEQUENCE_TAB";
            f.Text = dgvt.Name;
            f.Controls.Add(dgvt);
            dgvt.Dock = DockStyle.Fill;
            f.Update();

            f.Show();

            dgvt.dataGridView_Data.DataSource = Variables.PROD_WORK_ORDER_SEQUENCE_TAB;

            dgvt.Load_template_completed = false;
            dgvt.txt_filter_KeyDown(dgvt, new KeyEventArgs(Keys.Enter));
            dgvt.Load_template_completed = true;
        }
        //#####################################################################################################################################################
        private void Add_column_from_3PAR_Actual_Status_2018()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("3PAR_Actual_Status", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);
                Variables.PAR_Actual_Status = Excel_Reader(Resources.Patch_3PAR_Actual_Status_2019).Tables[0];

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    //fg pn = excel pn where status open
                    // vyber qty + WO number
                    var res_WO = Variables.PAR_Actual_Status.AsEnumerable().
                        Where(w => w["FG PN"].ToString() == dr["PN"].ToString()).
                        Where(w => !w["Status"].ToString().ToUpper().Contains("CLOSE")).
                        Select(s => s["Tested Qty"].ToString() + " for " + s["WO"].ToString()).
                        ToArray();

                    int res_qty = 0;
                    //pokud vice WO tak sectu pocet ks
                    try
                    {
                        res_qty = Variables.PAR_Actual_Status.AsEnumerable().
                            Where(w => w["FG PN"].ToString() == dr["PN"].ToString()).
                            Where(w => !w["Status"].ToString().ToUpper().Contains("CLOSE")).
                            Sum(s => Convert.ToInt32(s["Tested Qty"].ToString()));
                    }
                    catch { }

                    // a cislo WO
                    if (res_WO.Length > 0)
                    {
                        string Join_res_WO = string.Join("; ", res_WO);
                        dr["3PAR_Actual_Status"] = res_qty + " = " + Join_res_WO;
                    }
                    // pokud nic tak prazdna bunka
                    if (dr["3PAR_Actual_Status"].ToString() == string.Empty) dr["3PAR_Actual_Status"] = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        private void pARActualStatusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form();
            f.Icon = Resources.left;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Width = 1500;
            f.Height = 800;

            DGV_template dgvt = new DGV_template(Color.FromArgb(192, 255, 192));
            dgvt.load_BO_from_excel_btn.Visible = false;
            dgvt.Name = "3PAR_Actual_Status";
            f.Text = dgvt.Name;
            f.Controls.Add(dgvt);
            dgvt.Dock = DockStyle.Fill;
            f.Update();

            f.Show();

            dgvt.dataGridView_Data.DataSource = Variables.PAR_Actual_Status;

            dgvt.Load_template_completed = false;
            dgvt.txt_filter_KeyDown(dgvt, new KeyEventArgs(Keys.Enter));
            dgvt.Load_template_completed = true;
        }
        //#####################################################################################################################################################
        private void Add_column_from_General_Item_Info_AR_all_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("General_Item_Info_AR_all_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);
                //Variables.General_Item_Info = Excel_Reader(Resources.Patch_General_Item_Info).Tables[0];

                string[] HPPN_arr = Variables.DT.AsEnumerable().Select(s => "'" + s["HP PN"].ToString() + "'").ToArray();

                StringBuilder CmdTextGII = new StringBuilder("SELECT " +
                    "CUSTOMER_NUMBER, " +
                    "HP_PART_NO, " +
                    "OEM_PART_NO, " +
                    "TEST_HDD, " +
                    "ITEM_NUMBER, " +
                    "BRANCH_PLANT, " +
                    "DESCRIPTION, " +
                    "DESCRIPTION2, " +
                    "PRODUCT_FAMILY, " +
                    "SUPPLIER_NUMBER, " +
                    "SUPPLIER_NAME " +
                    "FROM RPTDBLINK.GII_ITEM_MERGED_INFO WHERE ");

                int i = 0;
                foreach (string HPPN in HPPN_arr)
                {
                    i++;
                    CmdTextGII.Append("HP_PART_NO = " + HPPN);
                    if (i < HPPN_arr.Length) CmdTextGII.Append(" OR ");
                }

                using (OracleDataAdapter OraSqlDa = new OracleDataAdapter(CmdTextGII.ToString(), Variables.OraConn_DBS2))
                {
                    Variables.General_Item_Info = new DataTable();
                    OraSqlDa.Fill(Variables.General_Item_Info);
                }
                
                
                //vytvari seznam vsech alternativ z GII a zapisuje do General_Item_Info_AR_all_AB krome sama sebe tedy dr.itemArray[0]


               
                foreach (DataRow dr in Variables.DT.Rows)
                {
                    var res_GII = Variables.General_Item_Info.AsEnumerable().
                        Where(w =>
                        w["HP_PART_NO"].ToString() == dr["HP PN"].ToString() && w["ITEM_NUMBER"].ToString() != dr["PN"].ToString()).
                        Select(s => s["ITEM_NUMBER"].ToString()).
                        ToArray();

                    if (res_GII.Length > 0)
                    {
                        string Join_res_GII = string.Join("; ", res_GII);
                        dr["General_Item_Info_AR_all_AB"] = Join_res_GII;
                    }

                    if (dr["General_Item_Info_AR_all_AB"].ToString() == string.Empty) dr["General_Item_Info_AR_all_AB"] = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString(), "Back_Order_Report");
            }
        }
        private void generalItemInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form();
            f.Icon = Resources.left;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Width = 1500;
            f.Height = 800;

            DGV_template dgvt = new DGV_template(Color.FromArgb(255, 255, 192));
            dgvt.load_BO_from_excel_btn.Visible = false;
            dgvt.Name = "General_Item_Info";
            f.Text = dgvt.Name;
            f.Controls.Add(dgvt);
            dgvt.Dock = DockStyle.Fill;
            f.Update();

            f.Show();

            dgvt.dataGridView_Data.DataSource = Variables.General_Item_Info;

            dgvt.Load_template_completed = false;
            dgvt.txt_filter_KeyDown(dgvt, new KeyEventArgs(Keys.Enter));
            dgvt.Load_template_completed = true;
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Add_column_from_SCM_STOCK_REPORT_AR_all_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("SCM_STOCK_REPORT_QAVAILABLE_AR_all_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                string[] res_All_General_Item_Info = Variables.DT.AsEnumerable().Select(s => s["General_Item_Info_AR_all_AB"].ToString()).ToArray();
                string join_GII = string.Join("; ", res_All_General_Item_Info);
                string[] split_GII = join_GII.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);

                split_GII = split_GII.GroupBy(gb => new { GII = gb }).Select(S => S.Key.GII).ToArray();


                string[] res_PN = Variables.DT.AsEnumerable().GroupBy(gb => new { PN = gb["PN"].ToString() }).Select(s => s.Key.PN).ToArray();

                string cmd = "SELECT * FROM RPTDBLINK.SCM_STOCK_REPORT WHERE \"Item Number\" IN ('" + string.Join("', '", split_GII) + "') or " +
                             "\"Item Number\" IN ('" + string.Join("', '", res_PN) + "')";

                Variables.SCM_STOCK_REPORT.Clear();

                using (OracleDataAdapter daOra = new OracleDataAdapter(cmd, Variables.OraConn_DBS2))
                {
                    daOra.Fill(Variables.SCM_STOCK_REPORT);
                }
                //ERROR in counting pravdepodobne 
                //zpocitava pocet ks na alternativy QAVAILIABLE

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    string[] res_item = dr["General_Item_Info_AR_all_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);

                    int QAVAILABLE = 0;
                    foreach (string itm in res_item)
                    {
                        QAVAILABLE += Variables.SCM_STOCK_REPORT.AsEnumerable()
                            .Where(w => w["Item Number"].ToString() == itm)
                            .Sum(s => Convert.ToInt32(s["QAVAILABLE"].ToString()));
                    }

                    dr["SCM_STOCK_REPORT_QAVAILABLE_AR_all_AB"] = QAVAILABLE.ToString();

                    if (dr["SCM_STOCK_REPORT_QAVAILABLE_AR_all_AB"].ToString() == string.Empty) dr["SCM_STOCK_REPORT_QAVAILABLE_AR_all_AB"] = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }

        }
        private void sCMSTOCKREPORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form
            {
                Icon = Resources.left,
                StartPosition = FormStartPosition.CenterScreen,
                Width = 1500,
                Height = 800
            };

            DGV_template dgvt = new DGV_template(Color.FromArgb(192, 255, 255));
            dgvt.load_BO_from_excel_btn.Visible = false;
            dgvt.Name = "SCM_STOCK_REPORT";
            f.Text = dgvt.Name;
            f.Controls.Add(dgvt);
            dgvt.Dock = DockStyle.Fill;
            f.Update();

            f.Show();

            dgvt.dataGridView_Data.DataSource = Variables.SCM_STOCK_REPORT;

            dgvt.Load_template_completed = false;
            dgvt.txt_filter_KeyDown(dgvt, new KeyEventArgs(Keys.Enter));
            dgvt.Load_template_completed = true;
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB_AR_all_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[2];
                dc_arr[0] = new DataColumn("sum_quantities_for_open_Work_orders_AR_all_AB", typeof(int));
                dc_arr[1] = new DataColumn("Data_for_open_Work_orders_AR_all_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    string[] res_item = dr["General_Item_Info_AR_all_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    //vyhledava vse ze sloupecku General_Item_Info_AR_all_AB a spocitava otevrene WO
                    var tmpPN = dr["General_Item_Info_AR_all_AB"].ToString().Split("','".ToCharArray(),StringSplitOptions.RemoveEmptyEntries);
                    int QAVAILABLE = 0;

                    //pro alternativu v prod order sequence tab spocita polozky QTY_OPEN
                    foreach (string itm in res_item)
                    {
                        QAVAILABLE += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == itm).Sum(sum => Convert.ToInt32(sum["MAX(T1.QTY_OPEN)"].ToString()));


                        string[] res_WO =
                            Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == itm).
                            Select(s =>
                            s["MAX(T1.QTY_OPEN)"].ToString() + " pcs " +
                            s["MAX(T1.ITEM_NUMBER)"].ToString() + " " +
                            s["MAX(T2.TIME_END_PLANNED)"].ToString() +
                            " Test=" + s["MAX(T2.TEST_YN)"].ToString()).
                            ToArray();

                        string Join_res_WO = string.Join("; ", res_WO);
                        dr["Data_for_open_Work_orders_AR_all_AB"] += (Join_res_WO == string.Empty ? null : Join_res_WO);

                        if (dr["Data_for_open_Work_orders_AR_all_AB"].ToString() == string.Empty) dr["Data_for_open_Work_orders_AR_all_AB"] = null;
                    }


                    dr["sum_quantities_for_open_Work_orders_AR_all_AB"] = QAVAILABLE.ToString();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Add_column_from_3PAR_Actual_Status_2018_AR_all_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("3PAR_Actual_Status_AR_all_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    // General_Item_Info_AR_all_AB
                    //vyhleda alternativy v testech
                    string[] split_GII = dr["General_Item_Info_AR_all_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);

                    var res_WO = Variables.PAR_Actual_Status.AsEnumerable().
                        Where(w => split_GII.Any(GIIarr => w["FG PN"].ToString() == GIIarr)).
                        Where(w => !w["Status"].ToString().ToUpper().Contains("CLOSE")).
                        Select(s => s["Tested Qty"].ToString() + " for " + s["WO"].ToString()).
                        ToArray();





                    int res_qty = 0;

                    try
                    {
                        res_qty = Variables.PAR_Actual_Status.AsEnumerable().
                            Where(w => split_GII.Any(l => w["FG PN"].ToString() == l)).
                            Where(w => !w["Status"].ToString().ToUpper().Contains("CLOSE")).
                            Sum(s => Convert.ToInt32(s["Tested Qty"].ToString()));
                    }
                    catch { }


                    if (res_WO.Length > 0)
                    {
                        string Join_res_WO = string.Join("; ", res_WO);
                        dr["3PAR_Actual_Status_AR_all_AB"] = res_qty + " = " + Join_res_WO;
                    }

                    if (dr["3PAR_Actual_Status_AR_all_AB"].ToString() == string.Empty) dr["3PAR_Actual_Status_AR_all_AB"] = null;
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        //#####################################################################################################################################################
        private void Add_column_from_General_Item_Info_AR_current_AB()
        {
            DataRow drtmp = null;

            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("General_Item_Info_AR_current_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);
                foreach (DataRow dr in Variables.DT.Rows)
                {
                    drtmp = dr;

                    // vyhleda item number podle PN a vrati customer numbers 
                    var res_GII_Customer = Variables.General_Item_Info.AsEnumerable().
                        Where(w => w["ITEM_NUMBER"].ToString() == dr["PN"].ToString()).
                        Select(s => s["CUSTOMER_NUMBER"].ToString()).
                        ToArray();
                    
                    var res_GII = Variables.General_Item_Info.AsEnumerable().
                        Where(w =>
                        w["HP_PART_NO"].ToString() == dr["HP PN"].ToString() &&
                        w["CUSTOMER_NUMBER"].ToString() == res_GII_Customer[0] &&
                        w["ITEM_NUMBER"].ToString() != dr["PN"].ToString()).
                        Select(s => s["ITEM_NUMBER"].ToString()).
                        ToArray();

                    if (res_GII.Length > 0)
                    {
                        string Join_res_GII = string.Join("; ", res_GII);
                        dr["General_Item_Info_AR_current_AB"] = Join_res_GII;
                    }

                    if (dr["General_Item_Info_AR_current_AB"].ToString() == string.Empty) dr["General_Item_Info_AR_current_AB"] = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString() + Environment.NewLine + Environment.NewLine +
                    "V reportu: General_Item_Info " +
                    "Neexistuje Item: " + drtmp["PN"].ToString() +
                    ", nebo HP PN: " + drtmp["HP PN"].ToString() + Environment.NewLine +
                    "Znamená to, že pro tento Item nebudou zjištěny údaje pro Alternativní Revize", "Back_Order_Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Add_column_from_SCM_STOCK_REPORT_AR_current_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[1];
                dc_arr[0] = new DataColumn("SCM_STOCK_REPORT_QAVAILABLE_AR_current_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    string[] res_item = dr["General_Item_Info_AR_current_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    //hleda res_item "Item number" ve stock reportu
                    //column 12 SCM_STOCK_REPORT_QAVAILABLE_AR_current_AB
                    int QAVAILABLE = 0;
                    foreach (string itm in res_item)
                    {
                        QAVAILABLE += Variables.SCM_STOCK_REPORT.AsEnumerable()
                            .Where(w => w["Item Number"].ToString() == itm)
                            .Sum(s => Convert.ToInt32(s["QAVAILABLE"].ToString()));
                    }

                    dr["SCM_STOCK_REPORT_QAVAILABLE_AR_current_AB"] = QAVAILABLE.ToString();

                    if (dr["SCM_STOCK_REPORT_QAVAILABLE_AR_current_AB"].ToString() == string.Empty) dr["SCM_STOCK_REPORT_QAVAILABLE_AR_current_AB"] = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------
        private void Add_column_from_PROD_WORK_ORDER_SEQUENCE_TAB_AR_current_AB()
        {
            try
            {
                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[2];
                dc_arr[0] = new DataColumn("sum_quantities_for_open_Work_orders_AR_current_AB", typeof(int));
                dc_arr[1] = new DataColumn("Data_for_open_Work_orders_AR_current_AB", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    string[] res_item = dr["General_Item_Info_AR_current_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    //column 13 sum_quantities_for_open_Work_orders_AR_current_AB
                    //column 14 Data_for_open_Work_orders_AR_current_AB
                    int QAVAILABLE = 0;
                    foreach (string itm in res_item)
                    {
                        QAVAILABLE += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable()
                            .Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == itm)
                            .Sum(sum => Convert.ToInt32(sum["MAX(T1.QTY_OPEN)"].ToString()));


                        string[] res_WO =
                            Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == itm).
                            Select(s => s["MAX(T1.QTY_OPEN)"].ToString() + " pcs " + s["MAX(T2.TIME_END_PLANNED)"].ToString() + " Test=" + s["MAX(T2.TEST_YN)"].ToString()).
                            ToArray();

                        string Join_res_WO = string.Join("; ", res_WO);
                        dr["Data_for_open_Work_orders_AR_current_AB"] += (Join_res_WO == string.Empty ? null : Join_res_WO);

                        if (dr["Data_for_open_Work_orders_AR_current_AB"].ToString() == string.Empty) dr["Data_for_open_Work_orders_AR_current_AB"] = null;
                    }


                    dr["sum_quantities_for_open_Work_orders_AR_current_AB"] = QAVAILABLE.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        //#####################################################################################################################################################
        private void Add_column_from_Rework_tracker()
        {
            try
            {
                using (MySqlDataAdapter DA = new MySqlDataAdapter("SELECT * FROM scm.rework_tracker;", Variables.MySqlConn_DBS4))
                {
                    DA.Fill(Variables.Rework_tracker_WO_2018_NEW_finance);
                }

                DataColumn[] dc_arr = null;
                //----------------------------------------------------------------
                dc_arr = new DataColumn[2];
                dc_arr[0] = new DataColumn("sum_quantities_Rework_tracker", typeof(int));
                dc_arr[1] = new DataColumn("Data_from_Rework_tracker", typeof(string));
                Variables.DT.Columns.AddRange(dc_arr);

                foreach (DataRow dr in Variables.DT.Rows)
                {
                    DataRow[] res_RT =
                        Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                        Where(w => w["TO HP"].ToString() == dr["HP PN"].ToString()).
                        Where(w => w["Status"].ToString().ToUpper().Trim() == "OPEN").
                        ToArray();

                    StringBuilder sb = new StringBuilder();
                    int i = 0;
                    foreach (DataRow res_dr in res_RT)
                    {
                        i++;
                        sb.Append(res_dr["Quantity"].ToString() + "pcs - ");
                        sb.Append(res_dr["BO comment"].ToString());
                        if (i < res_RT.Length) sb.Append("; ");
                    }

                    try
                    {
                        dr["sum_quantities_Rework_tracker"] = res_RT.AsEnumerable().Sum(sum => Convert.ToInt32(sum["Quantity"].ToString()));
                    }
                    catch
                    {
                        dr["sum_quantities_Rework_tracker"] = "0";
                    }

                    dr["Data_from_Rework_tracker"] = (sb.ToString() == string.Empty ? null : sb.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            }
        }
        private void ReworktrackerWO2018NEWfinanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form
            {
                Icon = Resources.left,
                StartPosition = FormStartPosition.CenterScreen,
                Width = 1500,
                Height = 800
            };

            DGV_template dgvt = new DGV_template(Color.FromArgb(255, 192, 192));
            dgvt.load_BO_from_excel_btn.Visible = false;
            dgvt.Name = "Rework_tracker_WO_2018_NEW_finance";
            f.Text = dgvt.Name;
            f.Controls.Add(dgvt);
            dgvt.Dock = DockStyle.Fill;
            f.Update();

            f.Show();

            dgvt.dataGridView_Data.DataSource = Variables.Rework_tracker_WO_2018_NEW_finance;

            dgvt.Load_template_completed = false;
            dgvt.txt_filter_KeyDown(dgvt, new KeyEventArgs(Keys.Enter));
            dgvt.Load_template_completed = true;
        }
        //#####################################################################################################################################################
        public DataSet Excel_Reader(string file_name)
        {
            var file = new FileInfo(file_name);
            using (FileStream fileStream = new FileStream(file_name, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader;

                if (file.Extension.Equals(".xls"))
                    reader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                else
                    reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

                //// reader.IsFirstRowAsColumnNames
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                return reader.AsDataSet(conf);
            }
        }
        //#####################################################################################################################################################
        private void Set_Final_Back_Order_Comment()
        {
            DataRow drtmp = null;

            try
            {
                DataRow[] res_dr = Variables.DT.AsEnumerable().ToArray();

                foreach (DataRow dr in res_dr)

                {
                    drtmp = dr;

                    StringBuilder Comments_sb = new StringBuilder();

                    // PROD_WORK_ORDER_SEQUENCE_TAB does not exist in the table
                    //PN bud v description nebo Item number v PROD_WORK_ORDER_SEQUENCE_TAB
                    #region
                    DataRow[] res_WO_does_not_exist_in_the_table =
                        Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                   w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                        ToArray();


                    if (res_WO_does_not_exist_in_the_table.Length == 0) Comments_sb.Append("no build scheduled");
                    #endregion

                    // PROD_WORK_ORDER_SEQUENCE_TAB Test=N
                    #region
                    DataRow[] res_WO =
                        Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                   w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                        Where(w => w["MAX(T2.TEST_YN)"].ToString().ToUpper().Equals("N")).
                        ToArray();

                    int cyklus = 0;
                    foreach (DataRow WO in res_WO)
                    { 
                        cyklus++;

                        DataRow[] res_RT_exist_WO =
                            Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                            Where(w => w["WO number"].ToString().Trim() == WO["WO_NUMBER"].ToString()).
                            Where(w => w["BO comment"].ToString().Trim() != string.Empty).
                            ToArray();

                        Comments_sb.Append(WO["MAX(T1.QTY_OPEN)"].ToString());

                        if (res_RT_exist_WO.Length > 0)
                        {
                            Comments_sb.Append(" rework " + res_RT_exist_WO[0]["BO comment"].ToString());
                        }
                        else
                        {
                            Comments_sb.Append(" build ");

                            DateTime TIME_END_PLANNED = DateTime.MinValue;
                            try { TIME_END_PLANNED = Convert.ToDateTime(WO["MAX(T2.TIME_END_PLANNED)"].ToString()); } catch { }

                            if (TIME_END_PLANNED == DateTime.MinValue)
                            {
                                #region
                                TIME_END_PLANNED = DateTime.Now;

                                if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                }

                                Comments_sb.Append("/");

                                if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(3).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                }
                                #endregion

                            }
                            else
                            {
                                #region
                                DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);

                                if (TIME_END_PLANNED < today)
                                {
                                    TIME_END_PLANNED = today;
                                }

                                if (TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString().Substring(0, 3));
                                }

                                Comments_sb.Append("/");

                                if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                }
                                #endregion

                            }
                        }


                        if (cyklus < res_WO.Length) Comments_sb.Append("; ");
                    }
                    #endregion

                    // PROD_WORK_ORDER_SEQUENCE_TAB Test=Y
                    #region
                    res_WO =
                        Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                   w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                        Where(w => w["MAX(T2.TEST_YN)"].ToString().ToUpper().Equals("Y")).
                        ToArray();

                    cyklus = 0;
                    foreach (DataRow WO in res_WO)
                    {
                        cyklus++;

                        DataRow[] res_RT_exist_WO =
                            Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                            Where(w => w["WO number"].ToString().Trim() == WO["WO_NUMBER"].ToString()).
                            Where(w => w["BO comment"].ToString().Trim() != string.Empty).
                            ToArray();

                        DateTime TIME_END_PLANNED = DateTime.MinValue;
                        try { TIME_END_PLANNED = Convert.ToDateTime(WO["MAX(T2.TIME_END_PLANNED)"].ToString()); } catch { }

                        if (TIME_END_PLANNED == DateTime.MinValue)
                        {
                            #region
                            Comments_sb.Append(WO["MAX(T1.QTY_OPEN)"].ToString());

                            if (res_RT_exist_WO.Length > 0)
                            {
                                Comments_sb.Append(" rework " + res_RT_exist_WO[0]["BO comment"].ToString());
                            }
                            else
                            {
                                Comments_sb.Append(" build ");

                                TIME_END_PLANNED = DateTime.Now;

                                if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                }

                                Comments_sb.Append("/");

                                if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString() == "Sunday")
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(3).DayOfWeek.ToString().Substring(0, 3));
                                }
                                else
                                {
                                    Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                }

                            }
                            #endregion
                        }
                        else
                        {
                            #region
                            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);

                            if (TIME_END_PLANNED < today)
                            {
                                Comments_sb.Append(WO["MAX(T1.QTY_OPEN)"].ToString() + " build in test");
                            }
                            else
                            {
                                Comments_sb.Append(WO["MAX(T1.QTY_OPEN)"].ToString());

                                if (res_RT_exist_WO.Length > 0)
                                {
                                    Comments_sb.Append(" rework " + res_RT_exist_WO[0]["BO comment"].ToString());
                                }
                                else
                                {
                                    Comments_sb.Append(" build ");

                                    if (TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                    {
                                        Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                    }
                                    else
                                    {
                                        Comments_sb.Append(TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString().Substring(0, 3));
                                    }

                                    Comments_sb.Append("/");

                                    if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                    {
                                        Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                    }
                                    else
                                    {
                                        Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                    }
                                }
                            }

                            #endregion
                        }

                        if (cyklus < res_WO.Length) Comments_sb.Append("; ");
                    }
                    #endregion

                    // General_Item_Info
                    #region
                    var res_GII_Customer = Variables.General_Item_Info.AsEnumerable().
                        Where(w => w["ITEM_NUMBER"].ToString() == dr["PN"].ToString()).
                        Select(s => s["CUSTOMER_NUMBER"].ToString()).
                        ToArray();

                    var res_GII = Variables.General_Item_Info.AsEnumerable().
                        Where(w => w["HP_PART_NO"].ToString() == dr["HP PN"].ToString() &&
                             w["CUSTOMER_NUMBER"].ToString() == res_GII_Customer[0] &&
                             w["ITEM_NUMBER"].ToString() != dr["PN"].ToString() &&
                             (dr["OEM"].ToString().Trim().ToUpper() == "SEA" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "HIT" ||
                              dr["OEM"].ToString().Trim().ToUpper() == "MIC" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "HYN" ||
                              dr["OEM"].ToString().Trim().ToUpper() == "INT" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "SAM" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "HGS" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "TOS" || 
                              dr["OEM"].ToString().Trim().ToUpper() == "RRDK")). 
                        Select(s => s["ITEM_NUMBER"].ToString()).
                        ToArray();

                    foreach (var GII in res_GII)
                    {
                        // pcs as alt. revision
                        #region
                        int QTY = 0;

                        try
                        {
                            QTY = Variables.SCM_STOCK_REPORT.AsEnumerable()
                                .Where(w => w["Item Number"].ToString() == GII.ToString())
                                .Sum(s => Convert.ToInt32(s["QAVAILABLE"].ToString()));
                        }
                        catch { }

                        if (QTY > 0)
                        {
                            if (Comments_sb.ToString() != string.Empty) Comments_sb.Append("; ");
                            Comments_sb.Append(QTY + " pcs as alt. revision " + GII);
                        }
                        #endregion

                        // build as alt. revision
                        #region

                        
                        res_WO =
                            Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == GII.ToString() || w["MAX(T1.DESCRIPTION)"].ToString().Trim() == GII.ToString()).
                            //Where(w => w["MAX(T2.TEST_YN)"].ToString().ToUpper().Equals("N")).
                            ToArray();

                        if (res_WO.Length > 0)
                        {
                            foreach (DataRow WO in res_WO)
                            {
                                DataRow[] res_RT_exist_WO =
                                    Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                                    Where(w => w["WO number"].ToString().Trim() == WO["WO_NUMBER"].ToString()).
                                    ToArray();

                                if (res_RT_exist_WO.Length == 0)
                                {
                                    if (Comments_sb.ToString() != string.Empty) Comments_sb.Append("; ");

                                    Comments_sb.Append(WO["MAX(T1.QTY_OPEN)"].ToString());
                                    Comments_sb.Append(" build ");

                                    //vytvor promennou datumu pro WO ktere nemaji end time planned

                                    DateTime TIME_END_PLANNED = DateTime.MinValue;
                                    //WORK_ORDER_QUANTITY == WO_BALANCE je 0 tak RD stock check
                                    //                      


                                    //pokud je vetsi 0 tak +2 dny
                                    //mensi nebo roven 0 tak awaiting RAW DRIVES
                                    //WO_QTY - WO BALANCE > 0 pak +1 den

                                    int status = Convert.ToInt32(WO[5]);
                                    // DateTime.Today()
                                    //if (TIME_END_PLANNED != Convert.ToDateTime(WO["MAX(T2.TIME_END_PLANNED)"].ToString())) 
                                    //{
                                        DateTime dateTime = new DateTime();
                                        dateTime = Convert.ToDateTime(WO["MAX(T2.TIME_END_PLANNED)"].ToString());
                                        

                                    if (25 <= status && status <= 43)

                                    //if FG qty build == 0
                                    //if OH RD = +2 if not = awaiting RD
                                    {
                                        dateTime.AddDays(2);
                                    }

                                    else if(43 <= status && status <= 80) 
                                    //and qty to build > 0
                                    {
                                        dateTime.AddDays(1);                                        
                                    }
                                    //}
                                    

                                    try 
                                    { 
                                        TIME_END_PLANNED = Convert.ToDateTime(WO["MAX(T2.TIME_END_PLANNED)"].ToString()); 
                                    } catch 
                                    { 
                                    
                                    }

                                    if (TIME_END_PLANNED == DateTime.MinValue)
                                    {
                                        #region
                                        TIME_END_PLANNED = DateTime.Now;

                                        if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday")
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        else
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                        }

                                        Comments_sb.Append("/");

                                        if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString() == "Sunday")
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(3).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        else
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region
                                        if (TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        else
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString().Substring(0, 3));
                                        }

                                        Comments_sb.Append("/");

                                        if (TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString() == "Sunday" || TIME_END_PLANNED.AddDays(0).DayOfWeek.ToString() == "Sunday")
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(2).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        else
                                        {
                                            Comments_sb.Append(TIME_END_PLANNED.AddDays(1).DayOfWeek.ToString().Substring(0, 3));
                                        }
                                        #endregion
                                    }

                                    Comments_sb.Append(" as alt. revision " + GII);
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    //REWORK TRACKER open HP PN
                    // Rework Tracker where "HP PN" is the same and "RRD PN" is the same
                    #region
                    DataRow[] res_RT =
                        Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                        Where(w => w["TO HP"].ToString() == dr["HP PN"].ToString()).
                        Where(w => w["TO RRD"].ToString() == dr["PN"].ToString()).
                        Where(w => w["Status"].ToString().ToUpper().Trim() == "OPEN").
                        ToArray();                    
                    //res_RT = opened rework tracker orders
                    cyklus = 0;
                    foreach (DataRow dr_RT in res_RT)
                    {
                        cyklus++;

                        DataRow[] res_exist_in_PROD_WORK_ORDER_SEQUENCE_TAB = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["WO_NUMBER"].ToString() == dr_RT["WO number"].ToString()).
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                            ToArray();

                        if (res_exist_in_PROD_WORK_ORDER_SEQUENCE_TAB.Length == 0)
                        {
                            if (Comments_sb.ToString() != string.Empty) Comments_sb.Append("; ");
                            Comments_sb.Append(dr_RT["Quantity"].ToString() + " rework " + dr_RT["BO comment"].ToString());

                            //DataRow[] res_dr_in_BO = res_dr.AsEnumerable().Where(w => w["HP PN"].ToString() == dr["HP PN"].ToString()).ToArray();

                            //if (res_dr_in_BO.Length > 1)
                            //{
                            //    Comments_sb.Append("                      !!! user control !!! is the same PN");
                            //    continue;
                            //}
                        }
                    }
                    #endregion

                    // Rework Tracker where "HP PN" is the same and "RRD PN" not the same 
                    #region
                    res_RT =
                        Variables.Rework_tracker_WO_2018_NEW_finance.AsEnumerable().
                        Where(w => w["TO HP"].ToString() == dr["HP PN"].ToString()).
                        Where(w => w["TO RRD"].ToString() != dr["PN"].ToString()).
                        Where(w => w["Status"].ToString().ToUpper().Trim() == "OPEN").
                        ToArray();

                    cyklus = 0;
                    foreach (DataRow dr_RT in res_RT)
                    {
                        cyklus++;
                        //description 
                        DataRow[] res_exist_in_PROD_WORK_ORDER_SEQUENCE_TAB = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(w => w["WO_NUMBER"].ToString() == dr_RT["WO number"].ToString()).
                            Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                                       w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
                            ToArray();

                        if (res_exist_in_PROD_WORK_ORDER_SEQUENCE_TAB.Length == 0)
                        {
                            //DataRow[] res_dr_in_BO = res_dr.AsEnumerable().Where(w => w["HP PN"].ToString() == dr["HP PN"].ToString()).ToArray();

                            //if (res_dr_in_BO.Length > 1)
                            //{
                            //    Comments_sb.Append("                      !!! user control !!!");
                            //    continue;
                            //}

                            if (Comments_sb.ToString() != string.Empty) Comments_sb.Append("; ");
                            Comments_sb.Append(dr_RT["Quantity"].ToString() + " rework " + dr_RT["BO comment"].ToString());
                        }
                    }
                    #endregion

                    // qty on Hold
                    #region
                    int QONHAND = 0;
                    try
                    {
                        QONHAND = Variables.SCM_STOCK_REPORT.AsEnumerable().
                            Where(w => w["Item Number"].ToString() == dr["PN"].ToString()).
                            Where(w => w["LSTTS"].ToString().Trim() != string.Empty).
                            Where(w => w["LSTTS"].ToString().Trim() != "A").
                            Where(w => w["LSTTS"].ToString().Trim() != "H").
                            Where(w => w["LSTTS"].ToString().Trim() != "X").
                            Where(w => w["LSTTS"].ToString().Trim() != "1").

                            Sum(s => Convert.ToInt32(s["QONHAND"].ToString()));
                    }
                    catch 
                    { 
                    }

                    if (QONHAND > 0)
                    {
                        if (Comments_sb.ToString() != string.Empty) Comments_sb.Append("; ");
                        Comments_sb.Append(QONHAND + " pcs on Hold");
                    }
                    #endregion

                    // no build scheduled
                    #region
                    DataRow[] res_empty_comment = Variables.DT.AsEnumerable().Where(w => w["Comments"].ToString().Trim() == string.Empty).ToArray();

                    foreach (DataRow dr_empty_comment in res_empty_comment)
                    {
                        dr_empty_comment["Comments"] = "Manual control is required !!!";
                    }
                    #endregion
                    // no build scheduled update
                    if(Comments_sb.ToString().Contains("no build scheduled") && Comments_sb.ToString().Length > 18) 
                    {
                        try
                        {
                            Comments_sb = Comments_sb.Replace("no build scheduled; ", "");
                        }
                        catch 
                        {
                            Comments_sb = Comments_sb.Replace("no build scheduled", "");                        
                        }
                    }
                    dr["Comments"] = (Comments_sb.ToString() == string.Empty ? null : Comments_sb.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString() + Environment.NewLine + Environment.NewLine +
                    "V reportu: General_Item_Info " +
                    "Neexistuje Item: " + drtmp["PN"].ToString() +
                    ", nebo HP PN: " + drtmp["HP PN"].ToString() + Environment.NewLine +
                    "Znamená to, že pro tento Item nebudou zjištěny údaje pro Alternativní Revize", "Back_Order_Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Print_WO_Alternatives_To_Excel()
        {
            //try
            //{
            foreach (DataRow dr in Variables.DT.Rows)
            {
                string[] res_item = dr["General_Item_Info_AR_all_AB"].ToString().Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in res_item.Distinct())
                {
                    string g = String.Join(",", res_item.Distinct().ToArray());
                    var sel = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable()
                        .Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == item.ToString() &&
                                    item.ToString() != dr["PN"].ToString())
                        .ToList();

                    DataRow[] res_WO_does_not_exist_in_the_table =
               Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
               Where(w => w["MAX(T1.ITEM_NUMBER)"].ToString() == dr["PN"].ToString() || 
                          w["MAX(T1.DESCRIPTION)"].ToString().Trim() == dr["PN"].ToString()).
               ToArray();

                }
            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace.ToString());
            //}

        }


        //#####################################################################################################################################################
        private void Writing_to_Excel_Comment()

        {
            //-----------------------------------------------------------------------------------
            xlApp = new EXCEL.Application
            {
                Visible = false
            };
            //-----------------------------------------------------------------------------------
            xlWb = xlApp.Workbooks.Open(BAck_Order_File_Name);
            xlWs = xlWb.Worksheets["Summary"];
            //-----------------------------------------------------------------------------------
            int i = 1;
            while (xlWs.Range["A" + i].Value != null)
            {
                string[] res_Comments = Variables.DT.AsEnumerable().
                    Where(w => w["PN"].ToString() == xlWs.Range["A" + i].Value.ToString()).
                    Select(s => s["Comments"].ToString()).
                    ToArray();

                if (res_Comments.Length > 0) xlWs.Range["E" + i].Value = res_Comments[0];
                i++;
            }

            System.Timers.Timer ttimer = new System.Timers.Timer();

            ttimer.Start();
            DateTime startTime = DateTime.Now;

            ExcelConnector er = new ExcelConnector(@"r:\Operations\Warehouse\Shipping\Backorder Reports\HPE commit report.xlsx", true);
            DataTable hpe = er.Select("SELECT * FROM [Main report$]");

            DataTable filtered = hpe.AsEnumerable().Where(a => a["HPE date"].ToString() != "#N/A").Where(b => b["HPE date"].ToString() != "").CopyToDataTable();
            ttimer.Stop();
            DateTime endTime = DateTime.Now;
            TimeSpan totaltime = endTime - startTime;



            EXCEL.Range start = xlWs.Range["A2"];
            int lastRow = xlWs.Range["A2"].get_End(EXCEL.XlDirection.xlDown).Row;
            EXCEL.Range end = xlWs.Range["A" + lastRow];
            EXCEL.Range rangeOfPN = xlWs.get_Range(start, end);

            object[] PNs = new object[rangeOfPN.Rows.Count];

            for (int o = 1; o <= rangeOfPN.Rows.Count; o++)
            {
                PNs[o - 1] = xlWs.Range["A" + (o + 1)].Value;
            }

            string PreviousPn = PNs.Last().ToString();


            xlWsBO = xlWb.Worksheets["BO"];
            int usedrn = xlWsBO.Range["C" + xlWsBO.Rows.Count + ""].get_End(EXCEL.XlDirection.xlUp).Row;

            List<string> Nobuild = Variables.DT.AsEnumerable().Where(a => a["Comments"].ToString().Contains("no build scheduled")).Select(a => a["PN"].ToString()).ToList();
            

            string PreviousPN = string.Empty;
            double left = 0;
            double ItemCount = 0;
            int indexNumber = 0;
            string CurrentPN = string.Empty;
            string CurrentHP = string.Empty;
            TimeSpan timeSpan = new TimeSpan();
            String format = "ddd";
            String formatddmmyyyy = "MM/dd/yyyy hh:mm:ss";

            DateTime from = new DateTime();
            DateTime to = new DateTime();


            xlWsBO.Range["AD:AD"].EntireColumn.NumberFormat = "MM/dd/yyyy";

            double piecesforActualProdWO = 0;

            for (int t = 2; t <= usedrn; t++)
            {
                try
                {
                    //if cell empty or null or special value "#VALUE!" skip
                    if (xlWsBO.Range["A" + t].Value == null
                        || xlWsBO.Range["A" + t].Value.ToString() == "PN"
                        || xlWsBO.Range["A" + t].Value.ToString() == "0"
                        || xlWsBO.Range["A" + t].Value.ToString() == string.Empty
                        || xlWsBO.Range["W" + t].Value.ToString() == null
                        || xlWsBO.Range["A" + t].Value.ToString() == "#VALUE!")
                    {
                        indexNumber = 0;
                        ItemCount = 0;
                        continue;
                    }

                    CurrentPN = xlWsBO.Range["A" + t].Value.ToString();
                    CurrentHP = xlWsBO.Range["S" + t].Value.ToString();

                    IEnumerable<DataRow> varss = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().Where(a => Nobuild.Any(b => b.ToString() == a["MAX(T1.ITEM_NUMBER)"].ToString())).Select(a => a).AsEnumerable();
                    var idiot = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().Where(a => Nobuild.Any(b => b.ToString() == a["MAX(T1.ITEM_NUMBER)"].ToString())).Select(a => a).ToList();
                    var heyidiot = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().Where(a => Nobuild.Any(b => b.ToString() == a["MAX(T1.ITEM_NUMBER)"].ToString())).Select(a => a).AsEnumerable();
                    //Datetime = xlWsBO.Range[""                  xlWsBO.Range["AE" + t].Value.ToString();
                    // Kuskova odstranit nobuild scheduled
                    //

                    //if (xlWsBO.Range["K" + t].Value.ToString().Contains("no build scheduled") && xlWsBO.Range["K" + t].Value.ToString().Length > 18)
                    //{
                    //    try
                    //    {
                    //        xlWsBO.Range["K" + t].Value = xlWsBO.Range["K" + t].Value.ToString().Replace("no build scheduled; ", "");
                    //    }
                    //    catch
                    //    {
                    //        xlWsBO.Range["K" + t].Value = xlWsBO.Range["K" + t].Value.ToString().Replace("no build scheduled", "");
                    //    }
                    //}

                    //xlWsBO.Range["K"+t].Value = xlWsBO.Range["K" + t].Value.ToString().

                    // if nobuild scheduled skip ItemNumber
                    if (CurrentPN.Where(b => Nobuild.Any(a => a.ToString() == CurrentPN)).ToList().Count() == 0 &&
                        filtered.AsEnumerable().Where(a=>a["Part Number"].ToString() == CurrentHP).ToList().Count == 0
                        )
                     {
                        xlWsBO.Range["AC" + t + ""].Value = "N";
                        xlWsBO.Range["AB" + t + ""].Value = "N/A";
                        PreviousPN = CurrentPN;
                        ItemCount = 0;
                        continue;
                    }

                    double currentBOReportQuantityOrdered = xlWsBO.Range["T" + t].Value;

                    // if there is a build with comment no build scheduled then skip
                    var Times = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"] != DBNull.Value).
                        Select(a => a["MAX(T2.TIME_END_PLANNED)"]).ToList();

                    // if there is no plan 
                    if (Times == null || Times.Count == 0)
                    {
                        // search in another excel file "HPE commit report" to search for dates there
                        var dates = filtered.AsEnumerable().Where(a => a["Part Number"].ToString() == xlWsBO.Range["S" + t].Value.ToString()).ToList();
                        // if there is a date 
                        if (dates.Count > 0)
                        {
                            // select the date and create time span
                            var hpe_commit_report_date = filtered.AsEnumerable().Where(a => a["Part Number"].ToString() == xlWsBO.Range["S" + t].Value.ToString()).Select(a => a["HPE date"]).First();
                            DateTime hpetime = DateTime.Parse(hpe_commit_report_date.ToString());
                            DateTime UserDate = xlWsBO.Range["W" + t + ""].Value;
                            TimeSpan tss = hpetime - UserDate;
                            if (tss < TimeSpan.Zero) 
                            {
                                xlWsBO.Range["AD" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                                xlWsBO.Range["AD" + t + ""].Value = hpetime;
                                xlWsBO.Range["AB" + t + ""].Value = "N/A";
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                PreviousPN = CurrentPN;
                                continue;
                            }
                            else if(tss > TimeSpan.Zero)
                            {
                                xlWsBO.Range["AD" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbGreen;
                                xlWsBO.Range["AD" + t + ""].Value = hpetime;
                                xlWsBO.Range["AB" + t + ""].Value = "N/A";
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                PreviousPN = CurrentPN;
                                continue;
                            }
                            else 
                            {
                                xlWsBO.Range["AB" + t + ""].Value = "no build scheduled";
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                PreviousPN = CurrentPN;
                                continue;
                            }

                        }
                        xlWsBO.Range["AB" + t + ""].Value = "no build scheduled";
                        xlWsBO.Range["AC" + t + ""].Value = "N";
                        PreviousPN = CurrentPN;
                        continue;
                    }

                    // if the current WO is to be tested then add two extra days to MAX(T2.TIME_END_PLANNED) value
                    bool test = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable()
                        .Where(a => a["MAX(T2.TEST_YN)"].ToString() == "Y" && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).ToList().Count() > 0;

                   
                    //start a new PN
                    if (CurrentPN != PreviousPN)
                    {
                        indexNumber = 0;
                        ItemCount = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable()
                        .Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN)
                        .Sum(b => Convert.ToDouble(b["MAX(T1.QTY_OPEN)"]));

                        var Time1 = xlWsBO.Range["W" + t].Value.ToString();
                        from = DateTime.ParseExact(Time1, "dd/MM/yyyy hh:mm:ss", null);

                        DateTime TimeforActualProdWO = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                        Select(a => a["MAX(T2.TIME_END_PLANNED)"]).Cast<DateTime>().First();

                        var WONumber = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                        Select(a => a["WO_NUMBER"]).First();

                        piecesforActualProdWO = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                        Select(a => Convert.ToDouble(a["MAX(T1.QTY_OPEN)"])).First();

                        piecesforActualProdWO -= currentBOReportQuantityOrdered;

                        if (piecesforActualProdWO < 0 && indexNumber == (Times.Count() - 1))
                        {
                            xlWsBO.Range["AB" + t + ""].Value = "N/A";
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                            PreviousPN = CurrentPN;
                            continue;
                        }

                        if (test)
                        {
                            TimeforActualProdWO = TimeforActualProdWO.AddDays(2);
                        }

                        timeSpan = TimeforActualProdWO - from;
                        if (timeSpan < TimeSpan.Zero)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;

                            //Add 1 day for a packing and warehouse stuff
                            //if that day is to be sunday add another one

                            if (TimeforActualProdWO.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(2).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                            else
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                xlWsBO.Range["AD" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }

                        }
                        else if (timeSpan > TimeSpan.Zero)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbGreen;
                            if (TimeforActualProdWO.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(2).ToString(format).Substring(0, 3).ToUpper() + " OK " + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "Y";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                            else
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + " OK " + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "Y";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                        }
                        else
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                            xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbYellow;
                            xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                        }

                        // if pieces are less than zero skip to another WO
                        if (piecesforActualProdWO <= 0 && indexNumber < (Times.Count() - 1))
                        {
                            indexNumber++;

                            piecesforActualProdWO += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                            Select(a => Convert.ToDouble(a["MAX(T1.QTY_OPEN)"])).First();

                            // if there is still negative amount increment pcs from next WO
                            if (piecesforActualProdWO < 0)
                            {
                                while (piecesforActualProdWO >= 0 ^ indexNumber < (Times.Count() - 1))
                                {
                                    indexNumber++;
                                    piecesforActualProdWO += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                                    Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                                    Select(a => Convert.ToDouble(a["MAX(T1.QTY_OPEN)"])).First();
                                }
                            }
                        }
                        else if (piecesforActualProdWO < 0)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                        }

                    }
                    // or continue using same PN
                    else
                    {


                        piecesforActualProdWO -= currentBOReportQuantityOrdered;
                        if (piecesforActualProdWO < 0 && indexNumber == (Times.Count() - 1))
                        {
                            xlWsBO.Range["AB" + t + ""].Value = "N/A";
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                            PreviousPN = CurrentPN;
                            continue;
                        }

                        var Time1 = xlWsBO.Range["W" + t].Value.ToString();
                        from = DateTime.ParseExact(Time1, "dd/MM/yyyy hh:mm:ss", null);

                        DateTime TimeforActualProdWO = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                        Select(a => a["MAX(T2.TIME_END_PLANNED)"]).Cast<DateTime>().First();

                        var WONumber = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                        Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                        Select(a => a["WO_NUMBER"]).First();


                        if (test)
                        {
                            TimeforActualProdWO = TimeforActualProdWO.AddDays(2);
                        }

                        timeSpan = TimeforActualProdWO - from;

                        if (timeSpan < TimeSpan.Zero)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                            if (TimeforActualProdWO.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(2).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                            else
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "N";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                        }
                        else if (timeSpan > TimeSpan.Zero)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbGreen;
                            if (TimeforActualProdWO.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(2).ToString(format).Substring(0, 3).ToUpper() + " OK " + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "Y";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                            else
                            {
                                xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + " OK " + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                                xlWsBO.Range["AC" + t + ""].Value = "Y";
                                xlWsBO.Range["Z" + t + ""].Value = TimeforActualProdWO.ToOADate();
                            }
                        }
                        else
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                            xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbYellow;
                            xlWsBO.Range["AB" + t + ""].Value = TimeforActualProdWO.AddDays(1).ToString(format).Substring(0, 3).ToUpper() + TimeforActualProdWO.ToString().Substring(0, 10) + " WO " + WONumber.ToString() + " PCS " + piecesforActualProdWO;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                        }

                        // if pieces are less than zero skip to another WO
                        if (piecesforActualProdWO <= 0 && indexNumber < (Times.Count() - 1))
                        {
                            indexNumber++;

                            piecesforActualProdWO += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                            Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                            Select(a => Convert.ToDouble(a["MAX(T1.QTY_OPEN)"])).First();

                            // if there is still negative amount increment pcs from next WO
                            if (piecesforActualProdWO < 0)
                            {
                                while (piecesforActualProdWO >= 0 ^ indexNumber < (Times.Count() - 1))
                                {
                                    indexNumber++;
                                    piecesforActualProdWO += Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable().
                                    Where(a => a["MAX(T1.ITEM_NUMBER)"].ToString() == CurrentPN && a["MAX(T2.TIME_END_PLANNED)"].ToString().Contains(Times[indexNumber].ToString())).
                                    Select(a => Convert.ToDouble(a["MAX(T1.QTY_OPEN)"])).First();
                                }
                            }
                        }
                        else if (piecesforActualProdWO < 0)
                        {
                            xlWsBO.Range["AB" + t + ""].Font.Color = EXCEL.XlRgbColor.rgbRed;
                            xlWsBO.Range["AC" + t + ""].Value = "N";
                        }

                    }

                    PreviousPN = CurrentPN;

                }
                catch (InvalidOperationException ex)
                {
                    xlWsBO.Range["AB" + t + ""].Value = "N/A";
                    xlWsBO.Range["AC" + t + ""].Value = "N";
                    PreviousPN = CurrentPN;
                    continue;
                }
                catch (RuntimeBinderException ex)
                {
                    xlWsBO.Range["AB" + t + ""].Value = "N/A";
                    xlWsBO.Range["AC" + t + ""].Value = "N";
                    PreviousPN = CurrentPN;
                    continue;
                }
                xlWsBO.Range["Z:Z"].EntireColumn.NumberFormat = "dd/MM/yyyy";
                xlWsBO.Range["AD:AD"].EntireColumn.NumberFormat = "dd/MM/yyyy";
            }


            xlWsBO.Range["Z1"].Value = "Presumed date";
            xlWsBO.Range["AB1"].Value = "WO Completed on";
            xlWsBO.Range["AC1"].Value = "Is Covered";
            xlWsBO.Range["AD1"].Value = "HPE Date";

            i = 0;

            string tmpValue = string.Empty;
            string[] AllAdressBookAlternatives;


            DataSet AlternativesTable = new DataSet();

            string AlternativesTableCmd = "SELECT * FROM RPTDBLINK.DIRECT_SHIP_FG_WITH_RD";

            using (OracleDataAdapter daOra = new OracleDataAdapter(AlternativesTableCmd, Variables.OraConn_DBS2))
            {
                daOra.Fill(AlternativesTable);
            }
            
            string[] drives_on_stock;

            List<string> fgs = new List<string>();
            List<string> hp = new List<string>();
            List<string> RDQTY = new List<string>();

            Dictionary<string, List<string>> AlternativesItemNumbers = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> AlternativesHPPN = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> AlternativesALTRDQTY = new Dictionary<string, List<string>>();
            
            try
            {

                while (xlWs.Range["A" + (i + 2)].Value != null)
                {

                    if (tmpValue == xlWs.Range["A" + (i + 2)].Value.ToString())
                    {
                        i++;
                        continue;
                    }
                    {
                        List<string> alternativeFG = AlternativesTable.Tables[0].AsEnumerable()
                            .Where(excelAlternatives => excelAlternatives["FG_HPE_PN"].ToString() == Variables.DT.Rows[i]["HP PN"].ToString())
                            .Select(excelAlternatives => excelAlternatives["RD"].ToString())
                            .ToList();


                        List<string> alternativeHPE = AlternativesTable.Tables[0].AsEnumerable()
                            .Where(excelAlternatives => excelAlternatives["FG_HPE_PN"].ToString() == Variables.DT.Rows[i]["HP PN"].ToString())
                            .Select(excelAlternatives => excelAlternatives["FG_HPE_PN"].ToString())
                            .ToList();

                        List<string> ALTRDQTY = AlternativesTable.Tables[0].AsEnumerable()
                            .Where(excelAlternatives => excelAlternatives["FG_HPE_PN"].ToString() == Variables.DT.Rows[i]["HP PN"].ToString())
                            .Select(excelAlternatives => excelAlternatives["FG_RRD_PN"].ToString())
                            .ToList();

                        hp.AddRange(alternativeHPE.Distinct());
                        fgs.AddRange(alternativeFG);
                        fgs.AddRange(ALTRDQTY);
                        RDQTY.AddRange(ALTRDQTY);

                        AlternativesItemNumbers.Add(xlWs.Range["A" + (i + 2)].Value.ToString(), alternativeFG);
                        AlternativesHPPN.Add(xlWs.Range["A" + (i + 2)].Value.ToString(), alternativeHPE.Distinct().ToList());
                        AlternativesALTRDQTY.Add(xlWs.Range["A" + (i + 2)].Value.ToString(), ALTRDQTY.Distinct().ToList());

                        tmpValue = xlWs.Range["A" + (i + 2)].Value.ToString();
                        i++;

                    }
                }
            }
            catch (IndexOutOfRangeException ex)
            {

            }

            Variables.SCM_STOCK_REPORT.Clear();


            string cmd = "SELECT * FROM RPTDBLINK.SCM_STOCK_REPORT WHERE \"Item Number\" IN ('" + string.Join("', '", fgs.Distinct()) + "') and branch = '728'";

            using (OracleDataAdapter daOra = new OracleDataAdapter(cmd, Variables.OraConn_DBS2))
            {
                daOra.Fill(Variables.SCM_STOCK_REPORT);
            }

            Dictionary<string, double> qtyOfAlternatives = new Dictionary<string, double>();
            List<List<string>> WoS = new List<List<string>>();
            
            int j = 0;
            int k = 0;
            foreach (var item in AlternativesItemNumbers.Values)
            {
                double count = 0;
                double tmpCount = 0;
                for (j = 0; j < item.Count; j++)
                {
                    tmpCount = Variables.SCM_STOCK_REPORT.AsEnumerable().Where(a => a["Item Number"].ToString() == item[j].ToString())
                        .Sum(a => Convert.ToDouble(a.ItemArray[28].ToString()));

                    var WOplanned = Variables.PROD_WORK_ORDER_SEQUENCE_TAB.AsEnumerable()
                        .Where(a => a.ItemArray[3].ToString() == item[j].ToString())
                        .ToList();

                    if (WOplanned.Count > 0)
                    {

                    }

                    //.Select(a=>a.ItemArray[3].ToString()).ToString();
                    if (tmpCount < 0)
                    {
                        continue;
                    }
                    else
                    {
                        count += tmpCount;
                    }

                    //WoS.Add(WOplanned);
                }
                qtyOfAlternatives.Add(AlternativesItemNumbers.ElementAt(k).Key.ToString(), count);
                k++;
            }

            DataSet ds = new DataSet();
            string command = "select * from brnodata.prod_work_order_sequence_tab where WORK_ORDER_STATUS not in (95,96,98,99) AND BRANCH_PLANT = '         728'  and TIME_START_PLANNED is not null and HPK_PART_NUMBER IN ('" + string.Join("', '", hp) + "')";

            using (OracleDataAdapter daOra = new OracleDataAdapter(command, Variables.OraConn_DBS2))
            {
                daOra.Fill(ds);
            }

            int p = 2;
            for (int q = 0; q < AlternativesHPPN.Keys.Count; q++)
            {

                foreach (var item in AlternativesHPPN[xlWs.Range["A" + p].Value.ToString()])
                {

                    var WOSplanned = ds.Tables[0].AsEnumerable().Where(a => a.ItemArray[4].ToString() == item.ToString()).Where(a => a.ItemArray[3].ToString() != xlWs.Range["A" + p].Value.ToString())
                        .Select(a => new { WO = a.ItemArray[13].ToString(), WOQTY = a.ItemArray[16].ToString(), Item = a.ItemArray[3].ToString(), Planned = a.ItemArray[34].ToString() }).ToList();

                    if (WOSplanned.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        for (int r = 0; r < WOSplanned.Count; r++)
                        {
                            sb.Append("" + WOSplanned[r].WO + " - " + WOSplanned[r].Item + " - " + WOSplanned[r].WOQTY + "");
                        }
                        xlWs.Range["G" + (q + 2)].Value = sb.ToString();
                    }
                }
                p++;
            }
            
            int it = 0;

            string CurrentBOPN = string.Empty;
            int SummaryPN = 2;
            int BoPN = 2;
            xlWs.Range["F1"].Value = "ALT FG SUM";
            xlWs.Range["G1"].Value = "ALT FG WO-PN-QTY";
            xlWs.Range["H1"].Value = "ALT RD QTY";
            xlWs.Range["I1"].Value = "ALT FG QTY PN";
            xlWs.Range["J1"].Value = "ALT RD QTY";
            try
            {
                while (xlWs.Range["A" + SummaryPN].Value != null)
                {
                    double cnt = 0;
                    var Alts2 = AlternativesALTRDQTY.Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString())
                        .Select(a => a.Value)
                        .ToList();

                    cnt = Variables.SCM_STOCK_REPORT.AsEnumerable()
                        .Where(a => Alts2[0].Any(b => b.ToString() == a["Item Number"].ToString()))
                        .Where(a => Convert.ToDouble(a["QAVAILABLE"]) > 0)
                        .Sum(a => Convert.ToDouble(a["QAVAILABLE"]));

                    string QtytoWrite = qtyOfAlternatives.Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString()).Select(a => a.Value.ToString()).First().ToString();

                    //AlternativesALTRDQTY.Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString())

                    //xlWs.Range["H" + SummaryPN].Value = cnt;
                    //xlWs.Range["F" + SummaryPN].Value = QtytoWrite;

                    xlWs.Range["H" + SummaryPN].Value = QtytoWrite;
                    xlWs.Range["F" + SummaryPN].Value = cnt;

                    var d = AlternativesALTRDQTY.AsEnumerable()
                        .Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString())
                        .Select(a => a.Value).ToList();

                    string out1 = string.Join(",", d[0].ToArray());

                    var de = AlternativesItemNumbers.AsEnumerable()
                        .Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString())
                        .Select(a => a.Value).ToList();

                    string out2 = string.Join(",", de[0].ToArray());

                    //xlWs.Range["I" + SummaryPN].Value = AlternativesALTRDQTY.AsEnumerable()
                    //    .Where(a => a.Key.ToString() == xlWs.Range["A" + SummaryPN].Value.ToString())
                    //    .Select(a=>a.Value.ToList()).ToString();

                    //string Alts = 
                    //xlWs.Range["I" + SummaryPN].Value = out1 + " -- " + QtytoWrite + "PCS ;";
                    //xlWs.Range["J" + SummaryPN].Value = out2 + " --  " + xlWs.Range["F" + SummaryPN].Value + " PCS ;";

                    xlWs.Range["I" + SummaryPN].Value = out1 + " -- " + xlWs.Range["F" + SummaryPN].Value + "PCS ;";
                    xlWs.Range["J" + SummaryPN].Value = out2 + " --  " + QtytoWrite + " PCS ;";
                    SummaryPN++;
                }
            }
            catch (IndexOutOfRangeException ex)
            {

            }


            xlWsBO.Columns.AutoFit();
            xlWs.Columns.AutoFit();
            xlWsBO.Range["Z1:AD"+ usedrn + ""].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbCornsilk;
            xlWs.Range["F1:J" + i + ""].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbCornsilk;

            xlApp.DisplayAlerts = false;

            xlWs.SaveAs(BAck_Order_File_Name);
            xlWs.SaveAs(Resources.Patch_Backorder_Reports_Comments_new);

            FreeWorkSheetResources();
            FreeWorkBookResources();
            FreeApplicationResources();
        }
        private void Writing_to_Excel_Details()
        {
            //-----------------------------------------------------------------------------------
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Back Order Comment Details     " + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";

            dataGridView_Data.SelectAll();
            dataGridView_Data.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DataObject dataObj = dataGridView_Data.GetClipboardContent();

            if (dataObj != null)
            {
                Clipboard.Clear();
                Clipboard.SetDataObject(dataObj);
            }

            dataGridView_Data.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            //-----------------------------------------------------------------------------------
            xlApp = new EXCEL.Application();
            xlApp.Visible = false;
            //-----------------------------------------------------------------------------------
            xlWb = xlApp.Workbooks.Add(Missing.Value);
            xlWs = xlWb.Worksheets[1];
            xlWs.Name = "Back Order Comment Details";
            //-----------------------------------------------------------------------------------
            DataGridViewColumn[] dgvcol_Arr = dataGridView_Data.Columns.Cast<DataGridViewColumn>().
                Where(w => w.Visible).
                OrderBy(ob => ob.DisplayIndex).
                ToArray();

            if (dgvcol_Arr.Length == 0)
            {
                FreeWorkSheetResources();
                FreeWorkBookResources();
                FreeApplicationResources();

                return;
            }

            xlRx = xlWs.Cells[1, 1];
            xlRy = xlWs.Cells[dataGridView_Data.Rows.Count + 1, dgvcol_Arr.Length + 1];
            xlRfr = xlWs.Cells[1, dgvcol_Arr.Length];
            //-----------------------------------------------------------------------------------
            xlRx.Select();
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (dataObj == null) return;
            xlRx.PasteSpecial(EXCEL.XlPasteType.xlPasteAll, EXCEL.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
            //---------------------------------------------------------------------------------------------------------------------------------------
            FormatAsTable(xlWs.get_Range(xlRx, xlRy), filename, "TableStyleMedium16");
            //---------------------------------------------------------------------------------------------------------------------------------------
            xlWs.get_Range(xlRx, xlRy).Select();
            xlWs.get_Range(xlRx, xlRy).Style.HorizontalAlignment = EXCEL.XlHAlign.xlHAlignCenter;
            //---------------------------------------------------------------------------------------------------------------------------------------
            EXCEL.Range oRange = xlWs.get_Range(xlRx, xlRy);
            oRange.Borders.get_Item(EXCEL.XlBordersIndex.xlEdgeLeft).LineStyle = EXCEL.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(EXCEL.XlBordersIndex.xlEdgeRight).LineStyle = EXCEL.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(EXCEL.XlBordersIndex.xlInsideHorizontal).LineStyle = EXCEL.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(EXCEL.XlBordersIndex.xlInsideVertical).LineStyle = EXCEL.XlLineStyle.xlContinuous;
            oRange.Borders.Color = Color.Black;
            oRange.Font.Size = 8;
            //---------------------------------------------------------------------------------------------------------------------------------------
            EXCEL.Range oRangeFR = xlWs.get_Range(xlRx, xlRfr);
            oRangeFR.Font.Size = 9;
            //---------------------------------------------------------------------------------------------------------------------------------------
            xlWs.Columns.AutoFit();
            xlWs.Rows.AutoFit();
            //---------------------------------------------------------------------------------------------------------------------------------------
            xlApp.DisplayAlerts = false;

            xlWb.SaveAs(
                filename, EXCEL.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                EXCEL.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            FreeWorkSheetResources();
            FreeWorkBookResources();
            FreeApplicationResources();
        }
        public void FormatAsTable(EXCEL.Range SourceRange, string TableName, string TableStyleName)
        {
            SourceRange.Worksheet.ListObjects.Add(EXCEL.XlListObjectSourceType.xlSrcRange,
            SourceRange, Type.Missing, EXCEL.XlYesNoGuess.xlYes, Type.Missing).Name = TableName;
            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        }



        private void FreeWorkSheetResources()
        {
            Marshal.ReleaseComObject(xlWs);
            xlWs = null;
        }
        private void FreeWorkBookResources()
        {
            if (xlWs != null)
            {
                Marshal.ReleaseComObject(xlWs);
                xlWs = null;
                xlWb.Close(false);
            }

            Marshal.ReleaseComObject(xlWb);
            xlWb = null;
        }
        private void FreeApplicationResources()
        {
            xlApp.DisplayAlerts = false;
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            GC.Collect();
        }
        //#####################################################################################################################################################
    }

    public class XML_Read_Write
    {
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_BO_form;
        private FileInfo FI_conditional_formatting_menu_BO_form;
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB;
        private FileInfo FI_conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB;
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_3PAR_Actual_Status;
        private FileInfo FI_conditional_formatting_menu_3PAR_Actual_Status;
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_General_Item_Info;
        private FileInfo FI_conditional_formatting_menu_General_Item_Info;
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_SCM_STOCK_REPORT;
        private FileInfo FI_conditional_formatting_menu_SCM_STOCK_REPORT;
        //---------------------------------------------------------------------------
        private FileInfo FI_setting_column_Rework_tracker;
        private FileInfo FI_conditional_formatting_menu_Rework_tracker;
        //---------------------------------------------------------------------------

        public XML_Read_Write()
        {
            string PathDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString();

            bool exists = Directory.Exists(PathDir);

            if (!exists)
                Directory.CreateDirectory(PathDir);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_BO_form =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_BO_form);
            FI_conditional_formatting_menu_BO_form =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_BO_form);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB);

            FI_conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_PROD_WORK_ORDER_SEQUENCE_TAB);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_3PAR_Actual_Status =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_3PAR_Actual_Status);

            FI_conditional_formatting_menu_3PAR_Actual_Status =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_3PAR_Actual_Status);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_General_Item_Info =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_General_Item_Info);

            FI_conditional_formatting_menu_General_Item_Info =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_General_Item_Info);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_SCM_STOCK_REPORT =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_SCM_STOCK_REPORT);

            FI_conditional_formatting_menu_SCM_STOCK_REPORT =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_SCM_STOCK_REPORT);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
            FI_setting_column_Rework_tracker =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_setting_column_Rework_Tracker);

            FI_conditional_formatting_menu_Rework_tracker =
                new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\.Net_Application_Setting_" + GetType().Namespace.ToString() + Resources.Patch_conditional_formatting_Rework_Tracker);
            //-------------------------------------------------------------------------------------------------------------------------------------------------------
        }
        public void Write_xml(DataTable table, String FI_Name)
        {
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_BO_form")
                table.WriteXml(FI_setting_column_BO_form.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_BO_form")
                table.WriteXml(FI_conditional_formatting_menu_BO_form.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB")
                table.WriteXml(FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB")
                table.WriteXml(FI_conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_3PAR_Actual_Status")
                table.WriteXml(FI_setting_column_3PAR_Actual_Status.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_3PAR_Actual_Status")
                table.WriteXml(FI_conditional_formatting_menu_3PAR_Actual_Status.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_General_Item_Info")
                table.WriteXml(FI_setting_column_General_Item_Info.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_General_Item_Info")
                table.WriteXml(FI_conditional_formatting_menu_General_Item_Info.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_SCM_STOCK_REPORT")
                table.WriteXml(FI_setting_column_SCM_STOCK_REPORT.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_SCM_STOCK_REPORT")
                table.WriteXml(FI_conditional_formatting_menu_SCM_STOCK_REPORT.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
            if (FI_Name == "FI_setting_column_Rework_tracker_WO_2018_NEW_finance")
                table.WriteXml(FI_setting_column_Rework_tracker.FullName, XmlWriteMode.WriteSchema, true);

            if (FI_Name == "conditional_formatting_menu_Rework_tracker_WO_2018_NEW_finance")
                table.WriteXml(FI_conditional_formatting_menu_Rework_tracker.FullName, XmlWriteMode.WriteSchema, true);
            //------------------------------------------------------------------------------------------
        }
        public DataTable Read_xml(String name)
        {
            DataTable table = new DataTable();
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_BO_form")
            {
                if (FI_setting_column_BO_form.Exists == false) return null;
                table.ReadXml(FI_setting_column_BO_form.FullName);
            }
            if (name == "conditional_formatting_menu_BO_form")
            {
                if (FI_conditional_formatting_menu_BO_form.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_BO_form.FullName);
            }
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB")
            {
                if (FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB.Exists == false) return null;
                table.ReadXml(FI_setting_column_PROD_WORK_ORDER_SEQUENCE_TAB.FullName);
            }
            if (name == "conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB")
            {
                if (FI_conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_PROD_WORK_ORDER_SEQUENCE_TAB.FullName);
            }
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_3PAR_Actual_Status")
            {
                if (FI_setting_column_3PAR_Actual_Status.Exists == false) return null;
                table.ReadXml(FI_setting_column_3PAR_Actual_Status.FullName);
            }
            if (name == "conditional_formatting_menu_3PAR_Actual_Status")
            {
                if (FI_conditional_formatting_menu_3PAR_Actual_Status.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_3PAR_Actual_Status.FullName);
            }
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_General_Item_Info")
            {
                if (FI_setting_column_General_Item_Info.Exists == false) return null;
                table.ReadXml(FI_setting_column_General_Item_Info.FullName);
            }
            if (name == "conditional_formatting_menu_General_Item_Info")
            {
                if (FI_conditional_formatting_menu_General_Item_Info.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_General_Item_Info.FullName);
            }
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_SCM_STOCK_REPORT")
            {
                if (FI_setting_column_SCM_STOCK_REPORT.Exists == false) return null;
                table.ReadXml(FI_setting_column_SCM_STOCK_REPORT.FullName);
            }
            if (name == "conditional_formatting_menu_SCM_STOCK_REPORT")
            {
                if (FI_conditional_formatting_menu_SCM_STOCK_REPORT.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_SCM_STOCK_REPORT.FullName);
            }
            //------------------------------------------------------------------------------------------
            if (name == "FI_setting_column_Rework_tracker_WO_2018_NEW_finance")
            {
                if (FI_setting_column_Rework_tracker.Exists == false) return null;
                table.ReadXml(FI_setting_column_Rework_tracker.FullName);
            }
            if (name == "conditional_formatting_menu_Rework_tracker_WO_2018_NEW_finance")
            {
                if (FI_conditional_formatting_menu_Rework_tracker.Exists == false) return null;
                table.ReadXml(FI_conditional_formatting_menu_Rework_tracker.FullName);
            }
            //------------------------------------------------------------------------------------------
            return table;

        }
    }

    class Setting_DataTable_Variables
    {
        #region
        private bool ColVisible;
        public bool GS_ColVisible
        {
            get { return ColVisible; }
            set { ColVisible = value; }
        }

        private string Header_Name;
        public string GS_Header_Name
        {
            get { return Header_Name; }
            set { Header_Name = value; }
        }

        private int Column_Width;
        public int GS_Column_Width
        {
            get { return Column_Width; }
            set { Column_Width = value; }
        }

        private Color BC;
        public Color GS_BC
        {
            get { return BC; }
            set { BC = value; }
        }

        private Color FC;
        public Color GS_FC
        {
            get { return FC; }
            set { FC = value; }
        }

        private int DI;
        public int GS_DI
        {
            get { return DI; }
            set { DI = value; }
        }
        //-----------------------------------------------
        private string Fnt_name;
        public string GS_Fnt_name
        {
            get { return Fnt_name; }
            set { Fnt_name = value; }
        }

        private string Fnt_style;
        public string GS_Fnt_style
        {
            get { return Fnt_style; }
            set { Fnt_style = value; }
        }

        private bool Fnt_underline;
        public bool GS_Fnt_underline
        {
            get { return Fnt_underline; }
            set { Fnt_underline = value; }
        }

        private bool Fnt_strikeout;
        public bool GS_Fnt_strikeout
        {
            get { return Fnt_strikeout; }
            set { Fnt_strikeout = value; }
        }

        private float Fnt_size;
        public float GS_Fnt_size
        {
            get { return Fnt_size; }
            set { Fnt_size = value; }
        }
        //-----------------------------------------------
        #endregion
    }

    class LOG
    {
        private FileInfo Log_file_info = new FileInfo(
                    "\\\\10.214.10.26\\Departments\\Public\\Temporary\\Jan Jelinek\\LOGs\\Back_Order_Report" +
                    "   Computer=" + Environment.MachineName.ToString() +
                    "   User=" + Environment.UserName.ToString() +
                    ".kix");

        public void Create_LOG_file(Exception ex, object Oracle, object MySql)
        {
            //-------------------------------------------------------------------------------------------------------------------------------------------------
            MySqlConnection MySql_connection = null;
            #region
            string MySql_command_text = null;

            if (MySql != null)
            {
                if (MySql is MySqlCommand)
                {
                    MySql_connection = ((MySqlCommand)MySql).Connection;
                    MySql_command_text = ((MySqlCommand)MySql).CommandText;
                }
                else
                {
                    MySql_connection = ((MySqlDataAdapter)MySql).SelectCommand.Connection;
                    MySql_command_text = ((MySqlDataAdapter)MySql).SelectCommand.CommandText;
                }
            }
            #endregion
            //-------------------------------------------------------------------------------------------------------------------------------------------------
            OracleConnection ORA_connection = null;
            #region
            string ORA_command_text = null;

            if (Oracle != null)
            {
                if (MySql is OracleCommand)
                {
                    ORA_connection = ((OracleCommand)Oracle).Connection;
                    ORA_command_text = ((OracleCommand)Oracle).CommandText;
                }
                else
                {
                    ORA_connection = ((OracleDataAdapter)Oracle).SelectCommand.Connection;
                    ORA_command_text = ((OracleDataAdapter)Oracle).SelectCommand.CommandText;
                }
            }
            #endregion
            //-------------------------------------------------------------------------------------------------------------------------------------------------
            if ((ex.Message.ToString().IndexOf("'PRIMARY'")) >= 0) return;


            StackTrace st = new StackTrace(ex, true);

            string CommandText = string.Empty;

            try
            {
                CommandText = new Regex("\r\n +").Replace(
                    (MySql == null ? ORA_command_text : MySql_command_text), Environment.NewLine);
            }
            catch { }

            int exist_char = CommandText.IndexOf("INTO");
            if (exist_char > 0)
            {
                CommandText = CommandText.Replace("VALUES", "VALUES" + Environment.NewLine).
                                          Replace("(", "(" + Environment.NewLine).
                                          Replace(", ", ", " + Environment.NewLine).
                                          Replace("'<", "'" + Environment.NewLine + "<").
                                          Replace(")", Environment.NewLine + ")");
            }

            string ConnectionString = string.Empty;

            try
            {
                ConnectionString = (MySql == null ?
                    ORA_connection.ConnectionString.Replace(";", Environment.NewLine).Replace(" ", "") +
                    "HostName=" + ORA_connection.HostName.ToUpper() + Environment.NewLine +
                    "State=" + ORA_connection.State :
                    MySql_connection.ConnectionString.Replace(";", Environment.NewLine).Replace(" ", "") +
                    Environment.NewLine +
                    "State=" + MySql_connection.State);
            }
            catch { }

            string ExceptionErrorsStackTrace = ex.ToString().
                Replace(",", "," + Environment.NewLine + "   ").
                Replace("   at", "at").
                Replace(" in", Environment.NewLine + "    in").
                Replace("):", "):" + Environment.NewLine + "   ");

            string Log_Information = DateTime.Now.ToString("dddd dd-MM-yyyy in HH:mm") +
                Environment.NewLine + "#############################" +
                Environment.NewLine +
                Environment.NewLine + (MySql == null ? "OracleCommand.CommandText:" : "MySqlCommand.CommandText:") +
                Environment.NewLine + "--------------------------" +
                Environment.NewLine + CommandText +
                Environment.NewLine +
                Environment.NewLine + (MySql == null ? "OracleCommand.Connection.ConnectionString:" : "MySqlCommand.Connection.ConnectionString:") +
                Environment.NewLine + "------------------------------------------" +
                Environment.NewLine + ConnectionString.Replace("userid", "UserId") +
                Environment.NewLine +
                Environment.NewLine + "StackTrace.RuntimeMethodInfo.Name:" +
                Environment.NewLine + "----------------------------------" +
                Environment.NewLine + st.GetFrame(st.FrameCount - 1).GetMethod().ToString().Replace("(", "(" + Environment.NewLine).
                                                                                            Replace(", ", ", " + Environment.NewLine).
                                                                                            Replace(")", Environment.NewLine + ")").
                                                                                            Replace("(" + Environment.NewLine + Environment.NewLine + ")", "()") +
                                                                                            Environment.NewLine + "in " + st.GetFrame(st.FrameCount - 1).GetMethod().ReflectedType.Name +
                Environment.NewLine +
                Environment.NewLine + "Exception.Errors.Message:" +
                Environment.NewLine + "-------------------------" +
                Environment.NewLine + ex.Message +
                Environment.NewLine +
                Environment.NewLine + "Exception.Errors.StackTrace.LineNumber:" +
                Environment.NewLine + "---------------------------------------" +
                Environment.NewLine + st.GetFrame(st.FrameCount - 1).GetFileLineNumber() +
                Environment.NewLine +
                Environment.NewLine + "Exception.Errors.StackTrace:" +
                Environment.NewLine + "----------------------------" +
                Environment.NewLine + ExceptionErrorsStackTrace +
                Environment.NewLine + "####################################################################################################################################################################" +
                Environment.NewLine + "####################################################################################################################################################################" +
                Environment.NewLine + "####################################################################################################################################################################" +
                Environment.NewLine + Environment.NewLine;


            File.AppendAllText(Log_file_info.FullName, Log_Information);
        }
    }
}