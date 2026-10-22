// ac2f pack - a stand-in for System.Windows.Forms and the GDI+ half of
// System.Drawing, so the window scripts run under PowerShell on Linux in
// the tests (tools/test_windows.py). Only what the scripts use exists; a
// member a script uses that is missing here fails the test, which is the
// point. Behaviour follows WinForms where a mistake would matter:
// SelectedIndex out of range throws, a DataGridView row can only be
// inserted when detached, AddPolygon wants 3 points, a Bitmap wants a
// positive size, Close() raises FormClosing, and so on.
//
// Color, Size, Point, PointF ... are the real ones (System.Drawing.Primitives
// works everywhere). The GDI+ types live in StubDrawing; the test rewrites
// System.Drawing.Font/Bitmap/Graphics/... in a script to StubDrawing.*.
#pragma warning disable 67, 169, 414, 649
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;

namespace StubDrawing
{
    public enum FontStyle { Regular = 0, Bold = 1, Italic = 2, Underline = 4, Strikeout = 8 }

    public static class Stats
    {
        public static int Bitmaps, Renders, Polygons, Points, Lines, Strings, Rects;
        public static List<string> Text = new List<string>();
        public static void Reset() { Bitmaps = Renders = Polygons = Points = Lines = Strings = Rects = 0; Text.Clear(); }
    }

    public class Font : IDisposable
    {
        public string Name; public float Size; public FontStyle Style;
        public Font(string name, float size) : this(name, size, FontStyle.Regular) { }
        public Font(string name, float size, FontStyle style)
        {
            if (size <= 0 || float.IsNaN(size)) throw new ArgumentException("Value of 'size' is not valid.");
            Name = name; Size = size; Style = style;
        }
        public bool Bold { get { return (Style & FontStyle.Bold) != 0; } }
        public void Dispose() { }
    }

    public static class SystemColors
    {
        public static Color Window { get { return Color.White; } }
        public static Color WindowText { get { return Color.Black; } }
        public static Color Control { get { return Color.LightGray; } }
        public static Color ControlText { get { return Color.Black; } }
        public static Color GrayText { get { return Color.Gray; } }
    }

    public abstract class Image : IDisposable
    {
        public int Width, Height;
        public bool Disposed;
        public void Dispose() { Disposed = true; }
    }

    public class Bitmap : Image
    {
        public Bitmap(int width, int height)
        {
            if (width <= 0 || height <= 0) throw new ArgumentException("Parameter is not valid.");
            Width = width; Height = height; Stats.Bitmaps++;
        }
    }

    public abstract class Brush : IDisposable { public void Dispose() { } }

    public class SolidBrush : Brush
    {
        public Color Color;
        public SolidBrush(Color color) { Color = color; }
    }

    public static class Brushes
    {
        public static Brush Black { get { return new SolidBrush(Color.Black); } }
        public static Brush White { get { return new SolidBrush(Color.White); } }
    }

    public class Pen : IDisposable
    {
        public Color Color;
        float width = 1;
        public Drawing2D.DashStyle DashStyle;
        public Pen(Color color) { Color = color; }
        public Pen(Color color, float width) { Color = color; Width = width; }
        public float Width
        {
            get { return width; }
            set
            {
                if (value < 0 || float.IsNaN(value) || float.IsInfinity(value)) throw new ArgumentException("Parameter is not valid.");
                width = value;
            }
        }
        public void Dispose() { }
    }

    public class Graphics : IDisposable
    {
        Image img;
        bool disposed;
        public Drawing2D.SmoothingMode SmoothingMode { get; set; }
        Graphics(Image i) { img = i; }

        public static Graphics FromImage(Image image)
        {
            if (image == null) throw new ArgumentNullException("image");
            if (image.Disposed) throw new ArgumentException("Parameter is not valid.");
            Stats.Renders++;
            return new Graphics(image);
        }
        void Live() { if (disposed) throw new ArgumentException("Parameter is not valid."); }
        static void Finite(params float[] v)
        {
            foreach (var x in v)
                if (float.IsNaN(x) || float.IsInfinity(x)) throw new OverflowException("Overflow error.");
        }
        public void Clear(Color c) { Live(); }
        public void TranslateTransform(float dx, float dy) { Live(); Finite(dx, dy); }
        public void ScaleTransform(float sx, float sy)
        {
            Live(); Finite(sx, sy);
            if (sx == 0 || sy == 0) throw new ArgumentException("Parameter is not valid.");
        }
        public void ResetTransform() { Live(); }
        public void FillRectangle(Brush b, float x, float y, float w, float h)
        {
            Live(); if (b == null) throw new ArgumentNullException("brush"); Finite(x, y, w, h); Stats.Rects++;
        }
        public void DrawRectangle(Pen p, float x, float y, float w, float h)
        {
            Live(); if (p == null) throw new ArgumentNullException("pen"); Finite(x, y, w, h); Stats.Rects++;
        }
        public void FillPath(Brush b, Drawing2D.GraphicsPath gp)
        {
            Live(); if (b == null || gp == null) throw new ArgumentNullException();
            Stats.Polygons += gp.PolyCount; Stats.Points += gp.PointCount;
        }
        public void DrawPath(Pen p, Drawing2D.GraphicsPath gp)
        {
            Live(); if (p == null || gp == null) throw new ArgumentNullException();
        }
        public void DrawLines(Pen p, PointF[] pts)
        {
            Live(); if (p == null || pts == null) throw new ArgumentNullException();
            if (pts.Length < 2) throw new ArgumentException("Parameter is not valid.");
            Stats.Lines++;
        }
        public void DrawString(string s, Font f, Brush b, float x, float y)
        {
            Live(); if (f == null || b == null) throw new ArgumentNullException(); Finite(x, y);
            Stats.Strings++; Stats.Text.Add(s);
        }
        public SizeF MeasureString(string s, Font f) { return new SizeF(7 * (s ?? "").Length, 14); }
        public void Dispose() { disposed = true; }
    }

    namespace Drawing2D
    {
        public enum SmoothingMode { Invalid = -1, Default = 0, HighSpeed = 1, HighQuality = 2, None = 3, AntiAlias = 4 }
        public enum FillMode { Alternate = 0, Winding = 1 }
        public enum DashStyle { Solid, Dash, Dot, DashDot, DashDotDot, Custom }

        public class GraphicsPath : IDisposable
        {
            public FillMode FillMode { get; set; }
            public int PolyCount, PointCount;
            public void AddPolygon(PointF[] points)
            {
                if (points == null) throw new ArgumentNullException("points");
                if (points.Length < 3) throw new ArgumentException("A polygon needs at least 3 points.");
                PolyCount++; PointCount += points.Length;
            }
            public void AddLines(PointF[] points)
            {
                if (points == null || points.Length < 2) throw new ArgumentException("Parameter is not valid.");
                PointCount += points.Length;
            }
            public void Dispose() { }
        }
    }
}

namespace System.Windows.Forms
{
    using StubDrawing;

    public enum DockStyle { None, Top, Bottom, Left, Right, Fill }
    [Flags] public enum AnchorStyles { None = 0, Top = 1, Bottom = 2, Left = 4, Right = 8 }
    public enum FormStartPosition { Manual, CenterScreen, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent }
    public enum SizeType { AutoSize, Absolute, Percent }
    public enum ComboBoxStyle { Simple, DropDown, DropDownList }
    public enum FlowDirection { LeftToRight, TopDown, RightToLeft, BottomUp }
    public enum DialogResult { None, OK, Cancel, Abort, Retry, Ignore, Yes, No }
    public enum DataGridViewSelectionMode { CellSelect, FullRowSelect, FullColumnSelect, RowHeaderSelect, ColumnHeaderSelect }
    public enum DataGridViewAutoSizeColumnsMode { None = 1, ColumnHeader = 2, AllCellsExceptHeader = 4, AllCells = 6, DisplayedCellsExceptHeader = 8, DisplayedCells = 10, Fill = 16 }
    public enum ScrollBars { None, Horizontal, Vertical, Both }
    public enum BorderStyle { None, FixedSingle, Fixed3D }
    public enum PictureBoxSizeMode { Normal, StretchImage, AutoSize, CenterImage, Zoom }
    public enum HorizontalAlignment { Left, Right, Center }
    public enum SelectionMode { None, One, MultiSimple, MultiExtended }
    public enum Keys { None = 0, Back = 8, Tab = 9, Enter = 13, Escape = 27, Space = 32, Up = 38, Down = 40, F5 = 116 }

    public struct Padding
    {
        public int Left, Top, Right, Bottom;
        public Padding(int all) { Left = Top = Right = Bottom = all; }
        public Padding(int left, int top, int right, int bottom) { Left = left; Top = top; Right = right; Bottom = bottom; }
    }

    // the test's hand on the window: called once the dialog is "on screen"
    public static class Harness
    {
        public static Action<Form> OnShow;
        public static List<string> Notes = new List<string>();
        public static List<string> Messages = new List<string>();
        public static int ClientW = 900, ClientH = 640;
    }

    public static class Application
    {
        public static void EnableVisualStyles() { }
        public static void DoEvents() { }
    }

    public class Control : IDisposable
    {
        string text = "";
        int width = 100, height = 30;
        public Control() { Controls = new ControlCollection(this); Visible = true; Enabled = true; }
        public string Name { get; set; }
        public virtual string Text { get { return text; } set { text = value ?? ""; if (TextChanged != null) TextChanged(this, EventArgs.Empty); } }
        public DockStyle Dock { get; set; }
        public AnchorStyles Anchor { get; set; }
        public Padding Margin { get; set; }
        public Padding Padding { get; set; }
        public bool AutoSize { get; set; }
        public bool AutoScroll { get; set; }
        public StubDrawing.Font Font { get; set; }
        public Color ForeColor { get; set; }
        public Color BackColor { get; set; }
        public int Width { get { return width; } set { width = value; FireResize(); } }
        public int Height { get { return height; } set { height = value; FireResize(); } }
        public Size Size
        {
            get { return new Size(width, height); }
            set { width = value.Width; height = value.Height; FireResize(); }
        }
        public Size ClientSize { get { return Size; } set { Size = value; } }
        public Size MinimumSize { get; set; }
        public Size MaximumSize { get; set; }
        public Point Location { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public object Tag { get; set; }
        public int TabIndex { get; set; }
        public Control Parent { get; internal set; }
        public ControlCollection Controls { get; private set; }
        public bool Focused { get; private set; }
        public bool Focus() { Focused = true; return true; }
        public void Select() { Focused = true; }
        public void Refresh() { }
        public void Invalidate() { }
        public void SuspendLayout() { }
        public void ResumeLayout() { }
        public void ResumeLayout(bool performLayout) { }
        public event EventHandler Click;
        public event EventHandler DoubleClick;
        public event EventHandler Resize;
        public event EventHandler TextChanged;
        public event EventHandler Leave;
        public event KeyEventHandler KeyDown;
        public void PerformClick() { OnClick(); }
        protected virtual void OnClick() { if (Click != null) Click(this, EventArgs.Empty); }
        public void FireDoubleClick() { if (DoubleClick != null) DoubleClick(this, EventArgs.Empty); }
        public void FireKey(KeyEventArgs e) { if (KeyDown != null) KeyDown(this, e); }
        void FireResize() { if (Resize != null) Resize(this, EventArgs.Empty); }
        public void Dispose() { }

        // test helpers
        public IEnumerable<Control> All()
        {
            yield return this;
            foreach (Control c in Controls)
                foreach (var d in c.All()) yield return d;
        }
        public Control Find(string text)
        {
            foreach (var c in All()) if (c != this && c.Text == text) return c;
            return null;
        }
        public T First<T>() where T : Control
        {
            foreach (var c in All()) if (c is T) return (T)c;
            return null;
        }
        public List<T> Every<T>() where T : Control
        {
            var l = new List<T>();
            foreach (var c in All()) if (c is T) l.Add((T)c);
            return l;
        }
        // the same for PowerShell, by type name
        public Control FirstOf(string typeName)
        {
            foreach (var c in All()) if (c.GetType().Name == typeName) return c;
            return null;
        }
        public List<Control> AllOf(string typeName)
        {
            var l = new List<Control>();
            foreach (var c in All()) if (c.GetType().Name == typeName) l.Add(c);
            return l;
        }
    }

    public class ControlCollection : IEnumerable
    {
        readonly List<Control> items = new List<Control>();
        readonly Control owner;
        public ControlCollection(Control owner) { this.owner = owner; }
        public void Add(Control c)
        {
            if (c == null) throw new ArgumentNullException("value");
            if (owner is TabControl && !(c is TabPage)) throw new ArgumentException("TabControl takes TabPages only.");
            c.Parent = owner; items.Add(c);
        }
        public void Add(Control c, int column, int row)
        {
            var t = owner as TableLayoutPanel;
            if (t == null) throw new InvalidOperationException("Add(control, column, row) is for a TableLayoutPanel.");
            if (column < -1 || row < -1) throw new ArgumentOutOfRangeException("column");
            if (t.ColumnCount > 0 && column >= t.ColumnCount)
                throw new ArgumentException("Column " + column + " is outside the " + t.ColumnCount + " columns.");
            Add(c);
        }
        public void Remove(Control c) { items.Remove(c); }
        public bool Contains(Control c) { return items.Contains(c); }
        public int Count { get { return items.Count; } }
        public Control this[int i] { get { return items[i]; } }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public class Label : Control { }
    public class ButtonBase : Control { }
    public class Button : ButtonBase { public DialogResult DialogResult { get; set; } }

    public class CheckBox : ButtonBase
    {
        bool isChecked;
        public bool Checked
        {
            get { return isChecked; }
            set { isChecked = value; if (CheckedChanged != null) CheckedChanged(this, EventArgs.Empty); }
        }
        public event EventHandler CheckedChanged;
        protected override void OnClick() { Checked = !Checked; base.OnClick(); }
    }

    public class TextBox : Control
    {
        public bool Multiline { get; set; }
        public bool ReadOnly { get; set; }
        public bool WordWrap { get; set; }
        public ScrollBars ScrollBars { get; set; }
        public HorizontalAlignment TextAlign { get; set; }
        public int MaxLength { get; set; }
        public void SelectAll() { }
    }

    public class ObjectCollection : IEnumerable
    {
        readonly List<object> items = new List<object>();
        readonly ListControlBase owner;
        public ObjectCollection(ListControlBase owner) { this.owner = owner; }
        public int Add(object item)
        {
            if (item == null) throw new ArgumentNullException("item");
            items.Add(item); return items.Count - 1;
        }
        public void Remove(object item)
        {
            int i = items.IndexOf(item);
            if (i < 0) return;
            items.RemoveAt(i);
            owner.ItemRemoved(i);
        }
        public void Clear() { items.Clear(); owner.ItemRemoved(-1); }
        public int Count { get { return items.Count; } }
        public bool Contains(object o) { return items.Contains(o); }
        public int IndexOf(object o) { return items.IndexOf(o); }
        public object this[int i] { get { return items[i]; } }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public abstract class ListControlBase : Control
    {
        int sel = -1;
        protected ListControlBase() { Items = new ObjectCollection(this); }
        public ObjectCollection Items { get; private set; }
        public event EventHandler SelectedIndexChanged;
        public virtual int SelectedIndex
        {
            get { return sel; }
            set
            {
                if (value < -1 || value >= Items.Count)
                    throw new ArgumentOutOfRangeException("SelectedIndex", "InvalidArgument=Value of '" + value + "' is not valid for 'SelectedIndex'.");
                sel = value;
                if (SelectedIndexChanged != null) SelectedIndexChanged(this, EventArgs.Empty);
            }
        }
        public object SelectedItem
        {
            get { return sel >= 0 ? Items[sel] : null; }
            set { SelectedIndex = Items.IndexOf(value); }
        }
        internal void ItemRemoved(int i)
        {
            if (i < 0 || i == sel) sel = -1;
            else if (i < sel) sel--;
        }
    }

    public class ComboBox : ListControlBase
    {
        string typed = "";
        public ComboBoxStyle DropDownStyle { get; set; }
        public override int SelectedIndex
        {
            get { return base.SelectedIndex; }
            set { base.SelectedIndex = value; if (value >= 0) typed = Items[value].ToString(); }
        }
        public override string Text
        {
            get
            {
                if (DropDownStyle == ComboBoxStyle.DropDownList) return SelectedIndex >= 0 ? Items[SelectedIndex].ToString() : "";
                return typed;
            }
            set
            {
                if (DropDownStyle == ComboBoxStyle.DropDownList)
                {
                    int i = Items.IndexOf(value);
                    if (i >= 0) SelectedIndex = i;
                }
                else typed = value ?? "";
            }
        }
    }

    public class ListBox : ListControlBase
    {
        public bool IntegralHeight { get; set; }
        public SelectionMode SelectionMode { get; set; }
    }

    public class GroupBox : Control { }
    public class Panel : Control { public BorderStyle BorderStyle { get; set; } }

    public class FlowLayoutPanel : Panel
    {
        public bool WrapContents { get; set; }
        public FlowDirection FlowDirection { get; set; }
    }

    public class TableLayoutStyle { public SizeType SizeType { get; set; } }
    public class ColumnStyle : TableLayoutStyle
    {
        public float Width;
        public ColumnStyle() { }
        public ColumnStyle(SizeType t) { SizeType = t; }
        public ColumnStyle(SizeType t, float width) { SizeType = t; Width = width; }
    }
    public class RowStyle : TableLayoutStyle
    {
        public float Height;
        public RowStyle() { }
        public RowStyle(SizeType t) { SizeType = t; }
        public RowStyle(SizeType t, float height) { SizeType = t; Height = height; }
    }
    public class TableLayoutStyleCollection : IEnumerable
    {
        readonly List<TableLayoutStyle> items = new List<TableLayoutStyle>();
        public int Add(TableLayoutStyle s) { if (s == null) throw new ArgumentNullException(); items.Add(s); return items.Count - 1; }
        public int Count { get { return items.Count; } }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public class TableLayoutPanel : Panel
    {
        public int ColumnCount { get; set; }
        public int RowCount { get; set; }
        public TableLayoutStyleCollection ColumnStyles { get; private set; }
        public TableLayoutStyleCollection RowStyles { get; private set; }
        public TableLayoutPanel() { ColumnStyles = new TableLayoutStyleCollection(); RowStyles = new TableLayoutStyleCollection(); }
        public void SetColumnSpan(Control c, int value)
        {
            if (c == null) throw new ArgumentNullException("control");
            if (value < 1) throw new ArgumentOutOfRangeException("value");
        }
        public void SetRowSpan(Control c, int value) { SetColumnSpan(c, value); }
    }

    public class TabPage : Panel { }
    public class TabPageCollection : IEnumerable
    {
        readonly TabControl owner;
        readonly List<TabPage> items = new List<TabPage>();
        public TabPageCollection(TabControl owner) { this.owner = owner; }
        public void Add(TabPage p) { if (p == null) throw new ArgumentNullException(); items.Add(p); owner.Controls.Add(p); }
        public int Count { get { return items.Count; } }
        public TabPage this[int i] { get { return items[i]; } }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }
    public class TabControl : Control
    {
        int sel;
        public TabControl() { TabPages = new TabPageCollection(this); }
        public TabPageCollection TabPages { get; private set; }
        public int SelectedIndex
        {
            get { return sel; }
            set { if (value < -1 || value >= TabPages.Count) throw new ArgumentOutOfRangeException("SelectedIndex"); sel = value; }
        }
    }

    public class PictureBox : Control
    {
        public StubDrawing.Image Image { get; set; }
        public PictureBoxSizeMode SizeMode { get; set; }
        public BorderStyle BorderStyle { get; set; }
    }

    public class ToolTip : IDisposable
    {
        public int AutoPopDelay { get; set; }
        public int InitialDelay { get; set; }
        public int ReshowDelay { get; set; }
        public Dictionary<Control, string> Tips = new Dictionary<Control, string>();
        public void SetToolTip(Control c, string caption)
        {
            if (c == null) throw new ArgumentNullException("control");
            Tips[c] = caption;
        }
        public void Dispose() { }
    }

    public class FormClosingEventArgs : EventArgs { public bool Cancel { get; set; } }
    public delegate void FormClosingEventHandler(object sender, FormClosingEventArgs e);
    public class KeyEventArgs : EventArgs
    {
        public KeyEventArgs(Keys k) { KeyCode = k; }
        public Keys KeyCode { get; private set; }
        public bool Handled { get; set; }
        public bool SuppressKeyPress { get; set; }
        public bool Control { get; set; }
        public bool Shift { get; set; }
    }
    public delegate void KeyEventHandler(object sender, KeyEventArgs e);

    public class Form : Control
    {
        public FormStartPosition StartPosition { get; set; }
        public bool TopMost { get; set; }
        public bool KeyPreview { get; set; }
        public bool MaximizeBox { get; set; }
        public bool MinimizeBox { get; set; }
        public bool ShowInTaskbar { get; set; }
        public Button AcceptButton { get; set; }
        public Button CancelButton { get; set; }
        public DialogResult DialogResult { get; set; }
        public bool IsClosed { get; private set; }
        public event EventHandler Load;
        public event EventHandler Shown;
        public event FormClosingEventHandler FormClosing;
        public event EventHandler FormClosed;

        public DialogResult ShowDialog() { return ShowDialog(null); }
        public DialogResult ShowDialog(object owner)
        {
            if (Load != null) Load(this, EventArgs.Empty);
            // lay out: whatever fills its cell gets a real size
            foreach (var c in All())
                if (c != this && c.Dock == DockStyle.Fill) c.Size = new Size(Harness.ClientW, Harness.ClientH);
            if (Shown != null) Shown(this, EventArgs.Empty);
            if (Harness.OnShow != null) Harness.OnShow(this);
            if (!IsClosed)
            {
                Harness.Notes.Add("the test left the dialog open; closed as by the X button");
                Close();
            }
            return DialogResult;
        }
        public void Close()
        {
            if (IsClosed) return;
            var e = new FormClosingEventArgs();
            if (FormClosing != null) FormClosing(this, e);
            if (e.Cancel) return;
            IsClosed = true;
            if (FormClosed != null) FormClosed(this, EventArgs.Empty);
        }
        public void Show() { }
        public void Activate() { }
        public void BringToFront() { }
    }

    public static class MessageBox
    {
        public static DialogResult Show(string text) { Harness.Messages.Add(text); return DialogResult.OK; }
        public static DialogResult Show(string text, string caption) { return Show(text); }
        public static DialogResult Show(Form owner, string text, string caption) { return Show(text); }
    }

    // ---------------- DataGridView ----------------
    public class DataGridViewCellStyle { public Color ForeColor { get; set; } public Color BackColor { get; set; } }

    public class DataGridViewColumn
    {
        public string HeaderText { get; set; }
        public float FillWeight { get; set; }
        public string ToolTipText { get; set; }
        public bool ReadOnly { get; set; }
        public string Name { get; set; }
        public int Width { get; set; }
    }
    public class DataGridViewTextBoxColumn : DataGridViewColumn { }
    public class DataGridViewCheckBoxColumn : DataGridViewColumn { }

    public class DataGridViewColumnCollection : IEnumerable
    {
        readonly List<DataGridViewColumn> items = new List<DataGridViewColumn>();
        public int Add(DataGridViewColumn c) { if (c == null) throw new ArgumentNullException(); items.Add(c); return items.Count - 1; }
        public void AddRange(params DataGridViewColumn[] cols) { foreach (var c in cols) Add(c); }
        public int Count { get { return items.Count; } }
        public DataGridViewColumn this[int i] { get { return items[i]; } }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public class DataGridViewCell
    {
        public object Value { get; set; }
        public bool ReadOnly { get; set; }
        public DataGridViewRow OwningRow { get; internal set; }
        public int ColumnIndex { get; internal set; }
    }

    public class DataGridViewCellCollection : IEnumerable
    {
        internal readonly List<DataGridViewCell> items = new List<DataGridViewCell>();
        public int Count { get { return items.Count; } }
        public DataGridViewCell this[int i]
        {
            get
            {
                if (i < 0 || i >= items.Count) throw new ArgumentOutOfRangeException("index", "Index was out of range.");
                return items[i];
            }
        }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public class DataGridViewRow
    {
        public DataGridViewRow() { Cells = new DataGridViewCellCollection(); DefaultCellStyle = new DataGridViewCellStyle(); }
        public DataGridView DataGridView { get; internal set; }
        public object Tag { get; set; }
        public DataGridViewCellCollection Cells { get; private set; }
        public DataGridViewCellStyle DefaultCellStyle { get; set; }
        public int Index { get { return DataGridView == null ? -1 : DataGridView.Rows.IndexOf(this); } }
    }

    public class DataGridViewRowCollection : IEnumerable
    {
        readonly DataGridView grid;
        readonly List<DataGridViewRow> items = new List<DataGridViewRow>();
        public DataGridViewRowCollection(DataGridView g) { grid = g; }
        public int Add() { return Add(new object[0]); }
        public int Add(int count) { int r = -1; for (int i = 0; i < count; i++) r = Add(); return r; }
        public int Add(DataGridViewRow row) { Attach(row); items.Add(row); return items.Count - 1; }
        public int Add(params object[] values)
        {
            if (grid.Columns.Count == 0) throw new InvalidOperationException("No row can be added to a DataGridView control that does not have columns.");
            if (values.Length > grid.Columns.Count) throw new ArgumentException("More values than columns.");
            var r = new DataGridViewRow();
            for (int i = 0; i < grid.Columns.Count; i++)
                r.Cells.items.Add(new DataGridViewCell { Value = i < values.Length ? values[i] : null, OwningRow = r, ColumnIndex = i });
            r.DataGridView = grid;
            items.Add(r);
            return items.Count - 1;
        }
        void Attach(DataGridViewRow row)
        {
            if (row == null) throw new ArgumentNullException("dataGridViewRow");
            if (row.DataGridView != null) throw new InvalidOperationException("Row provided already belongs to a DataGridView control.");
            row.DataGridView = grid;
        }
        public void Insert(int index, DataGridViewRow row)
        {
            if (index < 0 || index > items.Count) throw new ArgumentOutOfRangeException("rowIndex");
            Attach(row);
            items.Insert(index, row);
        }
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= items.Count) throw new ArgumentOutOfRangeException("index");
            var r = items[index];
            if (grid.CurrentCell != null && grid.CurrentCell.OwningRow == r) grid.ClearCurrent();
            items.RemoveAt(index);
            r.DataGridView = null;
        }
        public int IndexOf(DataGridViewRow r) { return items.IndexOf(r); }
        public int Count { get { return items.Count; } }
        public DataGridViewRow this[int i]
        {
            get
            {
                if (i < 0 || i >= items.Count) throw new ArgumentOutOfRangeException("index", "Index was out of range.");
                return items[i];
            }
        }
        public IEnumerator GetEnumerator() { return items.GetEnumerator(); }
    }

    public class DataGridView : Control
    {
        DataGridViewCell current;
        public DataGridView() { Columns = new DataGridViewColumnCollection(); Rows = new DataGridViewRowCollection(this); }
        public bool AllowUserToAddRows { get; set; }
        public bool AllowUserToDeleteRows { get; set; }
        public bool AllowUserToResizeRows { get; set; }
        public bool RowHeadersVisible { get; set; }
        public bool MultiSelect { get; set; }
        public bool ReadOnly { get; set; }
        public DataGridViewSelectionMode SelectionMode { get; set; }
        public DataGridViewAutoSizeColumnsMode AutoSizeColumnsMode { get; set; }
        public Color BackgroundColor { get; set; }
        public DataGridViewColumnCollection Columns { get; private set; }
        public DataGridViewRowCollection Rows { get; private set; }
        public DataGridViewCell CurrentCell
        {
            get { return current; }
            set
            {
                if (value != null && (value.OwningRow == null || value.OwningRow.DataGridView != this))
                    throw new ArgumentException("The cell does not belong to this DataGridView.");
                current = value;
            }
        }
        internal void ClearCurrent() { current = null; }
        public DataGridViewRow CurrentRow { get { return current == null ? null : current.OwningRow; } }
        public bool EndEdit() { return true; }
    }
}
