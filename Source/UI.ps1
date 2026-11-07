# Windows Forms presentation. Worker runspaces never access controls.
function Initialize-AppPresentation {
    if ('QBitLauncher.Presentation' -as [type]) { return }
    Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace QBitLauncher {
    public static class Presentation {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        [DllImport("uxtheme.dll", CharSet=CharSet.Unicode)] public static extern int SetWindowTheme(IntPtr window, string application, string list);
        [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr window, IntPtr device, uint flags);
        public static void Capture(Form form, string path) {
            using (Bitmap bitmap = new Bitmap(form.Width,form.Height)) {
                using (Graphics graphics = Graphics.FromImage(bitmap)) {
                    IntPtr device = graphics.GetHdc();
                    try { if (!PrintWindow(form.Handle,device,0)) throw new InvalidOperationException("The window could not be captured."); }
                    finally { graphics.ReleaseHdc(device); }
                }
                bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
            }
        }
        public static void TitleBar(Form form, bool dark, Color caption, Color text) {
            try {
                int enabled = dark ? 1 : 0;
                DwmSetWindowAttribute(form.Handle, 20, ref enabled, 4);
                int background = caption.R | (caption.G << 8) | (caption.B << 16);
                int foreground = text.R | (text.G << 8) | (text.B << 16);
                DwmSetWindowAttribute(form.Handle, 35, ref background, 4);
                DwmSetWindowAttribute(form.Handle, 36, ref foreground, 4);
            } catch (DllNotFoundException) { } catch (EntryPointNotFoundException) { }
        }
    }
    public class ThemedListView : ListView {
        public Color HeaderColor { get; set; }
        public Color MutedColor { get; set; }
        public Color SelectionColor { get; set; }
        public Color AccentColor { get; set; }
        public Color DividerColor { get; set; }
        public bool Dark { get; set; }
        bool fitting;
        int[] preferred;
        HeaderWindow header;
        ListViewItem editingItem;
        Panel editorSurface;
        TextBox nameEditor;
        public bool IsShortcutNameEditing { get { return editingItem != null; } }
        [StructLayout(LayoutKind.Sequential)] struct NativeRect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window, int message, IntPtr first, IntPtr second);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window, int message, IntPtr first, ref NativeRect rectangle);
        [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr window, out NativeRect rectangle);
        [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr window, IntPtr device);
        [DllImport("user32.dll")] static extern int GetScrollPos(IntPtr window, int bar);
        [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr window, IntPtr rectangle, bool erase);
        sealed class HeaderWindow : NativeWindow {
            readonly ThemedListView owner;
            public HeaderWindow(ThemedListView list, IntPtr handle) { owner = list; AssignHandle(handle); }
            protected override void WndProc(ref Message message) {
                base.WndProc(ref message);
                if (message.Msg != 0x000F && message.Msg != 0x0318) return;
                IntPtr device = message.Msg == 0x0318 ? message.WParam : GetDC(Handle);
                try { if (device != IntPtr.Zero) owner.PaintHeader(Handle, device); }
                finally { if (message.Msg == 0x000F && device != IntPtr.Zero) ReleaseDC(Handle, device); }
            }
        }
        public ThemedListView() {
            OwnerDraw = true; DoubleBuffered = true; BorderStyle = BorderStyle.None;
            HeaderColor = Color.FromArgb(40,42,54); MutedColor = Color.FromArgb(150,150,150);
            SelectionColor = Color.FromArgb(68,71,90); AccentColor = Color.FromArgb(98,114,164);
            DividerColor = Color.FromArgb(68,71,90); Dark = true;
        }
        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            Presentation.SetWindowTheme(Handle, Dark ? "DarkMode_Explorer" : "Explorer", null);
            IntPtr handle = SendMessage(Handle, 0x101F, IntPtr.Zero, IntPtr.Zero);
            if (handle != IntPtr.Zero) header = new HeaderWindow(this, handle);
        }
        protected override void OnHandleDestroyed(EventArgs e) {
            FinishShortcutNameEdit(true);
            if (header != null) { header.ReleaseHandle(); header = null; }
            base.OnHandleDestroyed(e);
        }
        void PaintHeader(IntPtr handle, IntPtr device) {
            NativeRect client;
            if (!GetClientRect(handle, out client)) return;
            using (Graphics graphics = Graphics.FromHdc(device))
            using (Brush surface = new SolidBrush(HeaderColor))
            using (Pen divider = new Pen(DividerColor)) {
                // Paint the entire native header, including its empty trailing area.
                graphics.FillRectangle(surface, 0, 0, client.Right, client.Bottom);
                for (int index = 0; index < Columns.Count; index++) {
                    NativeRect native = new NativeRect();
                    SendMessage(handle, 0x1207, new IntPtr(index), ref native);
                    Rectangle text = new Rectangle(native.Left+10,native.Top,Math.Max(0,native.Right-native.Left-18),native.Bottom-native.Top);
                    TextRenderer.DrawText(graphics, Columns[index].Text, Font, text, MutedColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                    graphics.DrawLine(divider,native.Right-1,native.Top+4,native.Right-1,native.Bottom-5);
                }
                graphics.DrawLine(divider,0,client.Bottom-1,client.Right,client.Bottom-1);
            }
        }
        public void RefreshHeader() {
            if (header != null) InvalidateRect(header.Handle, IntPtr.Zero, false);
        }
        public void FitColumns() {
            if (fitting || Columns.Count != 4 || ClientSize.Width < 80) return;
            fitting = true;
            try {
                int width = Math.Max(60, ClientSize.Width-1);
                if (preferred == null) preferred = new int[]{(int)(width*.30),(int)(width*.27),(int)(width*.17)};
                // Recover collapsed/invalid saved columns instead of hiding their labels.
                int first = Math.Max(100,preferred[0]), second = Math.Max(100,preferred[1]), third = Math.Max(78,preferred[2]);
                int used = Math.Max(1, first+second+third);
                if (used > width-60) {
                    double scale = Math.Max(1,width-60)/(double)used;
                    first = (int)(first*scale); second = (int)(second*scale); third = (int)(third*scale);
                }
                Columns[0].Width = first; Columns[1].Width = second; Columns[2].Width = third;
                Columns[3].Width = width - first - second - third;
                if (IsHandleCreated) {
                    int offset = GetScrollPos(Handle,0);
                    if (offset != 0) SendMessage(Handle,0x1014,new IntPtr(-offset),IntPtr.Zero);
                }
                RefreshHeader();
                PositionNameEditor();
            } finally { fitting = false; }
        }
        public void SetPreferredColumns(int[] widths) {
            if (widths != null && widths.Length >= 3) preferred = new int[]{widths[0] < 80 ? 180 : widths[0],widths[1] < 80 ? 170 : widths[1],widths[2] < 60 ? 95 : widths[2]};
            FitColumns();
        }
        protected override void OnColumnWidthChanged(ColumnWidthChangedEventArgs e) {
            base.OnColumnWidthChanged(e);
            if (!fitting && Columns.Count == 4 && e.ColumnIndex < 3) {
                if (preferred == null) preferred = new int[]{Columns[0].Width,Columns[1].Width,Columns[2].Width};
                preferred[e.ColumnIndex] = Columns[e.ColumnIndex].Width;
                FitColumns();
            }
        }
        protected override void OnResize(EventArgs e) { base.OnResize(e); FitColumns(); }
        Rectangle NameTextBounds(ListViewItem item) {
            Rectangle cell = item.Bounds;
            cell.Width = Columns.Count > 0 ? Columns[0].Width : cell.Width;
            cell.X += 10; cell.Width = Math.Max(0,cell.Width-14);
            if (SmallImageList != null && item.ImageIndex >= 0 && item.ImageIndex < SmallImageList.Images.Count) {
                int offset = Math.Min(20,item.Bounds.Height-4)+8;
                cell.X += offset; cell.Width = Math.Max(0,cell.Width-offset);
            }
            return cell;
        }
        void PositionNameEditor() {
            if (editingItem == null || editorSurface == null || nameEditor == null || editingItem.ListView != this) return;
            Rectangle text = NameTextBounds(editingItem);
            editorSurface.SetBounds(text.Left,text.Top+1,Math.Max(8,text.Width),Math.Max(8,text.Height-2));
            int height = nameEditor.PreferredHeight;
            nameEditor.SetBounds(3,Math.Max(1,(editorSurface.Height-height)/2),Math.Max(1,editorSurface.Width-6),height);
        }
        protected override void OnBeforeLabelEdit(LabelEditEventArgs e) {
            base.OnBeforeLabelEdit(e);
            if (e.CancelEdit || e.Item < 0 || e.Item >= Items.Count) return;
            // Native label editing ignores the owner-drawn icon/text padding.
            // Keep the normal BeginEdit/F2 entry point, using one aligned editor.
            e.CancelEdit = true;
            ListViewItem item = Items[e.Item];
            BeginInvoke((MethodInvoker)delegate { if (!IsDisposed && item.ListView == this) BeginShortcutNameEdit(item); });
        }
        public void BeginShortcutNameEdit(ListViewItem item) {
            if (IsDisposed || Disposing || item == null || item.ListView != this) return;
            if (editingItem == item && nameEditor != null) { nameEditor.Focus(); return; }
            FinishShortcutNameEdit(false);
            item.EnsureVisible(); editingItem = item;
            editorSurface = new Panel { Name = "ShortcutNameEditorSurface", BackColor = BackColor, TabStop = false };
            editorSurface.Paint += delegate(object sender, PaintEventArgs args) {
                Panel surface = (Panel)sender;
                using (Pen border = new Pen(AccentColor)) args.Graphics.DrawRectangle(border,0,0,Math.Max(0,surface.Width-1),Math.Max(0,surface.Height-1));
            };
            nameEditor = new TextBox { Name = "ShortcutNameEditor", AccessibleName = "Future desktop shortcut name", BorderStyle = BorderStyle.None, Font = Font, BackColor = BackColor, ForeColor = ForeColor, Text = item.Text };
            nameEditor.KeyDown += delegate(object sender, KeyEventArgs args) {
                if (args.KeyCode != Keys.Enter && args.KeyCode != Keys.Escape) return;
                args.Handled = true; args.SuppressKeyPress = true;
                FinishShortcutNameEdit(args.KeyCode == Keys.Escape); Focus();
            };
            nameEditor.LostFocus += delegate { FinishShortcutNameEdit(false); };
            editorSurface.Controls.Add(nameEditor); Controls.Add(editorSurface);
            PositionNameEditor(); editorSurface.BringToFront();
            Invalidate(item.Bounds); nameEditor.Focus(); nameEditor.SelectAll();
        }
        void FinishShortcutNameEdit(bool cancel) {
            if (editingItem == null) return;
            ListViewItem item = editingItem; string text = nameEditor == null ? item.Text : nameEditor.Text;
            Panel surface = editorSurface;
            editingItem = null; nameEditor = null; editorSurface = null;
            if (surface != null) { Controls.Remove(surface); surface.Dispose(); }
            if (!cancel && item.ListView == this && !IsDisposed && !Disposing) {
                LabelEditEventArgs result = new LabelEditEventArgs(item.Index,text);
                base.OnAfterLabelEdit(result);
                if (!result.CancelEdit) item.Text = text;
            }
            if (!IsDisposed && !Disposing) Invalidate();
        }
        protected override void WndProc(ref Message message) {
            // Finish editing before scrolling or programmatically replacing rows.
            if (message.Msg == 0x0115 || message.Msg == 0x020A || message.Msg == 0x1009) FinishShortcutNameEdit(false);
            base.WndProc(ref message);
        }
        protected override void OnDrawColumnHeader(DrawListViewColumnHeaderEventArgs e) {
            using (Brush brush = new SolidBrush(HeaderColor)) e.Graphics.FillRectangle(brush, e.Bounds);
            using (Pen line = new Pen(DividerColor)) e.Graphics.DrawLine(line, e.Bounds.Left, e.Bounds.Bottom-1, e.Bounds.Right, e.Bounds.Bottom-1);
            using (Pen line = new Pen(DividerColor)) e.Graphics.DrawLine(line, e.Bounds.Right-1, e.Bounds.Top+4, e.Bounds.Right-1, e.Bounds.Bottom-5);
            Rectangle text = e.Bounds; text.X += 10; text.Width = Math.Max(0, text.Width - 14);
            TextRenderer.DrawText(e.Graphics, e.Header.Text, Font, text, MutedColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
        protected override void OnDrawItem(DrawListViewItemEventArgs e) { if (View != View.Details) base.OnDrawItem(e); }
        protected override void OnDrawSubItem(DrawListViewSubItemEventArgs e) {
            bool selected = e.Item.Selected;
            Color background = selected ? SelectionColor : BackColor;
            using (Brush brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            Rectangle text = e.Bounds; text.X += 10; text.Width = Math.Max(0, text.Width - 14);
            if (e.ColumnIndex == 0) {
                if (selected) using (Brush marker = new SolidBrush(AccentColor)) e.Graphics.FillRectangle(marker, e.Bounds.Left, e.Bounds.Top + 3, 2, Math.Max(0, e.Bounds.Height - 6));
                if (SmallImageList != null && e.Item.ImageIndex >= 0 && e.Item.ImageIndex < SmallImageList.Images.Count) {
                    int size = Math.Min(20, e.Bounds.Height-4);
                    e.Graphics.DrawImage(SmallImageList.Images[e.Item.ImageIndex], new Rectangle(text.X, e.Bounds.Top + (e.Bounds.Height-size)/2, size, size));
                    text.X += size + 8; text.Width = Math.Max(0, text.Width-size-8);
                }
            }
            Color foreground = e.ColumnIndex == 0 || selected ? ForeColor : MutedColor;
            if (e.ColumnIndex != 0 || e.Item != editingItem) {
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, Font, text, foreground, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            } else { PositionNameEditor(); }
            if (selected && Focused && e.ColumnIndex == Columns.Count-1) {
                Rectangle focus = e.Item.Bounds; focus.Width = Math.Max(0, ClientSize.Width-2); focus.Height -= 1;
                using (Pen pen = new Pen(AccentColor)) { pen.DashStyle = DashStyle.Dot; e.Graphics.DrawRectangle(pen, focus); }
            }
        }
    }
    public class FieldSurface : Panel {
        public Color BorderColor { get; set; }
        public Color FocusColor { get; set; }
        public FieldSurface() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Padding = new Padding(9,7,9,5); BorderColor = Color.FromArgb(68,71,90); FocusColor = Color.FromArgb(98,114,164);
        }
        protected override void OnControlAdded(ControlEventArgs e) {
            base.OnControlAdded(e); e.Control.Enter += delegate { Invalidate(); }; e.Control.Leave += delegate { Invalidate(); };
        }
        protected override void OnPaint(PaintEventArgs e) {
            base.OnPaint(e);
            using (Pen pen = new Pen(ContainsFocus ? FocusColor : BorderColor, ContainsFocus ? 2 : 1)) {
                e.Graphics.DrawRectangle(pen, 1,1,Math.Max(0,Width-3),Math.Max(0,Height-3));
            }
        }
    }
    public class ThemedCheckBox : CheckBox {
        public Color AccentColor { get; set; }
        public Color MutedColor { get; set; }
        public ThemedCheckBox() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            AccentColor = Color.FromArgb(98,114,164); MutedColor = Color.FromArgb(150,150,150);
        }
        protected override void OnCheckedChanged(EventArgs e) { base.OnCheckedChanged(e); Invalidate(); }
        protected override void OnPaint(PaintEventArgs e) {
            e.Graphics.Clear(BackColor);
            float scale = e.Graphics.DpiX / 96f;
            int size = Math.Max(14,(int)Math.Round(14*scale)), y = (Height-size)/2;
            Rectangle box = new Rectangle((int)scale,y,size,size);
            using (Brush fill = new SolidBrush(Checked ? AccentColor : BackColor)) e.Graphics.FillRectangle(fill,box);
            using (Pen border = new Pen(Checked || Focused ? AccentColor : MutedColor)) e.Graphics.DrawRectangle(border,box);
            if (Checked) using (Pen mark = new Pen(BackColor,2*scale)) e.Graphics.DrawLines(mark,new Point[]{new Point((int)(4*scale),y+(int)(7*scale)),new Point((int)(7*scale),y+(int)(10*scale)),new Point((int)(12*scale),y+(int)(4*scale))});
            int offset = (int)(23*scale);
            Rectangle text = new Rectangle(offset,0,Math.Max(0,Width-offset),Height);
            TextRenderer.DrawText(e.Graphics,Text,Font,text,Enabled ? ForeColor : MutedColor,TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            if (Focused) ControlPaint.DrawFocusRectangle(e.Graphics,text,AccentColor,BackColor);
        }
    }
    public class MenuColors : ProfessionalColorTable {
        Color surface, selection, border;
        public MenuColors(Color background, Color selected, Color edge) { surface = background; selection = selected; border = edge; UseSystemColors = false; }
        public override Color ToolStripDropDownBackground { get { return surface; } }
        public override Color ImageMarginGradientBegin { get { return surface; } }
        public override Color ImageMarginGradientMiddle { get { return surface; } }
        public override Color ImageMarginGradientEnd { get { return surface; } }
        public override Color MenuItemSelected { get { return selection; } }
        public override Color MenuItemBorder { get { return selection; } }
        public override Color MenuBorder { get { return border; } }
        public override Color SeparatorDark { get { return border; } }
        public override Color SeparatorLight { get { return border; } }
    }
}
'@ -ErrorAction Stop
}

function New-AppTextField {
    param([Windows.Forms.TextBox]$TextBox)
    Initialize-AppPresentation
    $surface = New-Object QBitLauncher.FieldSurface
    $surface.Dock = 'Fill'; $surface.Margin = New-Object Windows.Forms.Padding(0, 0, 8, 4); $surface.Tag = 'FieldSurface'
    $surface.BackColor = $Global:CurrentTheme.ControlBack; $surface.BorderColor = $Global:CurrentTheme.Border; $surface.FocusColor = $Global:CurrentTheme.Accent
    $TextBox.BorderStyle = 'None'; $TextBox.Dock = 'Fill'; $TextBox.BackColor = $surface.BackColor; $TextBox.ForeColor = $Global:CurrentTheme.TextFore
    $surface.Controls.Add($TextBox)
    return $surface
}

function Set-AppFormAppearance {
    param([Windows.Forms.Form]$Form)
    Initialize-AppPresentation
    $colors = $Global:CurrentTheme
    [QBitLauncher.Presentation]::TitleBar($Form, ($Global:ThemeSelection -ne 'Light'), $colors.FormBack, $colors.TextFore)
    $pending = New-Object 'System.Collections.Generic.Stack[System.Windows.Forms.Control]'; $pending.Push($Form)
    while ($pending.Count) {
        $control = $pending.Pop(); $control.ForeColor = $colors.TextFore
        if ($control -is [Windows.Forms.Button]) { Set-ThemedButton $control }
        elseif ($control -is [Windows.Forms.TextBox] -or $control -is [Windows.Forms.ListView] -or $control -is [Windows.Forms.ListBox] -or $control -is [Windows.Forms.ComboBox] -or $control.Tag -eq 'FieldSurface') { $control.BackColor = $colors.ControlBack }
        else { $control.BackColor = $colors.FormBack }
        if ($control.Tag -eq 'Muted') { $control.ForeColor = $colors.SecondaryText }
        if ($control -is [QBitLauncher.ThemedListView]) {
            $control.Dark = $Global:ThemeSelection -ne 'Light'; $control.HeaderColor = $colors.FormBack; $control.MutedColor = $colors.SecondaryText
            $control.SelectionColor = $colors.Border; $control.AccentColor = $colors.Accent; $control.DividerColor = $colors.Border
            [void][QBitLauncher.Presentation]::SetWindowTheme($control.Handle, $(if ($control.Dark) { 'DarkMode_Explorer' } else { 'Explorer' }), $null)
            $control.Invalidate(); $control.RefreshHeader()
        }
        if ($control -is [QBitLauncher.FieldSurface]) { $control.BorderColor = $colors.Border; $control.FocusColor = $colors.Accent; $control.Invalidate() }
        if ($control -is [QBitLauncher.ThemedCheckBox]) { $control.AccentColor = $colors.Accent; $control.MutedColor = $colors.SecondaryText; $control.Invalidate() }
        foreach ($child in $control.Controls) { $pending.Push($child) }
    }
}

function Start-AppWorker {
    param([scriptblock]$Action, [object[]]$Arguments = @(), [string[]]$Functions = @(), [hashtable]$Variables = @{})
    $definitions = foreach ($name in $Functions) { "function $name {`n$((Get-Command $name -CommandType Function -ErrorAction Stop).Definition)`n}" }
    $code = 'param($TaskArguments, $TaskVariables)' + "`n" + ($definitions -join "`n") + "`n" +
        'foreach ($key in $TaskVariables.Keys) { Set-Variable -Name $key -Value $TaskVariables[$key] -Scope Global }; $ErrorActionPreference = "Stop"; & {' + $Action.ToString() + '} @TaskArguments'
    $powershell = [Management.Automation.PowerShell]::Create()
    [void]$powershell.AddScript($code).AddArgument($Arguments).AddArgument($Variables)
    return @{ PowerShell = $powershell; Handle = $powershell.BeginInvoke(); Disposed = $false; Result = $null }
}

function Receive-AppWorker {
    param($Task)
    if ($Task.Result) { return $Task.Result }
    if (-not $Task.Handle.IsCompleted) { return $null }
    try {
        $values = @($Task.PowerShell.EndInvoke($Task.Handle))
        if ($Task.PowerShell.HadErrors) { throw ($Task.PowerShell.Streams.Error | Select-Object -First 1).Exception }
        $Task.Result = @{ Values = $values; Error = $null }
    } catch { $Task.Result = @{ Values = @(); Error = $_.Exception.Message } }
    finally { $Task.PowerShell.Dispose(); $Task.Disposed = $true }
    return $Task.Result
}

function Stop-AppWorker {
    param($Task)
    if ($Task -and -not $Task.Disposed) {
        try { $Task.PowerShell.Stop() } catch {}
        $Task.PowerShell.Dispose(); $Task.Disposed = $true
    }
}

function Invoke-ResponsiveTask {
    param([string]$Title, [scriptblock]$Action, [object[]]$Arguments = @(), [string[]]$Functions = @(), [hashtable]$Variables = @{}, [switch]$NoCancel)
    $task = Start-AppWorker -Action $Action -Arguments $Arguments -Functions $Functions -Variables $Variables
    $form = New-Object Windows.Forms.Form
    $form.Text = $Title; $form.ClientSize = New-Object Drawing.Size(480, 145)
    $form.StartPosition = 'CenterParent'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false
    $form.BackColor = $Global:CurrentTheme.FormBack; $form.ForeColor = $Global:CurrentTheme.TextFore
    $label = New-Object Windows.Forms.Label
    $label.Text = $Title; $label.Dock = 'Top'; $label.Height = 45; $label.Padding = New-Object Windows.Forms.Padding(16)
    $progress = New-Object Windows.Forms.ProgressBar; $progress.Dock = 'Top'; $progress.Style = 'Marquee'; $progress.Height = 18
    $cancel = New-Object Windows.Forms.Button; $cancel.Text = 'Cancel'; $cancel.Dock = 'Bottom'; $cancel.Height = 34
    $cancel.Enabled = -not $NoCancel; Set-ThemedButton $cancel
    $cancel.Add_Click({ $form.Close() }.GetNewClosure())
    $form.Controls.Add($cancel); $form.Controls.Add($progress); $form.Controls.Add($label)
    $state = @{ Result = $null }
    if ($NoCancel) {
        $form.ControlBox = $false
        $form.Add_FormClosing({ param($sender, $e) if (-not $state.Result -and $e.CloseReason -eq 'UserClosing') { $e.Cancel = $true } }.GetNewClosure())
    }
    $timer = New-Object Windows.Forms.Timer; $timer.Interval = 75
    $timer.Add_Tick({ $result = Receive-AppWorker $task; if ($result) { $state.Result = $result; $timer.Stop(); $form.Close() } }.GetNewClosure())
    $form.Add_Shown({ Set-FormIcon $form; $timer.Start() }.GetNewClosure())
    try {
        if ($Global:MainForm -and -not $Global:MainForm.IsDisposed) { [void]$form.ShowDialog($Global:MainForm) } else { [void]$form.ShowDialog() }
        if (-not $state.Result) { throw [OperationCanceledException]::new('Operation cancelled.') }
        if ($state.Result.Error) { throw $state.Result.Error }
        return $state.Result.Values
    } finally { $timer.Stop(); $timer.Dispose(); Stop-AppWorker $task; $form.Dispose() }
}

function Add-UpdatePolling {
    param([Windows.Forms.Form]$Form, [scriptblock]$OnAvailable)
    $timer = New-Object Windows.Forms.Timer; $timer.Interval = 1000
    $timer.Add_Tick({
        if ($Global:PendingUpdateCheck) {
            $result = Receive-BackgroundUpdateCheck -Check $Global:PendingUpdateCheck
            if ($result) {
                $Global:PendingUpdateCheck = $null
                if ($result.Error) { Write-LogMessage ('Update check: ' + $result.Error) }
                if ($result.Version) { $Global:AvailableVersion = $result.Version; if ($OnAvailable) { & $OnAvailable $result.Version } }
                $timer.Stop()
            }
        } else { $timer.Stop() }
    }.GetNewClosure())
    $Form.Add_Shown({ $timer.Start() }.GetNewClosure())
    $Form.Add_FormClosed({ $timer.Stop(); $timer.Dispose() }.GetNewClosure())
}

function Get-InventoryResponsive {
    param([string]$Path)
    return Invoke-ResponsiveTask -Title 'Scanning files — you can cancel this scan' -Action { param($path) Get-FolderInventory $path } -Arguments @($Path) -Functions @('Get-FolderInventory', 'Get-RunnableRank')
}

function Get-AllRunnables {
    param([string]$RootFolderPath)
    $inventory = Get-InventoryResponsive $RootFolderPath
    return $inventory.Runnables
}

function Show-ArchiveChoice {
    param([IO.FileInfo[]]$Files, [bool]$HasRunnables = $false)
    $form = New-Object Windows.Forms.Form
    $form.Text = 'Choose an archive'; $form.Size = New-Object Drawing.Size(760, 390)
    $form.MinimumSize = New-Object Drawing.Size(560, 300); $form.StartPosition = 'CenterScreen'
    $form.BackColor = $Global:CurrentTheme.FormBack; $form.ForeColor = $Global:CurrentTheme.TextFore; $form.AutoScaleMode = 'Dpi'
    $list = New-Object Windows.Forms.ListBox; $list.Dock = 'Fill'; $list.HorizontalScrollbar = $true
    $list.BackColor = $Global:CurrentTheme.ControlBack; $list.ForeColor = $Global:CurrentTheme.TextFore
    foreach ($file in $Files) { [void]$list.Items.Add($file.FullName) }; if ($list.Items.Count) { $list.SelectedIndex = 0 }
    $buttons = New-Object Windows.Forms.FlowLayoutPanel; $buttons.Dock = 'Bottom'; $buttons.Height = 48
    $result = @{ Value = $null }
    $select = New-Object Windows.Forms.Button; $select.Text = 'Extract selected'; $select.AutoSize = $true; Set-ThemedButton $select
    $select.Add_Click({ if ($list.SelectedIndex -ge 0) { $result.Value = $Files[$list.SelectedIndex]; $form.Close() } }.GetNewClosure())
    $skip = New-Object Windows.Forms.Button; $skip.Text = 'Show executables'; $skip.AutoSize = $true; $skip.Enabled = $HasRunnables; Set-ThemedButton $skip
    $skip.Add_Click({ $result.Value = 'Runnables'; $form.Close() }.GetNewClosure())
    $cancel = New-Object Windows.Forms.Button; $cancel.Text = 'Cancel'; $cancel.AutoSize = $true; Set-ThemedButton $cancel
    $cancel.Add_Click({ $form.Close() }.GetNewClosure())
    $buttons.Controls.AddRange(@($select, $skip, $cancel)); $form.Controls.Add($list); $form.Controls.Add($buttons)
    try { [void]$form.ShowDialog(); return $result.Value } finally { $form.Dispose() }
}

function Show-CleanupConfirmForm {
    param([string]$TorrentLabel, [string]$TorrentHash, [string]$SourcePath, [string]$RootFolder, [string]$SelectedPath)
    Initialize-AppPresentation
    $form = New-Object Windows.Forms.Form
    $form.Name = 'CleanupForm'; $form.Text = 'Cleanup'; $form.Size = New-Object Drawing.Size(720, 590)
    $form.MinimumSize = New-Object Drawing.Size(650, 530); $form.AutoScaleMode = 'Dpi'; $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 10)
    $form.BackColor = $Global:CurrentTheme.FormBack; $form.ForeColor = $Global:CurrentTheme.TextFore
    $layout = New-Object Windows.Forms.TableLayoutPanel; $layout.Dock = 'Fill'; $layout.ColumnCount = 1; $layout.RowCount = 6; $layout.Padding = New-Object Windows.Forms.Padding(16)
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    foreach ($spec in @(@('Absolute', 32), @('Absolute', 204), @('Percent', 100), @('Absolute', 0), @('Absolute', 34), @('Absolute', 44))) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle($spec[0], $spec[1]))) }
    $details = New-Object Windows.Forms.Label; $details.Text = 'Select what to remove. Unticked items stay.'; $details.Dock = 'Fill'; $details.TextAlign = 'MiddleLeft'
    $choices = New-Object Windows.Forms.TableLayoutPanel; $choices.Dock = 'Fill'; $choices.ColumnCount = 1; $choices.RowCount = 3
    [void]$choices.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    $source = New-Object QBitLauncher.ThemedCheckBox; $source.Name = 'RemoveSourceFiles'; $source.Text = 'Source files'; $source.Dock = 'Top'; $source.Height = 28
    $source.Enabled = [bool]($SourcePath -and (Test-Path -LiteralPath $SourcePath))
    $extracted = New-Object QBitLauncher.ThemedCheckBox; $extracted.Name = 'RemoveExtractedFiles'; $extracted.Text = 'Extracted files'; $extracted.Dock = 'Top'; $extracted.Height = 28
    $extracted.Enabled = [bool]($RootFolder -and (Test-Path -LiteralPath $RootFolder))
    $torrent = New-Object QBitLauncher.ThemedCheckBox; $torrent.Name = 'RemoveTorrentEntry'; $torrent.Text = 'qBittorrent entry'; $torrent.Dock = 'Top'; $torrent.Height = 28
    $torrent.Enabled = -not [string]::IsNullOrWhiteSpace($TorrentHash)
    $tip = New-Object Windows.Forms.ToolTip
    $row = 0
    foreach ($option in @(@($source, $(if ($SourcePath) { $SourcePath } else { 'No source location' })), @($extracted, $(if ($RootFolder) { $RootFolder } else { 'No extracted location' })), @($torrent, $(if ($torrent.Enabled) { 'Stop seeding and remove the entry. Files are handled above.' } else { 'Available when opened from qBittorrent.' })))) {
        [void]$choices.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 33.333)))
        $panel = New-Object Windows.Forms.Panel; $panel.Dock = 'Fill'; $panel.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 8)
        $location = New-Object Windows.Forms.Label; $location.Text = $option[1]; $location.Name = $option[0].Name + 'Location'; $location.Dock = 'Fill'; $location.Padding = New-Object Windows.Forms.Padding(23, 0, 0, 0)
        $location.AutoEllipsis = $true; $location.UseMnemonic = $false; $location.Tag = 'Muted'; $tip.SetToolTip($location, $option[1])
        $panel.Controls.Add($location); $panel.Controls.Add($option[0]); $choices.Controls.Add($panel, 0, $row); $row++
    }
    $preview = New-Object Windows.Forms.TextBox; $preview.Name = 'CleanupPreview'; $preview.Multiline = $true; $preview.ReadOnly = $true; $preview.ScrollBars = 'None'; $preview.BorderStyle = 'None'; $preview.Dock = 'Fill'; $preview.WordWrap = $true
    $preview.BackColor = $Global:CurrentTheme.ControlBack; $preview.ForeColor = $Global:CurrentTheme.TextFore
    $ack = New-Object QBitLauncher.ThemedCheckBox; $ack.Name = 'AcknowledgeAppDeletion'; $ack.Text = 'Remove the extracted app / installer too.'; $ack.Dock = 'Fill'; $ack.Visible = $false
    $permanent = New-Object QBitLauncher.ThemedCheckBox; $permanent.Name = 'PermanentlyDeleteFiles'; $permanent.Text = 'Permanently delete files (skip Recycle Bin)'; $permanent.Dock = 'Fill'; $permanent.Enabled = $false
    $actions = New-Object Windows.Forms.FlowLayoutPanel; $actions.Dock = 'Fill'; $actions.FlowDirection = 'RightToLeft'
    $confirm = New-Object Windows.Forms.Button; $confirm.Name = 'ConfirmCleanup'; $confirm.Text = 'Remove selected'; $confirm.AutoSize = $true; $confirm.MinimumSize = New-Object Drawing.Size(130, 34); $confirm.Enabled = $false; Set-DestructiveButton $confirm
    $cancel = New-Object Windows.Forms.Button; $cancel.Text = 'Cancel'; $cancel.AutoSize = $true; $cancel.MinimumSize = New-Object Drawing.Size(80, 34); Set-ThemedButton $cancel
    $actions.Controls.AddRange(@($confirm, $cancel))
    $state = @{ Plan = $null; Task = $null; NeedsAck = $false; Error = $false; Result = $null }
    $updateConfirm = { $confirm.Enabled = -not $state.Task -and -not $state.Error -and (-not $state.NeedsAck -or $ack.Checked) -and ($torrent.Checked -or ($state.Plan -and $state.Plan.Targets.Count -gt 0)) }.GetNewClosure()
    $renderPreview = {
        $lines = New-Object 'System.Collections.Generic.List[string]'
        if ($state.Plan -and $state.Plan.Targets.Count) {
            $verb = if ($permanent.Checked) { 'Permanently delete' } else { 'Recycle' }
            $items = if ($source.Checked -and $extracted.Checked) { 'source and extracted files' } elseif ($source.Checked) { 'source files' } else { 'extracted files' }
            $lines.Add("$verb $items.")
            # Paths are already shown in the checklist. List exact children only
            # when the plan splits a parent around files that must be kept.
            $requested = @($state.Plan.ApprovedRoots | ForEach-Object { Get-CanonicalPath $_ })
            $splitTargets = @($state.Plan.Targets | Where-Object { $_.CanonicalPath -notin $requested })
            foreach ($target in $splitTargets) { $lines.Add($target.Path) }
            if ($state.Plan.Retained.Count) {
                $lines.Add($(if ($state.Plan.Mode -eq 'SourceOnly') { 'Keep extracted files and any junctions inside the source location.' } else { 'Keep junctions inside the selected locations.' }))
            }
        }
        elseif ($source.Checked -or $extracted.Checked) { $lines.Add('No separate files can be removed while keeping the unticked locations.') }
        if ($torrent.Checked) { $lines.Add('Remove the entry from qBittorrent.') }
        if ($state.NeedsAck) { $lines.Add(''); $lines.Add('The extracted app / installer will also be removed.') }
        $preview.Text = if ($lines.Count) { $lines -join "`r`n" } else { 'Nothing selected.' }
        $ack.Visible = [bool]$state.NeedsAck; $layout.RowStyles[3].Height = if ($state.NeedsAck) { 34 } else { 0 }
        $permanent.Enabled = [bool]($state.Plan -and $state.Plan.Targets.Count)
        & $updateConfirm
    }.GetNewClosure()
    $functions = @('Get-CleanupPlan', 'Get-CanonicalPath', 'Test-PathContains', 'Test-CleanupRoot', 'Test-TreeHasReparsePoint')
    $selectionChanged = {
        Stop-AppWorker $state.Task; $state.Task = $null; $state.Plan = $null; $state.Error = $false; $state.NeedsAck = $false; $ack.Checked = $false
        if (-not $source.Checked -and -not $extracted.Checked) { & $renderPreview; return }
        $mode = if ($source.Checked -and $extracted.Checked) { 'Everything' } elseif ($source.Checked) { 'SourceOnly' } else { 'ExtractedOnly' }
        $preview.Text = 'Checking selected locations...'; $confirm.Enabled = $false; $ack.Visible = $false; $layout.RowStyles[3].Height = 0; $permanent.Enabled = $false
        $state.Task = Start-AppWorker -Action { param($sourcePath, $root, $selected, $mode) Get-CleanupPlan -SourcePath $sourcePath -ExtractedPath $root -SelectedPath $selected -Mode $mode } -Arguments @($SourcePath, $RootFolder, $SelectedPath, $mode) -Functions $functions -Variables @{ ScriptDir = $Global:ScriptDir }
    }.GetNewClosure()
    foreach ($control in @($source, $extracted, $torrent)) { $control.Add_CheckedChanged($selectionChanged) }
    $timer = New-Object Windows.Forms.Timer; $timer.Interval = 100
    $timer.Add_Tick({
        if ($state.Task) {
            $result = Receive-AppWorker $state.Task
            if ($result) {
                $state.Task = $null; $state.Error = [bool]$result.Error
                if ($result.Error) { $preview.Text = $result.Error; & $updateConfirm }
                else {
                    $state.Plan = $result.Values[0]; $state.NeedsAck = $state.Plan.RequiresAcknowledgement
                    & $renderPreview
                }
            }
        }
    }.GetNewClosure())
    $ack.Add_CheckedChanged($updateConfirm)
    $permanent.Add_CheckedChanged($renderPreview)
    $confirm.Add_Click({
        & $updateConfirm
        if (-not $confirm.Enabled) { return }
        $state.Result = @{ RemoveSource = $source.Checked; RemoveExtracted = $extracted.Checked; RemoveTorrent = $torrent.Checked; Plan = $state.Plan; Acknowledged = $ack.Checked; Permanent = $permanent.Checked }
        $form.Close()
    }.GetNewClosure())
    $cancel.Add_Click({ $form.Close() }.GetNewClosure())
    $row = 0
    foreach ($control in @($details, $choices, $preview, $ack, $permanent, $actions)) { $layout.Controls.Add($control, 0, $row); $row++ }
    $form.Controls.Add($layout)
    $form.Add_Shown({ Set-FormIcon $form; Set-DestructiveButton $confirm; & $renderPreview; $timer.Start(); $source.Focus() | Out-Null; $form.Tag = 'Ready' }.GetNewClosure())
    try { if ($Global:MainForm) { [void]$form.ShowDialog($Global:MainForm) } else { [void]$form.ShowDialog() }; return $state.Result }
    finally { $timer.Stop(); $timer.Dispose(); Stop-AppWorker $state.Task; $tip.Dispose(); $form.Dispose() }
}

function Invoke-CleanupSelection {
    param($Choice, [string]$TorrentHash, [scriptblock]$OnMessage = {})
    if (-not ($Choice.RemoveSource -or $Choice.RemoveExtracted -or $Choice.RemoveTorrent)) { throw 'Select at least one cleanup item.' }
    $local = $Choice.RemoveSource -or $Choice.RemoveExtracted
    if ($local -and -not $Choice.Plan) { throw 'Prepare a file preview before cleanup.' }
    if ($local -and $Choice.Plan.RequiresAcknowledgement -and -not $Choice.Acknowledged) { throw 'Confirmation is required to remove the extracted app / installer.' }
    if ($Choice.RemoveTorrent -and [string]::IsNullOrWhiteSpace($TorrentHash)) { throw 'No qBittorrent entry is available for this source.' }
    $result = @{ TorrentRemoved = $false; LocalResult = $null }
    if ($Choice.RemoveTorrent) {
        # Release seeding handles first. File deletion belongs to the local plan,
        # so selecting all three never makes qBittorrent delete the files twice.
        $request = { param($hash) Remove-QBittorrentTorrent -TorrentHash $hash -DeleteFiles $false }
        try { [void](Invoke-ResponsiveTask -Title 'Removing qBittorrent entry' -Action $request -Arguments @($TorrentHash) -Functions @('Remove-QBittorrentTorrent') -Variables @{ UserSettings = $Global:UserSettings } -NoCancel) }
        catch {
            $repair = Invoke-QBittorrentLocalApiRepairFlow -ErrorMessage $_.Exception.Message -RetryAction { Remove-QBittorrentTorrent -TorrentHash $TorrentHash -DeleteFiles $false }.GetNewClosure()
            if ($repair.Status -ne 'RetrySucceeded') { throw $repair.Message }
        }
        $result.TorrentRemoved = $true; & $OnMessage 'qBittorrent entry removed'
    }
    if ($local -and $Choice.Plan.Targets.Count) {
        $functions = @('Invoke-CleanupPlan', 'Get-CanonicalPath', 'Test-PathContains', 'Test-CleanupRoot', 'Test-TreeHasReparsePoint')
        $result.LocalResult = Invoke-ResponsiveTask -Title 'Removing selected files' -Action { param($plan, $ack, $permanent) Invoke-CleanupPlan -Plan $plan -AcknowledgeSelectedFiles:$ack -PermanentlyDelete:$permanent } -Arguments @($Choice.Plan, $Choice.Acknowledged, $Choice.Permanent) -Functions $functions -Variables @{ ScriptDir = $Global:ScriptDir } -NoCancel
        & $OnMessage ("$($result.LocalResult.Deleted.Count) locations removed; method: " + $(if ($Choice.Permanent) { 'permanent' } else { 'Recycle Bin' }))
    }
    return $result
}

function Show-ExecutableSelectionForm {
    param([Management.Automation.PSObject[]]$FoundExecutables, [string]$WindowTitle = 'qBitLauncher', [string]$RootFolder, [string]$TorrentHash, [string]$TorrentName)
    if (-not $RootFolder -and $FoundExecutables.Count) { $RootFolder = $FoundExecutables[0].DirectoryName }
    [void](Add-RecentFolder $RootFolder)
    Initialize-AppPresentation
    $colors = $Global:CurrentTheme
    $sourcePath = $filePathFromQB
    $rankedFiles = @($FoundExecutables | Sort-Object @{ Expression = { (Get-RunnableRank $_ $RootFolder).Score }; Descending = $true }, FullName)
    $state = @{ SelectedPath = ''; LoadingPreference = $false; LastError = ''; IconTask = $null; Icons = @{}; AllFiles = $rankedFiles; Closing = $false }
    $form = New-Object Windows.Forms.Form; $form.Name = 'ExecutableSelectionForm'; $form.Text = $WindowTitle
    $form.Size = New-Object Drawing.Size(1000, 620); $form.MinimumSize = New-Object Drawing.Size(780, 500)
    $form.StartPosition = 'CenterScreen'; $form.AutoScaleMode = 'Dpi'; $form.Font = New-Object Drawing.Font('Segoe UI', 10)
    $form.BackColor = $colors.FormBack; $form.ForeColor = $colors.TextFore
    if ($Global:UserSettings.WindowState.Width -and $Global:UserSettings.WindowState.Height) {
        $area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $form.Size = New-Object Drawing.Size([Math]::Max(780, [Math]::Min([int]$Global:UserSettings.WindowState.Width, $area.Width)), [Math]::Max(500, [Math]::Min([int]$Global:UserSettings.WindowState.Height, $area.Height)))
    }
    $Global:MainForm = $form
    $layout = New-Object Windows.Forms.TableLayoutPanel; $layout.Dock = 'Fill'; $layout.Padding = New-Object Windows.Forms.Padding(18, 14, 18, 10); $layout.ColumnCount = 1; $layout.RowCount = 6
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    foreach ($spec in @(@('Absolute', 56), @('Absolute', 42), @('Percent', 100), @('Absolute', 0), @('Absolute', 52), @('Absolute', 22))) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle($spec[0], $spec[1]))) }
    $header = New-Object Windows.Forms.Panel; $header.Dock = 'Fill'
    $heading = New-Object Windows.Forms.Label; $heading.Dock = 'Top'; $heading.Height = 24; $heading.Font = New-Object Drawing.Font('Segoe UI', 12, [Drawing.FontStyle]::Bold)
    $heading.AutoEllipsis = $true; $heading.Text = [IO.Path]::GetFileName($RootFolder.TrimEnd([char]92)); $heading.UseMnemonic = $false
    $pathLabel = New-Object Windows.Forms.Label; $pathLabel.Dock = 'Bottom'; $pathLabel.Height = 24; $pathLabel.Font = New-Object Drawing.Font('Segoe UI', 9)
    $pathLabel.AutoEllipsis = $true; $pathLabel.Text = $RootFolder; $pathLabel.UseMnemonic = $false; $pathLabel.Tag = 'Muted'
    $header.Controls.Add($heading); $header.Controls.Add($pathLabel)
    $tip = New-Object Windows.Forms.ToolTip; $tip.SetToolTip($heading, $RootFolder)
    $filterRow = New-Object Windows.Forms.TableLayoutPanel; $filterRow.Dock = 'Fill'; $filterRow.ColumnCount = 3; $filterRow.RowCount = 1
    [void]$filterRow.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))
    [void]$filterRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 56))); [void]$filterRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100))); [void]$filterRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 135)))
    $searchLabel = New-Object Windows.Forms.Label; $searchLabel.Text = 'Search'; $searchLabel.Dock = 'Fill'; $searchLabel.TextAlign = 'MiddleLeft'; $searchLabel.Tag = 'Muted'
    $search = New-Object Windows.Forms.TextBox; $search.Name = 'SearchBox'; $search.AccessibleName = 'Search executables and shortcut names'
    $searchField = New-AppTextField $search
    $showAll = New-Object QBitLauncher.ThemedCheckBox; $showAll.Name = 'ShowAllCheckBox'; $showAll.Text = 'Show helpers'; $showAll.Dock = 'Fill'
    $tip.SetToolTip($showAll, 'Include uninstallers, crash handlers, and maintenance tools.')
    $filterRow.Controls.AddRange(@($searchLabel, $searchField, $showAll))
    $split = New-Object Windows.Forms.SplitContainer; $split.Size = New-Object Drawing.Size(1040, 400); $split.Dock = 'Fill'; $split.SplitterDistance = 700; $split.SplitterWidth = 16; $split.Panel1MinSize = 350; $split.Panel2MinSize = 180
    $split.Add_Paint({
        param($sender, $eventArgs)
        $divider = $sender.SplitterRectangle; $x = $divider.Left + [int]($divider.Width / 2)
        $pen = New-Object Drawing.Pen($Global:CurrentTheme.Border)
        try { $eventArgs.Graphics.DrawLine($pen, $x, $divider.Top, $x, $divider.Bottom) } finally { $pen.Dispose() }
    })
    $list = New-Object QBitLauncher.ThemedListView; $list.Name = 'RunnableList'; $list.Dock = 'Fill'; $list.View = 'Details'; $list.FullRowSelect = $true; $list.MultiSelect = $false
    $list.HideSelection = $false; $list.LabelEdit = $true; $list.ShowItemToolTips = $true; $list.BackColor = $colors.ControlBack; $list.ForeColor = $colors.TextFore
    foreach ($column in @(@('Shortcut name', 180), @('Original file', 180), @('Type', 140), @('Folder', 350))) { [void]$list.Columns.Add($column[0], $column[1]) }
    $list.SetPreferredColumns(@(180, 170, 95))
    if ($Global:UserSettings.WindowState.Columns -and $Global:UserSettings.WindowState.Columns.Count -eq 4) { $list.SetPreferredColumns([int[]]$Global:UserSettings.WindowState.Columns) }
    $imageList = New-Object Windows.Forms.ImageList; $imageList.ImageSize = New-Object Drawing.Size(24, 24); $imageList.ColorDepth = 'Depth32Bit'; $list.SmallImageList = $imageList
    $log = New-Object Windows.Forms.TextBox; $log.Name = 'ActivityLog'; $log.Multiline = $true; $log.ReadOnly = $true; $log.BorderStyle = 'None'; $log.ScrollBars = 'None'; $log.Dock = 'Fill'; $log.BackColor = $colors.ControlBack; $log.ForeColor = $colors.SecondaryText; $log.Tag = 'Muted'
    $log.Font = New-Object Drawing.Font('Segoe UI', 9)
    $logSurface = New-Object Windows.Forms.Panel; $logSurface.Dock = 'Fill'; $logSurface.Padding = New-Object Windows.Forms.Padding(10); $logSurface.Tag = 'FieldSurface'; $logSurface.Controls.Add($log)
    $logTitle = New-Object Windows.Forms.Label; $logTitle.Text = 'Activity'; $logTitle.Dock = 'Top'; $logTitle.Height = 26; $logTitle.Font = New-Object Drawing.Font('Segoe UI', 9); $logTitle.Tag = 'Muted'
    $logPanel = New-Object Windows.Forms.Panel; $logPanel.Dock = 'Fill'; $logPanel.Controls.Add($logSurface); $logPanel.Controls.Add($logTitle)
    $split.Panel1.Controls.Add($list); $split.Panel2.Controls.Add($logPanel)
    $status = New-Object Windows.Forms.Label; $status.Dock = 'Fill'; $status.Tag = 'Muted'; $status.Font = New-Object Drawing.Font('Segoe UI', 9); $status.Text = 'F2 changes the future shortcut name. Original files stay unchanged.'
    $addLog = { param($message) $safe = Protect-LogText ([string]$message); $log.AppendText("$(Get-Date -Format 'HH:mm')   $safe`r`n`r`n"); $log.SelectionStart = $log.TextLength; $log.ScrollToCaret(); Write-LogMessage $safe }.GetNewClosure()
    $launchRow = New-Object Windows.Forms.TableLayoutPanel; $launchRow.Dock = 'Fill'; $launchRow.ColumnCount = 3; $launchRow.RowCount = 1
    [void]$launchRow.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))
    [void]$launchRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 135))); [void]$launchRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100))); [void]$launchRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 155)))
    $argLabel = New-Object Windows.Forms.Label; $argLabel.Text = 'Launch arguments'; $argLabel.Dock = 'Fill'; $argLabel.TextAlign = 'MiddleLeft'; $argLabel.Tag = 'Muted'
    $argumentsBox = New-Object Windows.Forms.TextBox; $argumentsBox.Name = 'LaunchArgumentsTextBox'; $argumentsBox.AccessibleName = 'Launch arguments'
    $argumentsField = New-AppTextField $argumentsBox
    $asAdmin = New-Object QBitLauncher.ThemedCheckBox; $asAdmin.Name = 'RunAsAdminCheckBox'; $asAdmin.Text = 'Administrator'; $asAdmin.Dock = 'Fill'
    $launchRow.Controls.AddRange(@($argLabel, $argumentsField, $asAdmin)); $launchRow.Visible = $false
    $buttons = New-Object Windows.Forms.FlowLayoutPanel; $buttons.Dock = 'Fill'; $buttons.AutoScroll = $false; $buttons.WrapContents = $true; $buttons.Padding = New-Object Windows.Forms.Padding(0, 10, 0, 0)
    $buttonMap = @{}
    foreach ($spec in @(@('Run', '&Run'), @('Shortcut', '&Shortcut'), @('ShortcutLabel', 'Shortcut &Name'), @('OpenFolder', '&Open Folder'), @('Cleanup', 'Cleanup'), @('Settings', 'Se&ttings'), @('OpenLog', 'Open Log'), @('CopyError', 'Copy Error'), @('Close', '&Close'))) {
        $button = New-Object Windows.Forms.Button; $button.Name = $spec[0] + 'Button'; $button.Text = $spec[1]; $button.AutoSize = $true; $button.MinimumSize = New-Object Drawing.Size(76, 34); $button.Margin = New-Object Windows.Forms.Padding(0, 0, 6, 4)
        Set-ThemedButton $button; [void]$buttons.Controls.Add($button); $buttonMap[$spec[0]] = $button
    }
    $buttonMap.Run.Tag = 'Primary'; Set-ThemedButton $buttonMap.Run
    foreach ($name in @('Settings', 'OpenLog', 'CopyError')) { $buttonMap[$name].Visible = $false }
    $more = New-Object Windows.Forms.Button; $more.Name = 'MoreButton'; $more.Text = '&More'; $more.AutoSize = $true; $more.MinimumSize = New-Object Drawing.Size(76, 34); $more.Margin = New-Object Windows.Forms.Padding(0, 0, 6, 4); Set-ThemedButton $more
    $buttons.Controls.Add($more); $buttons.Controls.SetChildIndex($more, $buttons.Controls.IndexOf($buttonMap.Close))
    $moreMenu = New-Object Windows.Forms.ContextMenuStrip; $moreMenu.ShowImageMargin = $false; $moreMenu.ForeColor = $colors.TextFore
    $moreMenu.Renderer = New-Object Windows.Forms.ToolStripProfessionalRenderer((New-Object QBitLauncher.MenuColors($colors.ControlBack, $colors.ButtonBack, $colors.Border)))
    $optionsItem = $moreMenu.Items.Add('Launch options'); $optionsItem.CheckOnClick = $true
    $optionsItem.Add_CheckedChanged({ $launchRow.Visible = $optionsItem.Checked; $layout.RowStyles[3].Height = $(if ($optionsItem.Checked) { 42 } else { 0 }) }.GetNewClosure())
    [void]$moreMenu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
    foreach ($name in @('Settings', 'OpenLog', 'CopyError')) {
        $menuItem = $moreMenu.Items.Add($buttonMap[$name].Text.Replace('&', '')); $menuItem.Tag = $name
        $menuItem.Add_Click({
            $button = $buttonMap[[string]$this.Tag]
            $onClick = [Windows.Forms.Button].GetMethod('OnClick', [Reflection.BindingFlags]'NonPublic,Instance')
            [void]$onClick.Invoke($button, @([EventArgs]::Empty))
        }.GetNewClosure())
    }
    $more.ContextMenuStrip = $moreMenu
    $more.Add_Click({ $moreMenu.Show($more, (New-Object Drawing.Point(0, $more.Height))) }.GetNewClosure())
    Set-DestructiveButton $buttonMap.Cleanup
    $layoutRow = 0
    foreach ($control in @($header, $filterRow, $split, $launchRow, $buttons, $status)) { $layout.Controls.Add($control, 0, $layoutRow); $layoutRow++ }; $form.Controls.Add($layout)
    $savePreference = {
        if ($state.SelectedPath -and -not $state.LoadingPreference) {
            [void](Set-FilePreference -Path $state.SelectedPath -Preference @{ Arguments = $argumentsBox.Text; AsAdmin = [bool]$asAdmin.Checked })
        }
    }.GetNewClosure()
    $preferenceTimer = New-Object Windows.Forms.Timer; $preferenceTimer.Interval = 500
    $preferenceTimer.Add_Tick({ $preferenceTimer.Stop(); & $savePreference }.GetNewClosure())
    $changedPreference = { if (-not $state.LoadingPreference) { $preferenceTimer.Stop(); $preferenceTimer.Start() } }.GetNewClosure()
    $argumentsBox.Add_TextChanged($changedPreference); $asAdmin.Add_CheckedChanged($changedPreference)
    $asAdmin.Add_CheckedChanged({ $buttonMap.Run.Text = if ($asAdmin.Checked) { '&Run as admin' } else { '&Run' } }.GetNewClosure())
    $list.Add_SelectedIndexChanged({
        if (-not $list.SelectedItems.Count) { return }
        $preferenceTimer.Stop(); & $savePreference
        $state.SelectedPath = [string]$list.SelectedItems[0].Tag
        $preference = Get-FilePreference $state.SelectedPath; $state.LoadingPreference = $true
        $argumentsBox.Text = [string]$preference.Arguments
        $asAdmin.Checked = if ($preference.ContainsKey('AsAdmin')) { [bool]$preference.AsAdmin } else { [bool]$Global:UserSettings.DefaultAdmin }
        $state.LoadingPreference = $false
        [void](Set-FilePreference -Path $RootFolder -Preference @{ PreferredExecutable = $state.SelectedPath })
    }.GetNewClosure())
    $render = {
        $wanted = $state.SelectedPath
        if (-not $wanted) { $wanted = (Get-FilePreference $RootFolder).PreferredExecutable }
        $list.BeginUpdate()
        try {
            $list.Items.Clear(); $shown = 0; $hidden = 0
            foreach ($file in $state.AllFiles) {
                if (-not [IO.File]::Exists($file.FullName)) { continue }
                $rank = Get-RunnableRank $file $RootFolder
                $preference = Get-FilePreference $file.FullName
                $labelText = if ($preference.Label) { $preference.Label } else { $file.BaseName }
                $query = $search.Text.Trim()
                if ($query -and ($labelText + ' ' + $file.FullName).IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                if ($rank.Helper -and -not $showAll.Checked) { $hidden++; continue }
                $item = New-Object Windows.Forms.ListViewItem([string]$labelText); $item.Tag = $file.FullName; $item.ToolTipText = $file.FullName
                [void]$item.SubItems.Add($file.Name); [void]$item.SubItems.Add($rank.Reason); [void]$item.SubItems.Add($file.DirectoryName)
                if ($state.Icons.ContainsKey($file.FullName)) { $item.ImageIndex = $state.Icons[$file.FullName] }
                [void]$list.Items.Add($item); $shown++
                if ($file.FullName -eq $wanted) { $item.Selected = $true }
            }
            if (-not $list.SelectedItems.Count -and $list.Items.Count) { $list.Items[0].Selected = $true }
            $status.Text = "$shown shown  ·  $hidden helpers hidden  ·  F2 edits the shortcut name only"
            foreach ($name in @('Run', 'Shortcut', 'ShortcutLabel')) { $buttonMap[$name].Enabled = $list.Items.Count -gt 0 }
        } finally { $list.EndUpdate() }
    }.GetNewClosure()
    $search.Add_TextChanged($render); $showAll.Add_CheckedChanged($render)
    $list.Add_AfterLabelEdit({
        param($sender, $e)
        $e.CancelEdit = $true
        if ($null -eq $e.Label) { return }
        try {
            $labelValue = ConvertTo-ShortcutLabel $e.Label; $item = $list.Items[$e.Item]
            if (-not (Set-FilePreference -Path $item.Tag -Preference @{ Label = $labelValue })) { throw 'The shortcut name could not be saved.' }
            [void]$form.BeginInvoke([Action]({ $item.Text = $labelValue }.GetNewClosure()))
            & $addLog "Shortcut name: $labelValue (original file unchanged)"
        } catch { $state.LastError = $_.Exception.Message; Show-ThemedMessageBox -Message $state.LastError -Title 'Shortcut name' -Icon Warning | Out-Null }
    }.GetNewClosure())
    $editLabel = { if ($list.SelectedItems.Count) { $list.BeginShortcutNameEdit($list.SelectedItems[0]) } }.GetNewClosure()
    $list.Add_DoubleClick($editLabel); $buttonMap.ShortcutLabel.Add_Click($editLabel)
    $list.Add_KeyDown({ param($sender, $e) if ($e.KeyCode -eq 'F2') { & $editLabel; $e.Handled = $true; $e.SuppressKeyPress = $true } }.GetNewClosure())
    $context = New-Object Windows.Forms.ContextMenuStrip; $context.ShowImageMargin = $false; $context.ForeColor = $colors.TextFore
    $context.Renderer = New-Object Windows.Forms.ToolStripProfessionalRenderer((New-Object QBitLauncher.MenuColors($colors.ControlBack, $colors.ButtonBack, $colors.Border)))
    $labelMenu = $context.Items.Add('Edit shortcut name'); $labelMenu.Add_Click($editLabel); $list.ContextMenuStrip = $context
    $getSelected = { if (-not $list.SelectedItems.Count) { throw 'Select a file first.' }; return Get-Item -LiteralPath $list.SelectedItems[0].Tag -ErrorAction Stop }.GetNewClosure()
    $launch = {
        param($file, [bool]$admin)
        Start-Runnable -Path $file.FullName -Arguments $argumentsBox.Text -AsAdmin:$admin
    }.GetNewClosure()
    $buttonMap.Run.Add_Click({
        try { $file = & $getSelected; & $savePreference; & $launch $file $asAdmin.Checked; & $addLog "Launched: $($file.Name)"; Invoke-ActionSound Success }
        catch {
            $exception = $_.Exception; $cancelled = $false
            while ($exception) { if ($exception -is [ComponentModel.Win32Exception] -and $exception.NativeErrorCode -eq 1223) { $cancelled = $true }; $exception = $exception.InnerException }
            if ($cancelled -and $asAdmin.Checked) {
                $retry = Show-ThemedMessageBox -Message 'Administrator launch was cancelled. Run without administrator privileges?' -Title 'Launch cancelled' -Buttons YesNo -Icon Question
                if ($retry -eq [Windows.Forms.DialogResult]::Yes) { try { & $launch $file $false; & $addLog "Launched without admin: $($file.Name)"; return } catch { $state.LastError = $_.Exception.Message } } else { return }
            } else { $state.LastError = $_.Exception.Message }
            & $addLog ('Launch failed: ' + $state.LastError); Show-ThemedMessageBox -Message $state.LastError -Title 'Launch failed' -Icon Error | Out-Null
        }
    }.GetNewClosure())
    $buttonMap.Shortcut.Add_Click({
        try {
            $file = & $getSelected; & $savePreference
            $preference = Get-FilePreference $file.FullName; $name = if ($preference.Label) { $preference.Label } else { $file.BaseName }; $name = ConvertTo-ShortcutLabel $name
            $desktop = [Environment]::GetFolderPath('Desktop'); $linkPath = Join-Path $desktop ($name + '.lnk')
            if (Test-Path -LiteralPath $linkPath) {
                $collision = Show-ThemedMessageBox -Message "A desktop shortcut already exists:`n$name`n`nYes: replace it`nNo: keep both`nCancel: create nothing" -Title 'Shortcut already exists' -Buttons YesNoCancel -Icon Question
                if ($collision -eq [Windows.Forms.DialogResult]::Cancel) { return }
                if ($collision -eq [Windows.Forms.DialogResult]::No) { $linkPath = Get-AvailableShortcutPath -Directory $desktop -Label $name }
            }
            New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $file.FullName -WorkingDirectory $file.DirectoryName -Arguments $argumentsBox.Text
            & $addLog ('Desktop shortcut: ' + [IO.Path]::GetFileNameWithoutExtension($linkPath)); Invoke-ActionSound Success
        } catch { $state.LastError = $_.Exception.Message; & $addLog ('Shortcut failed: ' + $state.LastError); Show-ThemedMessageBox -Message $state.LastError -Title 'Shortcut failed' -Icon Error | Out-Null }
    }.GetNewClosure())
    $buttonMap.OpenFolder.Add_Click({ $path = if ($list.SelectedItems.Count) { [IO.Path]::GetDirectoryName([string]$list.SelectedItems[0].Tag) } else { $RootFolder }; if (Test-Path -LiteralPath $path) { Start-Process explorer.exe -ArgumentList ('"' + $path + '"') } }.GetNewClosure())
    $buttonMap.Cleanup.Add_Click({
        try {
            $selected = if ($list.SelectedItems.Count) { [string]$list.SelectedItems[0].Tag } else { '' }
            $choice = Show-CleanupConfirmForm -TorrentLabel $TorrentName -TorrentHash $TorrentHash -SourcePath $sourcePath -RootFolder $RootFolder -SelectedPath $selected
            if (-not $choice) { return }
            if ($choice.Permanent -and $choice.Plan -and $choice.Plan.Targets.Count) {
                $confirmation = Show-ThemedMessageBox -Message 'Permanently delete the selected files? This cannot be undone.' -Title 'Permanent deletion' -Buttons YesNo -Icon Warning
                if ($confirmation -ne [Windows.Forms.DialogResult]::Yes) { return }
            }
            [void](Invoke-CleanupSelection -Choice $choice -TorrentHash $TorrentHash -OnMessage $addLog)
            & $render
        } catch [OperationCanceledException] { & $addLog 'Cleanup cancelled' }
        catch { $state.LastError = $_.Exception.Message; & $addLog ('Cleanup failed: ' + $state.LastError); Show-ThemedMessageBox -Message $state.LastError -Title 'Cleanup' -Icon Error | Out-Null }
    }.GetNewClosure())
    $buttonMap.Settings.Add_Click({
        if (Show-SettingsForm) {
            Set-AppFormAppearance $form
            $current = $Global:CurrentTheme
            $moreMenu.ForeColor = $current.TextFore; $context.ForeColor = $current.TextFore
            $moreMenu.Renderer = New-Object Windows.Forms.ToolStripProfessionalRenderer((New-Object QBitLauncher.MenuColors($current.ControlBack, $current.ButtonBack, $current.Border)))
            $context.Renderer = New-Object Windows.Forms.ToolStripProfessionalRenderer((New-Object QBitLauncher.MenuColors($current.ControlBack, $current.ButtonBack, $current.Border)))
            Set-DestructiveButton $buttonMap.Cleanup
        }
    }.GetNewClosure())
    $buttonMap.OpenLog.Add_Click({ Write-LogMessage 'Log opened'; Start-Process -FilePath notepad.exe -ArgumentList ('"' + (Get-AppLogPath) + '"') }.GetNewClosure())
    $buttonMap.CopyError.Add_Click({ if ($state.LastError) { [Windows.Forms.Clipboard]::SetText((Protect-LogText $state.LastError)); & $addLog 'Error copied' } }.GetNewClosure())
    $buttonMap.Close.Add_Click({ $form.Close() }.GetNewClosure())
    $form.KeyPreview = $true
    $form.Add_KeyDown({ param($sender, $eventArgs)
        if ($eventArgs.Alt -and $eventArgs.KeyCode -eq 'T') {
            $onClick = [Windows.Forms.Button].GetMethod('OnClick', [Reflection.BindingFlags]'NonPublic,Instance')
            [void]$onClick.Invoke($buttonMap.Settings, @([EventArgs]::Empty)); $eventArgs.Handled = $true
        }
    }.GetNewClosure())
    $queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $iconTimer = New-Object Windows.Forms.Timer; $iconTimer.Interval = 60
    $iconTimer.Add_Tick({
        $entry = $null; $count = 0
        while ($count -lt 10 -and $queue.TryDequeue([ref]$entry)) {
            if ($entry.Bytes) {
                $stream = New-Object IO.MemoryStream(, [byte[]]$entry.Bytes)
                try { $bitmap = [Drawing.Bitmap]::FromStream($stream); try { $imageList.Images.Add($bitmap); $state.Icons[$entry.Path] = $imageList.Images.Count - 1 } finally { $bitmap.Dispose() } } finally { $stream.Dispose() }
                foreach ($item in $list.Items) { if ($item.Tag -eq $entry.Path) { $item.ImageIndex = $state.Icons[$entry.Path] } }
            }
            $entry = $null; $count++
        }
        if ($count) { $list.Invalidate() }
        if ($state.IconTask -and $state.IconTask.Handle.IsCompleted -and $queue.IsEmpty) { [void](Receive-AppWorker $state.IconTask); $state.IconTask = $null; $iconTimer.Stop() }
    }.GetNewClosure())
    $form.Add_Shown({
        Set-FormIcon $form; $list.FitColumns(); Set-DestructiveButton $buttonMap.Cleanup; & $render; $search.Focus() | Out-Null; $form.Tag = 'Ready'
        $cache = Join-Path (Get-AppDataDirectory) 'icons'
        $state.IconTask = Start-AppWorker -Action { param($files, $queue, $cache) Add-Type -AssemblyName System.Drawing; foreach ($file in $files) { $bytes = Get-RunnableIconBytes $file $cache; $queue.Enqueue([pscustomobject]@{ Path = $file.FullName; Bytes = $bytes }) } } -Arguments @($state.AllFiles, $queue, $cache) -Functions @('Get-RunnableIconBytes')
        $iconTimer.Start(); & $addLog ('Found ' + $state.AllFiles.Count + ' runnable files')
    }.GetNewClosure())
    Add-UpdatePolling -Form $form -OnAvailable { param($version) & $addLog "Update available: $version — open Settings to install it" }.GetNewClosure()
    try { [void]$form.ShowDialog() }
    finally {
        $state.Closing = $true; $preferenceTimer.Stop(); & $savePreference
        if ($form.WindowState -eq 'Normal') { $Global:UserSettings.WindowState = @{ Width = $form.Width; Height = $form.Height; Columns = @($list.Columns | ForEach-Object { $_.Width }) }; [void](Save-UserSettings) }
        $iconTimer.Stop(); $iconTimer.Dispose(); $preferenceTimer.Dispose(); Stop-AppWorker $state.IconTask
        $moreMenu.Dispose(); $context.Dispose(); $tip.Dispose(); $imageList.Dispose(); $form.Dispose(); $Global:MainForm = $null
    }
}

function Show-SettingsForm {
    Initialize-AppPresentation
    $form = New-Object Windows.Forms.Form; $form.Name = 'SettingsForm'; $form.Text = "Settings — qBitLauncher $Global:ScriptVersion"
    $form.Size = New-Object Drawing.Size(700, 570); $form.MinimumSize = New-Object Drawing.Size(560, 500); $form.StartPosition = 'CenterParent'; $form.AutoScaleMode = 'Dpi'
    $form.Font = New-Object Drawing.Font('Segoe UI', 10); $form.BackColor = $Global:CurrentTheme.FormBack; $form.ForeColor = $Global:CurrentTheme.TextFore
    $layout = New-Object Windows.Forms.TableLayoutPanel; $layout.Dock = 'Fill'; $layout.ColumnCount = 2; $layout.RowCount = 10; $layout.Padding = New-Object Windows.Forms.Padding(20)
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 170))); [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    for ($i = 0; $i -lt 7; $i++) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 40))) }
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100))); [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 42))); [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 42)))
    $theme = New-Object Windows.Forms.ComboBox; $theme.DropDownStyle = 'DropDownList'; $theme.FlatStyle = 'Flat'; [void]$theme.Items.AddRange(@('Dracula', 'Light')); $theme.SelectedItem = $Global:UserSettings.Theme
    $url = New-Object Windows.Forms.TextBox; $url.Text = $Global:UserSettings.QbittorrentWebUrl
    $username = New-Object Windows.Forms.TextBox; $username.Text = $Global:UserSettings.QbittorrentUsername
    $password = New-Object Windows.Forms.TextBox; $password.Text = $Global:UserSettings.QbittorrentPassword; $password.UseSystemPasswordChar = $true
    $destination = New-Object Windows.Forms.TextBox; $destination.Text = $Global:UserSettings.LastExtractionPath
    $defaultAdmin = New-Object QBitLauncher.ThemedCheckBox; $defaultAdmin.Text = 'Request administrator privileges by default'; $defaultAdmin.Checked = [bool]$Global:UserSettings.DefaultAdmin
    $autoUpdate = New-Object QBitLauncher.ThemedCheckBox; $autoUpdate.Text = 'Check for updates in the background'; $autoUpdate.Checked = [bool]$Global:UserSettings.CheckForUpdates
    $rows = @(@('Theme', $theme), @('qBittorrent API URL', $url), @('API username', $username), @('API password', $password), @('Extraction base folder', $destination), @('Launch preference', $defaultAdmin), @('Updates', $autoUpdate))
    $row = 0
    foreach ($spec in $rows) {
        $label = New-Object Windows.Forms.Label; $label.Text = $spec[0]; $label.Dock = 'Fill'
        $control = $spec[1]; $control.Dock = 'Fill'; $control.BackColor = $Global:CurrentTheme.ControlBack; $control.ForeColor = $Global:CurrentTheme.TextFore
        $presented = if ($control -is [Windows.Forms.TextBox]) { New-AppTextField $control } else { $control }
        $layout.Controls.Add($label, 0, $row); $layout.Controls.Add($presented, 1, $row); $row++
    }
    $info = New-Object Windows.Forms.Label; $info.Dock = 'Fill'; $info.Text = "Shortcut names, arguments, and admin mode are remembered per file.`r`nSaved passwords are encrypted for this Windows user.`r`nqBittorrent API repair preserves its existing username, password, HTTPS, and authentication settings."
    if ($Global:PasswordDecryptionFailed) { $info.Text += "`r`nThe saved password could not be decrypted. Enter it again to replace it." }
    $layout.Controls.Add($info, 0, 7); $layout.SetColumnSpan($info, 2)
    $update = New-Object Windows.Forms.Button; $update.Text = 'Check for updates'; $update.Dock = 'Fill'; Set-ThemedButton $update
    $updateStatus = New-Object Windows.Forms.Label; $updateStatus.Dock = 'Fill'
    $layout.Controls.Add($update, 0, 8); $layout.Controls.Add($updateStatus, 1, 8)
    $actions = New-Object Windows.Forms.FlowLayoutPanel; $actions.Dock = 'Fill'; $actions.FlowDirection = 'RightToLeft'
    $save = New-Object Windows.Forms.Button; $save.Text = 'Save'; $save.AutoSize = $true; Set-ThemedButton $save
    $cancel = New-Object Windows.Forms.Button; $cancel.Text = 'Cancel'; $cancel.AutoSize = $true; Set-ThemedButton $cancel
    $actions.Controls.AddRange(@($save, $cancel)); $layout.Controls.Add($actions, 0, 9); $layout.SetColumnSpan($actions, 2)
    $state = @{ Saved = $false; Check = $null }; $timer = New-Object Windows.Forms.Timer; $timer.Interval = 150
    $update.Add_Click({
        if ($Global:AvailableVersion) { [void](Show-UpdatePrompt -NewVersion $Global:AvailableVersion); return }
        $state.Check = Start-BackgroundUpdateCheck -Force; $update.Enabled = $false; $updateStatus.Text = 'Checking...'; $timer.Start()
        if (-not $state.Check) { $updateStatus.Text = 'Checks are disabled for this diagnostic session.'; $update.Enabled = $true; $timer.Stop() }
    }.GetNewClosure())
    $timer.Add_Tick({
        if (-not $state.Check) { return }
        $result = Receive-BackgroundUpdateCheck $state.Check
        if ($result) {
            $timer.Stop(); $state.Check = $null; $update.Enabled = $true
            if ($result.Error) { $updateStatus.Text = 'Could not check. Open Log for details.'; Write-LogMessage $result.Error }
            elseif ($result.Version) { $Global:AvailableVersion = $result.Version; $updateStatus.Text = "Version $($result.Version) available"; $update.Text = 'Install update' }
            else { $updateStatus.Text = 'You are up to date.' }
        }
    }.GetNewClosure())
    $save.Add_Click({
        $parsed = $null
        if (-not [Uri]::TryCreate($url.Text.Trim(), [UriKind]::Absolute, [ref]$parsed) -or $parsed.Scheme -notin @('http', 'https') -or $parsed.UserInfo) {
            Show-ThemedMessageBox -Message 'Enter an HTTP or HTTPS API URL without an embedded username/password.' -Title 'API URL' -Icon Warning | Out-Null; return
        }
        if ($destination.Text.Trim() -and -not [IO.Path]::IsPathRooted($destination.Text.Trim())) { Show-ThemedMessageBox -Message 'Use an absolute extraction base folder, or leave it blank.' -Title 'Extraction folder' -Icon Warning | Out-Null; return }
        $Global:UserSettings.Theme = [string]$theme.SelectedItem; $Global:UserSettings.QbittorrentWebUrl = $url.Text.Trim(); $Global:UserSettings.QbittorrentUsername = $username.Text.Trim(); $Global:UserSettings.QbittorrentPassword = $password.Text
        $Global:UserSettings.LastExtractionPath = $destination.Text.Trim(); $Global:UserSettings.DefaultAdmin = $defaultAdmin.Checked; $Global:UserSettings.CheckForUpdates = $autoUpdate.Checked
        if (-not (Save-UserSettings)) { Show-ThemedMessageBox -Message 'Settings could not be saved. Check the log.' -Title 'Settings' -Icon Error | Out-Null; return }
        $Global:ThemeSelection = $Global:UserSettings.Theme; $Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]; $state.Saved = $true; $form.Close()
    }.GetNewClosure())
    $cancel.Add_Click({ $form.Close() }.GetNewClosure()); $form.Controls.Add($layout); $form.AcceptButton = $save; $form.CancelButton = $cancel
    $form.Add_Shown({ Set-FormIcon $form }.GetNewClosure())
    try { if ($Global:MainForm) { [void]$form.ShowDialog($Global:MainForm) } else { [void]$form.ShowDialog() }; return $state.Saved }
    finally { $timer.Stop(); $timer.Dispose(); if ($state.Check) { Stop-BackgroundUpdateCheck $state.Check }; $form.Dispose() }
}

function Show-SetupMenu {
    $form = New-Object Windows.Forms.Form; $form.Name = 'DashboardForm'; $form.Text = "qBitLauncher $Global:ScriptVersion"
    $form.Size = New-Object Drawing.Size(620, 650); $form.MinimumSize = New-Object Drawing.Size(500, 540); $form.StartPosition = 'CenterScreen'; $form.AutoScaleMode = 'Dpi'
    $form.BackColor = $Global:CurrentTheme.FormBack; $form.ForeColor = $Global:CurrentTheme.TextFore; $form.Font = New-Object Drawing.Font('Segoe UI', 10)
    $layout = New-Object Windows.Forms.TableLayoutPanel; $layout.Dock = 'Fill'; $layout.Padding = New-Object Windows.Forms.Padding(22); $layout.ColumnCount = 1; $layout.RowCount = 10
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    $title = New-Object Windows.Forms.Label; $title.Text = 'qBitLauncher'; $title.Font = New-Object Drawing.Font('Segoe UI', 20, [Drawing.FontStyle]::Bold); $title.Dock = 'Fill'
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 54))); [void]$layout.Controls.Add($title)
    $desc = New-Object Windows.Forms.Label; $desc.Text = 'Open a folder, manage integrations, or return to a recent download.'; $desc.Dock = 'Fill'
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 48))); [void]$layout.Controls.Add($desc)
    $recent = New-Object Windows.Forms.ComboBox; $recent.Name = 'RecentFolders'; $recent.DropDownStyle = 'DropDownList'; $recent.FlatStyle = 'Flat'; $recent.Dock = 'Fill'
    $refreshRecent = { $recent.Items.Clear(); foreach ($path in @($Global:UserSettings.RecentFolders)) { [void]$recent.Items.Add([string]$path) }; if ($recent.Items.Count) { $recent.SelectedIndex = 0 } }.GetNewClosure()
    & $refreshRecent
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 38))); [void]$layout.Controls.Add($recent)
    $state = @{ Result = [Windows.Forms.DialogResult]::Cancel }
    $specs = @(@('Open recent folder', 'Recent'), @('Choose a folder...', 'Browse'), @('Forget selected recent folder', 'Forget'), @('Install / update qBittorrent integration', 'InstallQB'), @('Add Shift+right-click context menus', 'InstallContext'), @('Remove integrations...', 'Remove'), @('Settings', 'Settings'))
    foreach ($spec in $specs) {
        $button = New-Object Windows.Forms.Button; $button.Text = $spec[0]; $button.Tag = $spec[1]; $button.Dock = 'Fill'; Set-ThemedButton $button
        [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', (100.0 / 7)))); [void]$layout.Controls.Add($button)
        $button.Add_Click({
            switch ($this.Tag) {
                'Recent' {
                    if ($recent.SelectedIndex -lt 0) { return }
                    $path = [string]$recent.SelectedItem
                    if (-not (Test-Path -LiteralPath $path -PathType Container)) { Show-ThemedMessageBox -Message 'This folder has moved or was deleted. Choose it again, or forget this entry.' -Title 'Folder missing' -Icon Warning | Out-Null; return }
                    $Global:PendingManualFolder = $path; $state.Result = [Windows.Forms.DialogResult]::OK; $form.Close()
                }
                'Browse' { $Global:PendingManualFolder = $null; $state.Result = [Windows.Forms.DialogResult]::OK; $form.Close() }
                'Forget' { if ($recent.SelectedIndex -ge 0) { [void](Remove-RecentFolder ([string]$recent.SelectedItem)); & $refreshRecent } }
                'InstallQB' { [void](Install-QBittorrentIntegration) }
                'InstallContext' { [void](Install-ContextMenu) }
                'Remove' { $choice = Show-ThemedMessageBox -Message "Yes: remove qBittorrent integration`nNo: remove Explorer context menus" -Title 'Remove integrations' -Buttons YesNoCancel -Icon Question; if ($choice -eq [Windows.Forms.DialogResult]::Yes) { [void](Remove-QBittorrentIntegration) } elseif ($choice -eq [Windows.Forms.DialogResult]::No) { [void](Remove-ContextMenu) } }
                'Settings' { if (Show-SettingsForm) { $state.Result = [Windows.Forms.DialogResult]::Retry; $form.Close() } }
            }
        }.GetNewClosure())
    }
    $form.Controls.Add($layout); $Global:MainForm = $form
    Add-UpdatePolling -Form $form -OnAvailable { param($version) $desc.Text = "Update $version available — open Settings to install it." }.GetNewClosure()
    $form.Add_Shown({ Set-FormIcon $form }.GetNewClosure())
    try { [void]$form.ShowDialog(); return $state.Result } finally { $form.Dispose(); $Global:MainForm = $null }
}
