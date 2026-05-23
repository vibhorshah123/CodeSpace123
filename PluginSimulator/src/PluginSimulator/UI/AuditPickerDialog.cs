using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PluginSimulator.Execution;

namespace PluginSimulator.UI
{
    /// <summary>
    /// Shows recent audit history entries for a record and lets the user pick one to replay.
    /// </summary>
    public class AuditPickerDialog : Form
    {
        private DataGridView _grid;
        private Button _btnUse, _btnCancel;
        private Label _lblHint;
        private readonly List<AuditEntry> _entries;

        public AuditEntry SelectedEntry { get; private set; }

        public AuditPickerDialog(List<AuditEntry> entries)
        {
            _entries = entries;
            BuildUI();
            PopulateGrid();
        }

        private void BuildUI()
        {
            Text = "Select Audit Entry to Replay";
            Width = 820;
            Height = 420;
            MinimumSize = new Size(600, 300);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9F);

            _lblHint = new Label
            {
                Text = "Double-click a row or select and click Use This Entry. " +
                       "Pre-Image (old values) and Target (changed fields) will be loaded automatically.",
                Dock = DockStyle.Top,
                Height = 36,
                Padding = new Padding(8, 8, 8, 0),
                ForeColor = Color.DimGray
            };

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                Font = new Font("Consolas", 9F)
            };
            _grid.DoubleClick += (s, e) => UseSelected();

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 45 };

            _btnUse = new Button
            {
                Text = "▶  Use This Entry",
                Left = 10, Top = 8, Width = 160, Height = 30,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            _btnUse.FlatAppearance.BorderSize = 0;
            _btnUse.Click += (s, e) => UseSelected();

            _btnCancel = new Button
            {
                Text = "Cancel",
                Left = 180, Top = 8, Width = 80, Height = 30,
                FlatStyle = FlatStyle.Flat
            };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            btnPanel.Controls.Add(_btnUse);
            btnPanel.Controls.Add(_btnCancel);

            Controls.Add(_grid);
            Controls.Add(btnPanel);
            Controls.Add(_lblHint);
        }

        private void PopulateGrid()
        {
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#",            Width = 40,  FillWeight = 5 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Changed On",   Width = 160, FillWeight = 20 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Changed By",   Width = 180, FillWeight = 25 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Fields Changed",            FillWeight = 50 });

            for (int i = 0; i < _entries.Count; i++)
            {
                var e = _entries[i];
                _grid.Rows.Add(
                    i + 1,
                    e.ChangedOn == DateTime.MinValue ? "Unknown" : e.ChangedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    e.ChangedBy,
                    e.ChangedFieldsSummary);
            }
        }

        private void UseSelected()
        {
            if (_grid.SelectedRows.Count == 0) return;
            var idx = _grid.SelectedRows[0].Index;
            if (idx < 0 || idx >= _entries.Count) return;
            SelectedEntry = _entries[idx];
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
