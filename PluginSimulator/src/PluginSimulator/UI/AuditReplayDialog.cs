using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using PluginSimulator.Execution;
using PluginSimulator.Models;

namespace PluginSimulator.UI
{
    /// <summary>
    /// Audit-driven Pre/Post image reconstruction popup.
    /// User picks an audit entry → seeds the deltas grid → edits Before/After values →
    /// live preview shows Target / PreImage / PostImage that will be handed to the plugin.
    /// </summary>
    public class AuditReplayDialog : Form
    {
        private readonly List<AuditEntry> _auditEntries;
        private readonly Entity _currentRecord;
        private readonly string _entityLogicalName;
        private readonly Guid _recordId;
        private readonly string _messageName;
        private readonly int _stage;

        private DataGridView _gridAudit;
        private DataGridView _gridDeltas;
        private TextBox _txtPreviewTarget;
        private TextBox _txtPreviewPreImage;
        private TextBox _txtPreviewPostImage;
        private Button _btnAddManual, _btnRemoveDelta, _btnLoadAudit, _btnApply, _btnCancel;
        private System.Windows.Forms.Label _lblWarning;

        public List<FieldDelta> ConfiguredDeltas { get; private set; } = new List<FieldDelta>();

        public AuditReplayDialog(
            List<AuditEntry> auditEntries,
            Entity currentRecord,
            string entityLogicalName,
            Guid recordId,
            string messageName,
            int stage)
        {
            _auditEntries = auditEntries ?? new List<AuditEntry>();
            _currentRecord = currentRecord;
            _entityLogicalName = entityLogicalName;
            _recordId = recordId;
            _messageName = messageName ?? "Update";
            _stage = stage;

            BuildUI();
            PopulateAuditGrid();
            RebuildPreview();
        }

        private void BuildUI()
        {
            Text = "Audit-Driven Image Reconstruction";
            Width = 1180;
            Height = 800;
            MinimumSize = new Size(900, 650);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9F);

            // ─── Warning banner ─────────────────────────────────────────────
            _lblWarning = new System.Windows.Forms.Label
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = Color.FromArgb(255, 248, 220),
                ForeColor = Color.FromArgb(120, 80, 0),
                Padding = new Padding(12, 8, 12, 8),
                Font = new Font("Segoe UI", 9F),
                Text =
                    "⚠  NEAR-CORRECT RECONSTRUCTION  —  This rebuilds the plugin context from audit history " +
                    "and your edits, not from a live plugin capture.\n" +
                    "Audit-disabled fields, fields written with same-value updates, and changes outside the " +
                    "audit retention window will NOT appear. Treat the result as a faithful approximation, not ground truth."
            };
            Controls.Add(_lblWarning);

            // ─── Bottom buttons ─────────────────────────────────────────────
            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            _btnApply = new Button
            {
                Text = "▶  Apply",
                Left = 12, Top = 10, Width = 130, Height = 32,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            _btnApply.FlatAppearance.BorderSize = 0;
            _btnApply.Click += (s, e) => ApplyAndClose();

            _btnCancel = new Button
            {
                Text = "Cancel",
                Left = 150, Top = 10, Width = 90, Height = 32,
                FlatStyle = FlatStyle.Flat
            };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            btnPanel.Controls.Add(_btnApply);
            btnPanel.Controls.Add(_btnCancel);
            Controls.Add(btnPanel);

            // ─── Main split: top half = audit picker + deltas grid, bottom half = preview ───
            var rootSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 380
            };
            Controls.Add(rootSplit);

            // ─── TOP PANEL: audit picker (left) + deltas (right) ────────────
            var topSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 460
            };
            rootSplit.Panel1.Controls.Add(topSplit);

            // Step 1 — Audit picker
            var auditPanel = new Panel { Dock = DockStyle.Fill };
            var lblStep1 = new System.Windows.Forms.Label
            {
                Text = "Step 1 — Pick an audit entry to seed the deltas (optional)",
                Dock = DockStyle.Top, Height = 24, Padding = new Padding(8, 4, 0, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 180)
            };
            _gridAudit = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                Font = new Font("Consolas", 8.5F)
            };
            _gridAudit.DoubleClick += (s, e) => LoadAuditIntoDeltas();
            var auditBtnPanel = new Panel { Dock = DockStyle.Bottom, Height = 36 };
            _btnLoadAudit = new Button
            {
                Text = "↓  Seed deltas from selected audit entry",
                Left = 6, Top = 4, Width = 280, Height = 28,
                FlatStyle = FlatStyle.Flat
            };
            _btnLoadAudit.Click += (s, e) => LoadAuditIntoDeltas();
            auditBtnPanel.Controls.Add(_btnLoadAudit);
            auditPanel.Controls.Add(_gridAudit);
            auditPanel.Controls.Add(auditBtnPanel);
            auditPanel.Controls.Add(lblStep1);
            topSplit.Panel1.Controls.Add(auditPanel);

            // Step 2 — Deltas grid
            var deltaPanel = new Panel { Dock = DockStyle.Fill };
            var lblStep2 = new System.Windows.Forms.Label
            {
                Text = "Step 2 — Edit field deltas (Before / After). These drive Target & images.",
                Dock = DockStyle.Top, Height = 24, Padding = new Padding(8, 4, 0, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 180)
            };
            _gridDeltas = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                Font = new Font("Consolas", 8.5F),
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            _gridDeltas.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Field", FillWeight = 22, Name = "colField" });
            var colType = new DataGridViewComboBoxColumn { HeaderText = "Type", FillWeight = 13, Name = "colType" };
            colType.Items.AddRange(AttributeTypes.All);
            _gridDeltas.Columns.Add(colType);
            _gridDeltas.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Before", FillWeight = 22, Name = "colBefore" });
            _gridDeltas.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "After",  FillWeight = 22, Name = "colAfter" });
            _gridDeltas.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "EntityRef Target", FillWeight = 13, Name = "colEr" });
            _gridDeltas.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Source", FillWeight = 8, Name = "colSource", ReadOnly = true });
            _gridDeltas.CellValueChanged += (s, e) => RebuildPreview();
            _gridDeltas.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_gridDeltas.IsCurrentCellDirty) _gridDeltas.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _gridDeltas.DataError += (s, e) => { e.Cancel = true; e.ThrowException = false; };

            var deltaBtnPanel = new Panel { Dock = DockStyle.Bottom, Height = 36 };
            _btnAddManual = new Button
            {
                Text = "+  Add manual field",
                Left = 6, Top = 4, Width = 140, Height = 28,
                FlatStyle = FlatStyle.Flat
            };
            _btnAddManual.Click += (s, e) => AddManualRow();
            _btnRemoveDelta = new Button
            {
                Text = "−  Remove selected",
                Left = 152, Top = 4, Width = 140, Height = 28,
                FlatStyle = FlatStyle.Flat
            };
            _btnRemoveDelta.Click += (s, e) => RemoveSelectedRow();
            deltaBtnPanel.Controls.Add(_btnAddManual);
            deltaBtnPanel.Controls.Add(_btnRemoveDelta);

            deltaPanel.Controls.Add(_gridDeltas);
            deltaPanel.Controls.Add(deltaBtnPanel);
            deltaPanel.Controls.Add(lblStep2);
            topSplit.Panel2.Controls.Add(deltaPanel);

            // ─── BOTTOM PANEL: live preview (3 read-only text panes) ────────
            var lblStep3 = new System.Windows.Forms.Label
            {
                Text = $"Step 3 — Preview ({_messageName} @ Stage {_stage})  —  what the plugin will see",
                Dock = DockStyle.Top, Height = 24, Padding = new Padding(8, 4, 0, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 180)
            };
            var previewSplit = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };
            previewSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            previewSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            previewSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34F));

            _txtPreviewTarget    = BuildPreviewPane("Target");
            _txtPreviewPreImage  = BuildPreviewPane("PreImage");
            _txtPreviewPostImage = BuildPreviewPane("PostImage");

            previewSplit.Controls.Add(WrapWithHeader(_txtPreviewTarget,    "🎯  Target"),    0, 0);
            previewSplit.Controls.Add(WrapWithHeader(_txtPreviewPreImage,  "📥  PreImage"),  1, 0);
            previewSplit.Controls.Add(WrapWithHeader(_txtPreviewPostImage, "📤  PostImage"), 2, 0);

            rootSplit.Panel2.Controls.Add(previewSplit);
            rootSplit.Panel2.Controls.Add(lblStep3);
        }

        private static TextBox BuildPreviewPane(string name) => new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Dock = DockStyle.Fill,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font("Consolas", 8.5F),
            BackColor = Color.FromArgb(248, 248, 248)
        };

        private static Panel WrapWithHeader(Control inner, string headerText)
        {
            var p = new Panel { Dock = DockStyle.Fill, Padding = new Padding(2) };
            var lbl = new System.Windows.Forms.Label
            {
                Text = headerText,
                Dock = DockStyle.Top,
                Height = 22,
                Padding = new Padding(4, 4, 0, 0),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            inner.Dock = DockStyle.Fill;
            p.Controls.Add(inner);
            p.Controls.Add(lbl);
            return p;
        }

        private void PopulateAuditGrid()
        {
            _gridAudit.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#",          FillWeight = 5  });
            _gridAudit.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Changed On", FillWeight = 22 });
            _gridAudit.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Changed By", FillWeight = 25 });
            _gridAudit.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Fields",     FillWeight = 48 });

            for (int i = 0; i < _auditEntries.Count; i++)
            {
                var e = _auditEntries[i];
                _gridAudit.Rows.Add(
                    i + 1,
                    e.ChangedOn == DateTime.MinValue ? "Unknown" : e.ChangedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    e.ChangedBy,
                    e.ChangedFieldsSummary);
            }
        }

        private void LoadAuditIntoDeltas()
        {
            if (_gridAudit.SelectedRows.Count == 0)
            {
                MessageBox.Show(this, "Select an audit row first.", "No selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var idx = _gridAudit.SelectedRows[0].Index;
            if (idx < 0 || idx >= _auditEntries.Count) return;
            var entry = _auditEntries[idx];

            _gridDeltas.Rows.Clear();

            var fieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (entry.NewValue != null)
                foreach (var k in entry.NewValue.Attributes.Keys) fieldNames.Add(k);
            if (entry.OldValue != null)
                foreach (var k in entry.OldValue.Attributes.Keys) fieldNames.Add(k);

            foreach (var field in fieldNames)
            {
                object oldV = entry.OldValue != null && entry.OldValue.Attributes.ContainsKey(field) ? entry.OldValue[field] : null;
                object newV = entry.NewValue != null && entry.NewValue.Attributes.ContainsKey(field) ? entry.NewValue[field] : null;
                string type = AttributeTypes.DetectFrom(newV ?? oldV);
                string entRef = (newV as EntityReference)?.LogicalName
                                ?? (oldV as EntityReference)?.LogicalName
                                ?? "";
                _gridDeltas.Rows.Add(
                    field,
                    type,
                    ContextReconstructor.FormatValue(oldV),
                    ContextReconstructor.FormatValue(newV),
                    entRef,
                    "Audit");
            }
            RebuildPreview();
        }

        private void AddManualRow()
        {
            _gridDeltas.Rows.Add("", "String", "", "", "", "Manual");
            RebuildPreview();
        }

        private void RemoveSelectedRow()
        {
            if (_gridDeltas.SelectedRows.Count == 0) return;
            foreach (DataGridViewRow row in _gridDeltas.SelectedRows)
                if (!row.IsNewRow) _gridDeltas.Rows.Remove(row);
            RebuildPreview();
        }

        private List<FieldDelta> CollectDeltas()
        {
            var deltas = new List<FieldDelta>();
            foreach (DataGridViewRow row in _gridDeltas.Rows)
            {
                if (row.IsNewRow) continue;
                var field = row.Cells["colField"].Value?.ToString();
                if (string.IsNullOrWhiteSpace(field)) continue;

                deltas.Add(new FieldDelta
                {
                    FieldName = field.Trim(),
                    AttributeType = row.Cells["colType"].Value?.ToString() ?? "String",
                    BeforeRaw = row.Cells["colBefore"].Value?.ToString(),
                    AfterRaw = row.Cells["colAfter"].Value?.ToString(),
                    EntityRefLogicalName = row.Cells["colEr"].Value?.ToString(),
                    Source = ParseSource(row.Cells["colSource"].Value?.ToString())
                });
            }
            return deltas;
        }

        private static DeltaSource ParseSource(string s)
        {
            if (Enum.TryParse<DeltaSource>(s, true, out var v)) return v;
            return DeltaSource.Manual;
        }

        private void RebuildPreview()
        {
            try
            {
                var deltas = CollectDeltas();
                var rebuilt = ContextReconstructor.Reconstruct(
                    _entityLogicalName, _recordId, _currentRecord, deltas, _messageName, _stage);

                _txtPreviewTarget.Text = FormatEntityPane(rebuilt.Target, deltas, includeAllAttrs: true);
                _txtPreviewPreImage.Text = rebuilt.PreImage == null
                    ? "(not applicable for Create)"
                    : FormatPreImagePane(rebuilt.PreImage, rebuilt.PreImageProvenance);
                _txtPreviewPostImage.Text = rebuilt.PostImage == null
                    ? $"(not applicable — PostImage requires Stage 40 + Create/Update)"
                    : FormatEntityPane(rebuilt.PostImage, deltas, includeAllAttrs: true);
            }
            catch (Exception ex)
            {
                _txtPreviewTarget.Text = $"⚠ Preview error:\r\n{ex.Message}";
                _txtPreviewPreImage.Text = "";
                _txtPreviewPostImage.Text = "";
            }
        }

        private string FormatEntityPane(Entity e, List<FieldDelta> deltas, bool includeAllAttrs)
        {
            if (e == null) return "(null)";
            var sb = new StringBuilder();
            sb.AppendLine($"[{e.LogicalName}] {e.Id}");
            sb.AppendLine($"Attributes: {e.Attributes.Count}");
            sb.AppendLine(new string('─', 40));
            foreach (var attr in e.Attributes.OrderBy(a => a.Key))
            {
                sb.AppendLine($"{attr.Key,-30} = {ContextReconstructor.FormatValue(attr.Value)}");
            }
            return sb.ToString();
        }

        private string FormatPreImagePane(Entity preImage, Dictionary<string, DeltaSource> provenance)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[{preImage.LogicalName}] {preImage.Id}");
            sb.AppendLine($"Attributes: {preImage.Attributes.Count}    📋=Audit  ✏=Manual  📥=CurrentRetrieve");
            sb.AppendLine(new string('─', 40));
            foreach (var attr in preImage.Attributes.OrderBy(a => a.Key))
            {
                string icon = "📥";
                if (provenance.TryGetValue(attr.Key, out var src))
                {
                    switch (src)
                    {
                        case DeltaSource.Audit:           icon = "📋"; break;
                        case DeltaSource.Manual:          icon = "✏"; break;
                        case DeltaSource.CurrentRetrieve: icon = "📥"; break;
                    }
                }
                sb.AppendLine($"{icon} {attr.Key,-28} = {ContextReconstructor.FormatValue(attr.Value)}");
            }
            return sb.ToString();
        }

        private void ApplyAndClose()
        {
            try
            {
                var deltas = CollectDeltas();
                // Validate by parsing — surface bad values now, not at runtime.
                foreach (var d in deltas)
                {
                    ContextReconstructor.ParseValue(d.BeforeRaw, d.AttributeType, d.EntityRefLogicalName);
                    ContextReconstructor.ParseValue(d.AfterRaw, d.AttributeType, d.EntityRefLogicalName);
                }
                ConfiguredDeltas = deltas;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "One or more values could not be parsed with the chosen type:\n\n" + ex.Message,
                    "Validation failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
