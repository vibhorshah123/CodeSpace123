using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using PluginSimulator.Authentication;
using PluginSimulator.Execution;
using PluginSimulator.Models;
using PluginSimulator.Reporting;

namespace PluginSimulator.UI
{
    public class MainForm : Form
    {
        private TextBox txtEnvUrl, txtEntity, txtRecordId, txtAssemblyPath, txtPluginType, txtChangedFields, txtPreImageName;
        private ComboBox cboMessage, cboStage;
        private CheckBox chkDebug;
        private Button btnBrowse, btnRun, btnClear, btnAudit;
        private RichTextBox rtbOutput;
        private System.Windows.Forms.Label lblStatus, lblAuditStatus;
        private bool _isRunning;
        private AuditEntry _loadedAuditEntry;
        private IOrganizationService _cachedService;
        private string _cachedEnvUrl;

        public MainForm()
        {
            BuildUI();
            LoadDefaults();
        }

        private void BuildUI()
        {
            Text = "D365 Plugin Simulator v1.0";
            Width = 950;
            Height = 750;
            MinimumSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            // Top panel - config fields
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 295, Padding = new Padding(10) };

            int y = 10;
            AddRow(topPanel, "Environment URL:", out txtEnvUrl, 10, y, 400);
            AddRow(topPanel, "Entity:", out txtEntity, 480, y, 400);
            y += 30;
            AddRow(topPanel, "Record ID (GUID):", out txtRecordId, 10, y, 400);
            AddComboRow(topPanel, "Message:", out cboMessage, 480, y, 180,
                new[] { "Create", "Update", "Delete" }, 0);
            y += 30;
            AddRow(topPanel, "Assembly (DLL):", out txtAssemblyPath, 10, y, 370);
            btnBrowse = new Button { Text = "...", Left = 505, Top = y, Width = 30, Height = 23 };
            btnBrowse.Click += BtnBrowse_Click;
            topPanel.Controls.Add(btnBrowse);
            AddComboRow(topPanel, "Stage:", out cboStage, 550, y, 180,
                new[] { "10 - Pre-Validation", "20 - Pre-Operation", "40 - Post-Operation" }, 2);
            y += 30;
            AddRow(topPanel, "Plugin Type (optional):", out txtPluginType, 10, y, 400);
            AddRow(topPanel, "Changed Fields:", out txtChangedFields, 480, y, 400);
            y += 30;
            AddRow(topPanel, "Pre-Image Name:", out txtPreImageName, 10, y, 300);
            txtPreImageName.Text = "PreImage";

            btnAudit = new Button
            {
                Text = "📋 Load from Audit",
                Left = 360, Top = y, Width = 150, Height = 23,
                FlatStyle = FlatStyle.Flat
            };
            btnAudit.Click += BtnAudit_Click;
            topPanel.Controls.Add(btnAudit);

            y += 30;
            lblAuditStatus = new System.Windows.Forms.Label
            {
                Text = "No audit loaded — pre-image uses current D365 record values",
                Left = 10, Top = y, AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic)
            };
            topPanel.Controls.Add(lblAuditStatus);
            y += 30;

            // Buttons row
            btnRun = new Button
            {
                Text = "▶  Run Simulation",
                Left = 10, Top = y, Width = 160, Height = 35,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnRun.FlatAppearance.BorderSize = 0;
            btnRun.Click += BtnRun_Click;
            topPanel.Controls.Add(btnRun);

            btnClear = new Button
            {
                Text = "Clear Log",
                Left = 180, Top = y, Width = 90, Height = 35,
                FlatStyle = FlatStyle.Flat
            };
            btnClear.Click += (s, e) => rtbOutput.Clear();
            topPanel.Controls.Add(btnClear);

            chkDebug = new CheckBox
            {
                Text = "Debug (attaches VS before plugin executes)",
                Left = 290, Top = y + 8, AutoSize = true
            };
            topPanel.Controls.Add(chkDebug);

            lblStatus = new System.Windows.Forms.Label
            {
                Text = "Ready",
                Left = 10, Top = y + 40, AutoSize = true, ForeColor = Color.Gray
            };
            topPanel.Controls.Add(lblStatus);

            // Output panel
            rtbOutput = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(220, 220, 220),
                Font = new Font("Consolas", 9F),
                WordWrap = false
            };

            Controls.Add(rtbOutput);
            Controls.Add(topPanel);
        }

        private void AddRow(Panel p, string label, out TextBox txt, int x, int y, int width)
        {
            var lbl = new System.Windows.Forms.Label { Text = label, Left = x, Top = y + 3, AutoSize = true };
            txt = new TextBox { Left = x + 140, Top = y, Width = width - 140 };
            p.Controls.Add(lbl);
            p.Controls.Add(txt);
        }

        private void AddComboRow(Panel p, string label, out ComboBox cbo, int x, int y, int width, string[] items, int selectedIndex)
        {
            var lbl = new System.Windows.Forms.Label { Text = label, Left = x, Top = y + 3, AutoSize = true };
            cbo = new ComboBox
            {
                Left = x + 70, Top = y, Width = width - 70,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cbo.Items.AddRange(items);
            cbo.SelectedIndex = selectedIndex;
            p.Controls.Add(lbl);
            p.Controls.Add(cbo);
        }

        private void LoadDefaults()
        {
            // All fields empty — user fills in for any plugin
            // Last-used values are saved/loaded from config file
            var configPath = Path.Combine(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                "PluginSimulator.settings.json");

            if (File.Exists(configPath))
            {
                try
                {
                    var json = File.ReadAllText(configPath);
                    var settings = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                    if (settings.ContainsKey("EnvUrl")) txtEnvUrl.Text = settings["EnvUrl"];
                    if (settings.ContainsKey("Entity")) txtEntity.Text = settings["Entity"];
                    if (settings.ContainsKey("AssemblyPath")) txtAssemblyPath.Text = settings["AssemblyPath"];
                    if (settings.ContainsKey("PluginType")) txtPluginType.Text = settings["PluginType"];
                    if (settings.ContainsKey("RecordId")) txtRecordId.Text = settings["RecordId"];
                    if (settings.ContainsKey("Message")) cboMessage.SelectedItem = settings["Message"];
                    if (settings.ContainsKey("PreImageName")) txtPreImageName.Text = settings["PreImageName"];
                }
                catch { /* ignore corrupt settings */ }
            }
        }

        private void SaveSettings()
        {
            var settings = new Dictionary<string, string>
            {
                ["EnvUrl"] = txtEnvUrl.Text,
                ["Entity"] = txtEntity.Text,
                ["AssemblyPath"] = txtAssemblyPath.Text,
                ["PluginType"] = txtPluginType.Text,
                ["RecordId"] = txtRecordId.Text,
                ["Message"] = cboMessage.SelectedItem?.ToString() ?? "Create",
                ["PreImageName"] = txtPreImageName.Text
            };

            var configPath = Path.Combine(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                "PluginSimulator.settings.json");

            try
            {
                File.WriteAllText(configPath, Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented));
            }
            catch { /* ignore */ }
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select Plugin Assembly";
                dlg.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*";
                if (!string.IsNullOrEmpty(txtAssemblyPath.Text))
                    dlg.InitialDirectory = Path.GetDirectoryName(txtAssemblyPath.Text);

                if (dlg.ShowDialog() == DialogResult.OK)
                    txtAssemblyPath.Text = dlg.FileName;
            }
        }

        private async void BtnAudit_Click(object sender, EventArgs e)
        {
            var envUrl  = txtEnvUrl.Text.Trim();
            var entity  = txtEntity.Text.Trim();
            var recordId = txtRecordId.Text.Trim();

            if (string.IsNullOrWhiteSpace(envUrl) || string.IsNullOrWhiteSpace(entity) || string.IsNullOrWhiteSpace(recordId))
            {
                MessageBox.Show("Fill in Environment URL, Entity, and Record ID first.", "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!Guid.TryParse(recordId, out var guid))
            {
                MessageBox.Show("Record ID must be a valid GUID.", "Invalid Record ID", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnAudit.Enabled = false;
            btnAudit.Text = "⏳ Loading...";
            lblAuditStatus.Text = "Connecting to D365...";

            List<AuditEntry> entries = null;
            string errorMsg = null;

            try
            {
                await Task.Run(() =>
                {
                    // Reuse cached connection if URL unchanged
                    if (_cachedService == null || _cachedEnvUrl != envUrl)
                    {
                        _cachedService = AuthManager.Connect(envUrl);
                        _cachedEnvUrl  = envUrl;
                    }
                    entries = AuditLoader.LoadRecentChanges(_cachedService, entity, guid);
                });
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }

            btnAudit.Enabled = true;
            btnAudit.Text = "📋 Load from Audit";

            if (errorMsg != null)
            {
                lblAuditStatus.ForeColor = Color.Red;
                lblAuditStatus.Text = $"❌ {errorMsg}";
                return;
            }

            if (entries == null || entries.Count == 0)
            {
                lblAuditStatus.ForeColor = Color.DarkOrange;
                lblAuditStatus.Text = "⚠ No audit entries found — is auditing enabled on this table?";
                return;
            }

            using (var dlg = new AuditPickerDialog(entries))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                _loadedAuditEntry = dlg.SelectedEntry;

                // Auto-populate Changed Fields from NewValue
                if (_loadedAuditEntry.NewValue != null)
                    txtChangedFields.Text = AuditLoader.BuildChangedFieldsString(_loadedAuditEntry.NewValue);

                // Auto-set message to Update
                cboMessage.SelectedItem = "Update";

                lblAuditStatus.ForeColor = Color.DarkGreen;
                lblAuditStatus.Text = $"📋 Audit loaded: {_loadedAuditEntry.ChangedOn.ToLocalTime():yyyy-MM-dd HH:mm} " +
                                      $"by {_loadedAuditEntry.ChangedBy} — " +
                                      $"{_loadedAuditEntry.ChangedFieldsSummary}";
            }
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            if (_isRunning)
            {
                MessageBox.Show("Simulation is already running.", "Busy", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validate
            var errors = new List<string>();
            if (string.IsNullOrWhiteSpace(txtEnvUrl.Text)) errors.Add("Environment URL is required.");
            if (string.IsNullOrWhiteSpace(txtEntity.Text)) errors.Add("Entity name is required.");
            if (string.IsNullOrWhiteSpace(txtRecordId.Text)) errors.Add("Record ID is required.");
            if (string.IsNullOrWhiteSpace(txtAssemblyPath.Text)) errors.Add("Assembly path is required.");
            if (!File.Exists(txtAssemblyPath.Text)) errors.Add($"Assembly not found: {txtAssemblyPath.Text}");

            if (errors.Count > 0)
            {
                MessageBox.Show(string.Join("\n", errors), "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Save settings for next launch
            SaveSettings();

            // Build config
            var config = new SimulationConfig
            {
                EnvironmentUrl = txtEnvUrl.Text.Trim(),
                EntityLogicalName = txtEntity.Text.Trim(),
                RecordId = txtRecordId.Text.Trim(),
                AssemblyPath = txtAssemblyPath.Text.Trim(),
                PluginTypeName = txtPluginType.Text.Trim(),
                MessageName = cboMessage.SelectedItem.ToString(),
                Stage = GetStageValue(),
                Debug = chkDebug.Checked,
                PreImageName = string.IsNullOrWhiteSpace(txtPreImageName.Text) ? "PreImage" : txtPreImageName.Text.Trim()
            };

            // If an audit entry was loaded, apply old values as pre-image overrides
            if (_loadedAuditEntry?.OldValue != null)
            {
                foreach (var attr in _loadedAuditEntry.OldValue.Attributes)
                    config.PreImageOverrides[attr.Key] = attr.Value;
            }

            // Parse changed fields
            if (!string.IsNullOrWhiteSpace(txtChangedFields.Text))
            {
                foreach (var pair in txtChangedFields.Text.Split(','))
                {
                    var parts = pair.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                        config.ChangedFields[parts[0].Trim()] = parts[1].Trim();
                }
            }

            // Run simulation
            _isRunning = true;
            btnRun.Enabled = false;
            btnRun.Text = "⏳  Running...";
            lblStatus.Text = "Connecting to D365...";
            rtbOutput.Clear();

            // Redirect console to output panel
            var redirector = new ConsoleRedirector(rtbOutput);
            var originalOut = Console.Out;
            Console.SetOut(redirector);

            try
            {
                await Task.Run(() =>
                {
                    // Step 1: Connect (reuse cached connection if URL unchanged)
                    UpdateStatus("Connecting to D365...");
                    if (_cachedService == null || _cachedEnvUrl != config.EnvironmentUrl)
                    {
                        _cachedService = AuthManager.Connect(config.EnvironmentUrl, config.ConnectionString);
                        _cachedEnvUrl  = config.EnvironmentUrl;
                    }
                    var realService = _cachedService;

                    // Step 2: Load plugin
                    UpdateStatus("Loading plugin assembly...");
                    var plugin = PluginLoader.Load(config.AssemblyPath, config.PluginTypeName);

                    // Step 3: Run simulation
                    UpdateStatus("Executing plugin...");
                    var result = SimulationRunner.Run(plugin, realService, config);

                    // Step 4: Report
                    UpdateStatus("Generating report...");
                    ConsoleReporter.Report(result, config);

                    UpdateStatus(result.Success ? "✅ Simulation completed successfully" : "❌ Simulation failed — see output");
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n  FATAL ERROR: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"  Inner: {ex.InnerException.Message}");
                UpdateStatus($"❌ Error: {ex.Message}");
            }
            finally
            {
                Console.SetOut(originalOut);
                _isRunning = false;

                if (!IsDisposed)
                {
                    BeginInvoke(new Action(() =>
                    {
                        btnRun.Enabled = true;
                        btnRun.Text = "▶  Run Simulation";
                    }));
                }
            }
        }

        private int GetStageValue()
        {
            var selected = cboStage.SelectedItem?.ToString() ?? "";
            if (selected.StartsWith("10")) return 10;
            if (selected.StartsWith("20")) return 20;
            return 40;
        }

        private void UpdateStatus(string text)
        {
            if (InvokeRequired)
                BeginInvoke(new Action(() => lblStatus.Text = text));
            else
                lblStatus.Text = text;
        }
    }
}
