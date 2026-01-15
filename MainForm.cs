using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Win11Debloat
{
    public class MainForm : Form
    {
        private FlowLayoutPanel mainPanel = null!;
        private ProgressBar progressBar = null!;
        private Label statusLabel = null!;
        private Button btnRemoveEdge = null!;
        private FlowLayoutPanel tweaksPanel;

        private readonly List<Tweak> tweaks = new();
        private readonly List<WingetApp> wingetApps = new();

        public MainForm()
        {
            Text = "Windows 11 Optimisation Utility";
            Size = new Size(1100, 700);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;

            BuildUI();
            ApplyTheme();
            RegisterTweaks();
            RegisterWingetApps();
        }

        #region UI

        private void BuildUI()
        {
            // ----------------------
            // Sidebar
            // ----------------------
            var sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 300,
                Padding = new Padding(12),
                BackColor = Color.FromArgb(225, 225, 225) // slightly darker than panel for contrast
            };

            // Apply / Recommended buttons
            var btnApply = CreateButton("Apply Selected Tweaks", async () => await ApplyTweaksAsync());
            var btnRecommended = CreateButton("Recommended Tweaks", SelectRecommended);

            // Microsoft Edge removal button
            btnRemoveEdge = CreateButton("Remove Microsoft Edge (Advanced)", RemoveEdge);
            btnRemoveEdge.Height = 48;

            // Add buttons to sidebar
            sidebar.Controls.Add(btnApply);
            sidebar.Controls.Add(btnRecommended);
            sidebar.Controls.Add(new Label { Height = 16 }); // spacing
            sidebar.Controls.Add(btnRemoveEdge);

            // ----------------------
            // Main tweaks panel
            // ----------------------
            tweaksPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(16),
                BackColor = Color.FromArgb(240, 240, 240) // light grey for Win11 look
            };

            // Add toggles for each registered tweak
            foreach (var tweak in tweaks)
            {
                var toggle = CreateToggle(tweak.Name);
                toggle.Checked = tweak.Recommended;
                tweak.Toggle = toggle;
                tweaksPanel.Controls.Add(toggle);
            }

            // ----------------------
            // Bottom status bar
            // ----------------------
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 18
            };

            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Text = "Ready",
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0)
            };

            // ----------------------
            // Add everything to Form
            // ----------------------
            Controls.Add(tweaksPanel);
            Controls.Add(progressBar);
            Controls.Add(statusLabel);
            Controls.Add(sidebar);
        }


        private Button CreateButton(string text, Action onClick)
        {
            var btn = new Button
            {
                Text = text,
                Dock = DockStyle.Top, // ensures it fills width of container
                Height = 46,
                Margin = new Padding(0, 4, 0, 4), // vertical spacing
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI Variable", 10.5f)
            };
            btn.Click += (_, _) => onClick();
            return btn;
        }


        private static CheckBox CreateToggle(string text)
        {
            return new CheckBox
            {
                Text = text,
                AutoSize = false,
                Width = 720,
                Height = 42,
                Padding = new Padding(12, 0, 0, 0),
                Font = new Font("Segoe UI Variable", 10.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                UseCompatibleTextRendering = false
            };
        }

        private static Label CreateSectionHeader(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Segoe UI Variable", 12f, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(0, 20, 0, 10)
            };
        }

        #endregion

        #region Theme

        private void ApplyTheme()
        {
            bool dark = IsDarkMode();
            var bg = dark ? Color.FromArgb(32, 32, 32) : Color.FromArgb(240, 240, 240); // light grey
            var fg = dark ? Color.White : Color.Black;

            BackColor = bg;
            ForeColor = fg;

            ApplyThemeRecursive(this, bg, fg);
        }




        private static void ApplyThemeRecursive(Control c, Color bg, Color fg)
        {
            c.BackColor = bg;
            c.ForeColor = fg;

            foreach (Control child in c.Controls)
                ApplyThemeRecursive(child, bg, fg);
        }

        private static bool IsDarkMode()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                return key?.GetValue("AppsUseLightTheme") is int v && v == 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Tweaks

        private sealed record Tweak(string Name, bool Recommended, Action Apply)
        {
            public CheckBox Toggle { get; set; } = null!;
        }

        private void RegisterTweaks()
        {
            tweaksPanel.Controls.Add(CreateSectionHeader("System Tweaks"));

            AddTweak("Disable Advertising ID", true,
                () => SetCU(@"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo", "Enabled", 0));

            AddTweak("Show File Extensions", true,
                () => SetCU(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0));

            AddTweak("Disable Windows Tips", true,
                () => SetCU(@"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
                    "SubscribedContent-338388Enabled", 0));

            AddTweak("Disable Bing Search in Start Menu", true,
                () => SetCU(@"Software\Microsoft\Windows\CurrentVersion\Search",
                    "BingSearchEnabled", 0));

            AddTweak("Disable Windows Copilot", true,
                () => DisableCopilot(IsAdmin()));
        }

        private void AddTweak(string name, bool recommended, Action apply)
        {
            var tweak = new Tweak(name, recommended, apply)
            {
                Toggle = CreateToggle(name)
            };

            tweaks.Add(tweak);
            tweaksPanel.Controls.Add(tweak.Toggle);
        }

        #endregion

        #region Actions

        private async Task ApplyTweaksAsync()
        {
            progressBar.Value = 0;
            statusLabel.Text = "Applying tweaks...";

            var selected = tweaks.FindAll(t => t.Toggle.Checked);
            if (selected.Count == 0)
            {
                statusLabel.Text = "No tweaks selected.";
                return;
            }

            int step = 100 / selected.Count;

            foreach (var tweak in selected)
            {
                await Task.Run(() => tweak.Apply());
                progressBar.Value = Math.Min(100, progressBar.Value + step);
            }

            progressBar.Value = 100;
            statusLabel.Text = "Tweaks applied successfully.";
        }

        private void SelectRecommended()
        {
            foreach (var tweak in tweaks)
                tweak.Toggle.Checked = tweak.Recommended;

            statusLabel.Text = "Recommended tweaks selected.";
        }

        #endregion


        #region WinGet

        private sealed record WingetApp(string Name, string Id)
        {
            public CheckBox Toggle { get; set; } = null!;
        }

        private void RegisterWingetApps()
        {
            tweaksPanel.Controls.Add(CreateSectionHeader("Essential Apps (WinGet)"));

            AddWingetApp("Google Chrome", "Google.Chrome");
            AddWingetApp("Mozilla Firefox", "Mozilla.Firefox");
            AddWingetApp("Brave Browser", "Brave.Brave");
            AddWingetApp("7-Zip", "7zip.7zip");
            AddWingetApp("Notepad++", "Notepad++.Notepad++");
            AddWingetApp("Everything (Voidtools)", "voidtools.Everything");
            AddWingetApp("Visual Studio Code", "Microsoft.VisualStudioCode");
            AddWingetApp("Git", "Git.Git");
            AddWingetApp("VLC Media Player", "VideoLAN.VLC");
        }

        private void AddWingetApp(string name, string id)
        {
            var app = new WingetApp(name, id)
            {
                Toggle = CreateToggle(name)
            };

            wingetApps.Add(app);
            tweaksPanel.Controls.Add(app.Toggle);
        }

        private async Task InstallWingetAppsAsync()
        {
            if (!IsWingetAvailable())
            {
                MessageBox.Show("WinGet is not available on this system.");
                return;
            }

            var selected = wingetApps.FindAll(a => a.Toggle.Checked);
            if (selected.Count == 0)
                return;

            progressBar.Value = 0;
            statusLabel.Text = "Installing applications...";

            int step = 100 / selected.Count;

            foreach (var app in selected)
            {
                await Task.Run(() =>
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "winget",
                        Arguments = $"install --id {app.Id} --silent --accept-package-agreements --accept-source-agreements",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    })?.WaitForExit();
                });

                progressBar.Value = Math.Min(100, progressBar.Value + step);
            }

            progressBar.Value = 100;
            statusLabel.Text = "Applications installed.";
        }

        private static bool IsWingetAvailable()
        {
            try
            {
                var p = Process.Start(new ProcessStartInfo
                {
                    FileName = "winget",
                    Arguments = "--version",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                p?.WaitForExit(2000);
                return p?.ExitCode == 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Edge / Helpers

        private static void SetCU(string path, string name, object value)
        {
            using var key = Registry.CurrentUser.CreateSubKey(path);
            key?.SetValue(name, value);
        }

        private static bool IsAdmin()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void DisableCopilot(bool admin)
        {
            var hive = admin ? Registry.LocalMachine : Registry.CurrentUser;

            using (var k = hive.CreateSubKey(
                @"SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"))
            {
                k?.SetValue("TurnOffWindowsCopilot", 1, RegistryValueKind.DWord);
            }

            using (var t = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"))
            {
                t?.SetValue("ShowCopilotButton", 0, RegistryValueKind.DWord);
            }

            RestartExplorer();
        }

        private static void RestartExplorer()
        {
            foreach (var p in Process.GetProcessesByName("explorer"))
                p.Kill();

            Process.Start("explorer.exe");
        }

        private void RemoveEdge()
        {
            if (!IsAdmin())
            {
                MessageBox.Show("Administrator privileges required.");
                return;
            }

            var confirm = MessageBox.Show(
                "This will remove Microsoft Edge.\nWindows Update may reinstall it.\n\nContinue?",
                "Confirm Edge Removal",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
                return;

            try
            {
                var edgePath = @"C:\Program Files (x86)\Microsoft\Edge\Application";
                var versionDir = Directory.GetDirectories(edgePath)[0];
                var installer = Path.Combine(versionDir, "Installer", "setup.exe");

                Process.Start(new ProcessStartInfo
                {
                    FileName = installer,
                    Arguments = "--uninstall --system-level --force-uninstall",
                    Verb = "runas",
                    UseShellExecute = true
                });

                statusLabel.Text = "Removing Microsoft Edge...";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to remove Edge:\n{ex.Message}");
            }
        }

        #endregion
    }
}
