using System.Text.Json.Nodes;

namespace PdfTableExtractorDesktop
{
    public static class Components
    {
        private static string? cachedUserSelectedXlsDirectory = string.Empty;

        public static Panel CreateSettingsPanel()
        {

            var comparisonMethods = new[] { "less/equal", "equal", "greater/equal" };
            var oneThruTen = Enumerable.Range(1, 10).ToArray();
            var bigBoldFont = new Font(FontFamily.GenericSansSerif, 20, FontStyle.Bold);
            var emptyColumnSkipGroup = new List<RadioButton>();
            var emptyRowSkipGroup = new List<RadioButton>();
            var panel = new Panel { Size = new Size(600, 1000) };

            AddSettingsSection("Page Filters", 10, panel, bigBoldFont);
            panel.Controls.Add(NewLabel(new Point(20, 50), "Keep pages with number of rows"));
            var rowComparisonSelector = NewComboBox(new Point(255, 45), 100, comparisonMethods[Settings.RowComparisonMethod], comparisonMethods);
            panel.Controls.Add(rowComparisonSelector);
            panel.Controls.Add(NewLabel(new Point(360, 50), "than/to"));
            var rowsPerPageSelector = NewComboBox(new Point(420, 45), 50, Settings.RowsPerPage, oneThruTen);
            panel.Controls.Add(rowsPerPageSelector);
            panel.Controls.Add(NewLabel(new Point(20, 90), "Keep pages with number of columns"));
            var columnComparisonSelector = NewComboBox(new Point(255, 85), 100, comparisonMethods[Settings.ColumnComparisonMethod], comparisonMethods);
            panel.Controls.Add(columnComparisonSelector);
            panel.Controls.Add(NewLabel(new Point(360, 90), "than/to"));
            var columnsPerPageSelector = NewComboBox(new Point(420, 85), 50, Settings.ColumnsPerPage, oneThruTen);
            panel.Controls.Add(columnsPerPageSelector);


            panel.Controls.Add(NewLabel(new Point(20, 130), "Skip empty rows:"));
            var emptyRowGroup = new GroupBox
            {
                Location = new Point(255, 115),
                Size = new Size(300, 30)
            };
            panel.Controls.Add(emptyRowGroup);
            var emptyRowOptions = new[] { "None", "Leading", "Trailing", "Both" };
            var emptyRowPositions = new[] { new Point(0, 5), new Point(65, 5), new Point(150, 5), new Point(240, 5) };
            var emptyRowWidths = new[] { 60, 80, 80, 60 };

            for (int i = 0; i < emptyRowOptions.Length; i++)
            {
                var radio = NewRadioButton(emptyRowPositions[i], emptyRowWidths[i], emptyRowOptions[i], i, Settings.EmptyRowSkipMethod == i);
                emptyRowSkipGroup.Add(radio);
                emptyRowGroup.Controls.Add(radio);
            }

            panel.Controls.Add(NewLabel(new Point(20, 160), "Skip empty columns:"));
            var emptyColumnGroup = new GroupBox
            {
                Location = new Point(255, 150),
                Size = new Size(300, 30)
            };
            panel.Controls.Add(emptyColumnGroup);
            var emptyColumnPositions = new[] { new Point(0, 5), new Point(65, 5), new Point(150, 5), new Point(240, 5) };

            for (int i = 0; i < emptyRowOptions.Length; i++)
            {
                var radio = NewRadioButton(emptyColumnPositions[i], emptyRowWidths[i], emptyRowOptions[i], i,
                    Settings.EmptyColumnSkipMethod == i);
                emptyColumnSkipGroup.Add(radio);
                emptyColumnGroup.Controls.Add(radio);
            }

            AddSettingsSection("Page Output", 200, panel, bigBoldFont);
            panel.Controls.Add(NewLabel(new Point(20, 240), "Page naming strategy: "));
            var pageNamingMethods = new[] { "Counting", "PageOrdinal - TableOrdinal" };
            var pageNamingComboBox = NewComboBox(new Point(255, 240), 150, pageNamingMethods[Settings.PageNamingMethod], pageNamingMethods);
            panel.Controls.Add(pageNamingComboBox);
            var autosizeColumnsCheckBox = NewCheckBox(new Point(255, 280), "Autosize columns after extraction", Settings.AutosizeColumns);
            panel.Controls.Add(autosizeColumnsCheckBox);

            AddSettingsSection("File Output", 320, panel, bigBoldFont);
            panel.Controls.Add(NewLabel(new Point(20, 360), "Output Path:"));
            var outputPathGroup = new GroupBox
            {
                Location = new Point(20, 390),
                Size = new Size(450, 30)
            };
            panel.Controls.Add(outputPathGroup);
            var outputDirectoryGroup = new List<RadioButton>();
            var outDirPdfDirOption = NewRadioButton(new Point(0, 5), 150, "Input .pdf directory", 0, Settings.OutputDirectoryMethod == 0);
            var outDirSelectOption = NewRadioButton(new Point(150, 5), 150, "Select before extraction", 1, Settings.OutputDirectoryMethod == 1);
            var outDirCustomOption = NewRadioButton(new Point(300, 5), 150, "Predefined directory", 2, Settings.OutputDirectoryMethod == 2);
            outputPathGroup.Controls.Add(outDirPdfDirOption);
            outputPathGroup.Controls.Add(outDirSelectOption);
            outputPathGroup.Controls.Add(outDirCustomOption);

            var customOutDirField = NewTextField(Settings.UserSelectedOutputDirectory, new Point(255, 360), 250);
            var customOutDirChooserButton = NewButton("...", new Point(520, 355), 40, (s, e) =>
            {
                customOutDirField.Text = ShowDirectorySelection(customOutDirField.Text);
            });

            outDirPdfDirOption.CheckedChanged += (s, e) => HandleOutDirOptionChange(false, customOutDirField, customOutDirChooserButton);
            outDirSelectOption.CheckedChanged += (s, e) => HandleOutDirOptionChange(false, customOutDirField, customOutDirChooserButton);
            outDirCustomOption.CheckedChanged += (s, e) => HandleOutDirOptionChange(true, customOutDirField, customOutDirChooserButton);

            HandleOutDirOptionChange(Settings.OutputDirectoryMethod == 2, customOutDirField, customOutDirChooserButton);

            panel.Controls.Add(customOutDirField);
            panel.Controls.Add(customOutDirChooserButton);

            var parallelCheckBox = NewCheckBox(new Point(15, 470), "Enable parallel file processing", Settings.ParallelExtraction);
            var pdfContextMenuCheckBox = NewCheckBox(new Point(15, 500), "Enable PDF extraction context menu", Settings.ContextMenuOptionEnabled);
            var versionCheckingDisabledBox = NewCheckBox(new Point(15, 530), "Disable version checking", Settings.VersionCheckingDisabled);
            pdfContextMenuCheckBox.Enabled = Settings.IsWindows();

            AddSettingsSection("App Settings", 430, panel, bigBoldFont);
            panel.Controls.Add(parallelCheckBox);
            panel.Controls.Add(pdfContextMenuCheckBox);
            panel.Controls.Add(versionCheckingDisabledBox);

            Application.ApplicationExit += (s, e) =>
            {
                var userSelectedDirectoryString = customOutDirField.Text;
                //var userSelectedDirectoryMethod = -1;
                //foreach (Control control in outputPathGroup.Controls)
                //{
                //    if (control is RadioButton radioButton && radioButton.Checked)
                //    {
                //        userSelectedDirectoryMethod = (int)radioButton.Tag!;
                //        break;
                //    }
                //}

                var userSelectedDirectoryMethod = outputPathGroup.Controls
                                                                    .OfType<RadioButton>()
                                                                    .ToList()
                                                                    .FindIndex(radioButton => radioButton.Checked);

                var isUserDirectoryValid = !string.IsNullOrWhiteSpace(userSelectedDirectoryString) && Directory.Exists(userSelectedDirectoryString);

                var settings = new JsonObject
                {
                    ["RowsPerPage"] = (int)rowsPerPageSelector.SelectedItem!,
                    ["RowComparisonMethod"] = Array.IndexOf(comparisonMethods, (string)rowComparisonSelector.SelectedItem!),
                    ["ColumnsPerPage"] = (int)columnsPerPageSelector.SelectedItem!,
                    ["ColumnComparisonMethod"] = Array.IndexOf(comparisonMethods, (string)columnComparisonSelector.SelectedItem!),
                    ["AutosizeColumns"] = autosizeColumnsCheckBox.Checked,
                    ["ParallelFileProcess"] = parallelCheckBox.Checked,
                    ["PageNamingMethod"] = Array.IndexOf(pageNamingMethods, (string)pageNamingComboBox.SelectedItem!),
                    ["EmptyColumnSkipMethod"] = emptyColumnSkipGroup.FindIndex(r => r.Checked),
                    ["EmptyRowSkipMethod"] = emptyRowSkipGroup.FindIndex(r => r.Checked),
                    ["OutputDirectoryMethod"] = userSelectedDirectoryMethod != 2 ? userSelectedDirectoryMethod : isUserDirectoryValid ? 2 : 0,
                    ["OutputDirectoryCustom"] = isUserDirectoryValid ? userSelectedDirectoryString : "",
                    ["VersionCheckingDisabled"] = versionCheckingDisabledBox.Checked,
                    ["ContextMenuOptionEnabled"] = pdfContextMenuCheckBox.Checked
                };

                Settings.Save(settings, pdfContextMenuCheckBox.Checked);
            };

            return panel;
        }

        public static string ShowDirectorySelection(string rootDir)
        {
            string selectedPath = "";

            var t = new Thread(() =>
            {
                using var folderDialog = new FolderBrowserDialog
                {
                    SelectedPath = rootDir,
                    Description = "Select Directory"
                };

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedPath = folderDialog.SelectedPath;
                }
            });

            t.SetApartmentState(ApartmentState.STA); // Required for UI components
            t.Start();
            t.Join(); // Wait for thread to complete

            return selectedPath;
        }


        public static string ShowExcelDirectoryPicker(string rootDir)
        {
            //try
            //{
            //    Application.SetHighDpiMode(HighDpiMode.SystemAware);
            //    Application.EnableVisualStyles();
            //    Application.SetCompatibleTextRenderingDefault(false);
            //}
            //catch (Exception) { }

            if (string.IsNullOrEmpty(cachedUserSelectedXlsDirectory))
            {
                cachedUserSelectedXlsDirectory = ShowDirectorySelection(rootDir);
            }

            return cachedUserSelectedXlsDirectory;
        }

        private static void HandleOutDirOptionChange(bool enabled, TextBox customOutDirField, Button customOutDirChooserButton)
        {
            customOutDirField.Enabled = enabled;
            customOutDirChooserButton.Enabled = enabled;
        }

        private static Label NewLabel(Point location, string text)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Size = new Size(text.Length * 7, 30),
                AutoSize = false
            };
        }

        private static CheckBox NewCheckBox(Point location, string text, bool selected)
        {
            return new CheckBox
            {
                Text = text,
                Location = location,
                Size = new Size(text.Length * 18, 30),
                Checked = selected
            };
        }

        private static TextBox NewTextField(string defaultValue, Point location, int width)
        {
            return new TextBox
            {
                Text = defaultValue,
                Location = location,
                Size = new Size(width, 30)
            };
        }

        private static Button NewButton(string text, Point location, int width, EventHandler handler)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(width, 40),
                BackColor = Color.DarkGray
            };
            button.Click += handler;
            return button;
        }

        private static RadioButton NewRadioButton(Point location, int width, string text, int index, bool selected)
        {
            return new RadioButton
            {
                Text = text,
                Location = location,
                Size = new Size(width, 30),
                Tag = index,
                Checked = selected
            };
        }

        private static ComboBox NewComboBox<T>(Point location, int width, T selectedValue, T[] elements)
        {
            var comboBox = new ComboBox
            {
                Location = location,
                Size = new Size(width, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            comboBox.Items.AddRange(elements.Cast<object>().ToArray());
            comboBox.SelectedItem = selectedValue;

            return comboBox;
        }

        private static void AddSettingsSection(string text, int y, Panel contentPanel, Font font)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(20, y),
                Size = new Size(text.Length * 18, 30),
                Font = font
            };
            contentPanel.Controls.Add(label);

            var separator = new Panel
            {
                Location = new Point(0, y + 30),
                Size = new Size(600, 2),
                BackColor = Color.Gray
            };
            contentPanel.Controls.Add(separator);
        }
    }
}