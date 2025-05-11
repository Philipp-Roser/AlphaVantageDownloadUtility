using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace FinDataAlphaVantage
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    internal class AVDataSet
    {
        internal const string MyApiKey = "";
        internal readonly string ApiKey = "";
        internal readonly string Symbol = "";
        internal readonly TimeSeries Series = TimeSeries.None;

        internal string RawData { get; private set; } = "";

        internal bool ProcessDataSuccessful { get; private set; } = false;

        internal AVDataSet(string apiKey, string symbol, TimeSeries series)
        {
            ApiKey = apiKey;
            Symbol = symbol;
            Series = series;
        }



        internal enum TimeSeries { None, Intraday, Daily, DailyAdjusted, Weekly, WeeklyAdjusted, Monthly, MonthlyAdjusted };
    }

    public partial class MainForm : Form
    {
        string _selectedDirectory;
        string _iniFullPath;

        internal string _rawData { get; private set; } = "";
        private bool _retrieveRawDataSuccessful = false;

        public MainForm()
        {
            InitializeComponent();
            InitializeUI();
            InitializeDefaultSettings();
        }

        #region Windows Form Template Code
        
        
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if ( disposing && (components != null) )
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }


        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 540);
            this.Text = "Retrieve Alpha Vantage Financial Data";
        }

        #endregion


        #region Layout, UI, WinForms components, etc.

        Label _titleLabel;

        Label _directoryLabel;
        Button _selectDirectoryButton;

        Label _apiKeyBoxLabel;
        TextBox _apiKeyBox;

        Label _symbolBoxLabel;
        TextBox _symbolBox;

        Label _timeSeriesLabel;
        ComboBox _timeSeriesBox;

        Label _targetDirectoryTitleLabel;

        Button _getDataButton;
        Label _getDataButtonLabel;


        void InitializeUI()
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            _titleLabel = new();
            _titleLabel.Text = "Retrieve Data from Alpha Vantage:";
            _titleLabel.Font = new Font(_titleLabel.Font, FontStyle.Bold);
            _titleLabel.Location = new Point(60, 20);
            _titleLabel.Size = new Size(400, 40);
            this.Controls.Add(_titleLabel);

            _apiKeyBoxLabel = new();
            _apiKeyBoxLabel.Text = "API Key:";
            _apiKeyBoxLabel.Location = new Point(60, 80);
            _apiKeyBoxLabel.Size = new Size(80, 40);
            this.Controls.Add(_apiKeyBoxLabel);

            _apiKeyBox = new();
            _apiKeyBox.Text = "";
            _apiKeyBox.Location = new Point(200, 80);
            _apiKeyBox.Size = new Size(300, 60);
            this.Controls.Add(_apiKeyBox);


            _symbolBoxLabel = new();
            _symbolBoxLabel.Text = "Symbol:";
            _symbolBoxLabel.Location = new Point(60, 120);
            _symbolBoxLabel.Size = new Size(80, 40);
            this.Controls.Add(_symbolBoxLabel);

            _symbolBox = new();
            _symbolBox.Text = "";
            _symbolBox.Location = new Point(200, 120);
            _symbolBox.Size = new Size(300, 60);
            this.Controls.Add(_symbolBox);

            _timeSeriesLabel = new();
            _timeSeriesLabel.Text = "Time Series:";
            _timeSeriesLabel.Location = new Point(60, 160);
            _timeSeriesLabel.Size = new Size(120, 40);
            this.Controls.Add(_timeSeriesLabel);

            _timeSeriesBox = new();
            _timeSeriesBox.Items.AddRange(new string[]
                { "", "Intraday", "Daily", "Daily Adjusted", "Weekly", "Weekly Adjusted", "Monthly", "Monthly Adjusted" });
            _timeSeriesBox.Location = new Point(200, 160);
            this.Controls.Add(_timeSeriesBox);


            _targetDirectoryTitleLabel = new();
            _targetDirectoryTitleLabel.Text = "Target Data Directory:";
            _targetDirectoryTitleLabel.Font = new Font(_targetDirectoryTitleLabel.Font, FontStyle.Bold);
            _targetDirectoryTitleLabel.Location = new Point(60, 240);
            _targetDirectoryTitleLabel.AutoSize = true;
            this.Controls.Add(_targetDirectoryTitleLabel);

            _directoryLabel = new();
            _directoryLabel.Text = _selectedDirectory;
            _directoryLabel.Location = new Point(60, 280);
            _directoryLabel.AutoSize = true;
            this.Controls.Add(_directoryLabel);

            _selectDirectoryButton = new();
            _selectDirectoryButton.Text = "Select Directory";
            _selectDirectoryButton.Location = new Point(60, 320);
            _selectDirectoryButton.Size = new Size(180, 40);
            _selectDirectoryButton.Click += SelectDirectoryButton_Click;
            this.Controls.Add(_selectDirectoryButton);

            _getDataButton = new();
            _getDataButton.Text = "Get Data";
            _getDataButton.Location = new Point(60, 400);
            _getDataButton.Size = new(120, 40);
            _getDataButton.Click += GetDataButton_Click;
            this.Controls.Add(_getDataButton);

            _getDataButtonLabel = new();
            _getDataButtonLabel.Text = "Ready ...";
            _getDataButtonLabel.Location = new Point(60, 460);
            _getDataButtonLabel.AutoSize = true;
            this.Controls.Add(_getDataButtonLabel);
        }

        #endregion


        #region Button Responses
        async void GetDataButton_Click(object sender, EventArgs e)
        {
            if(!Path.Exists(_selectedDirectory))
            {
                MessageBox.Show("No valid directory provided.");
                return;
            }
            
            _getDataButtonLabel.Text = "Data request sent.";
            SaveDataToIni();
                                    
            AVDataSet avDataSet = new AVDataSet(_apiKeyBox.Text, _symbolBox.Text, AVDataSet.TimeSeries.Daily);
            await GetRawDataFromAlphaVantage();

            string fileName = avDataSet.Symbol + "_" + avDataSet.Series + ".csv";
            WriteRawDataToFile(Path.Combine(_selectedDirectory, fileName));
            
            Process.Start("explorer.exe", _selectedDirectory);

            _getDataButtonLabel.Text = "Data request complete.";
        }

        void SelectDirectoryButton_Click(object sender, EventArgs e)
        {
            using ( FolderBrowserDialog dialog = new() ) 
            {
                dialog.Description = "Select Directory";
                dialog.UseDescriptionForTitle = true;

                DialogResult result = dialog.ShowDialog();

                if ( result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath) )
                {
                    {
                        _selectedDirectory = dialog.SelectedPath;
                        _directoryLabel.Text = _selectedDirectory;
                    }
                }
            }
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string section, string key, string defaultValue, StringBuilder retVal, int size, string filePath);

        #endregion


        #region Save to / Load from .ini
        void SaveDataToIni()
        {
            WritePrivateProfileString("Defaults", "API_Key", _apiKeyBox.Text, _iniFullPath);
            WritePrivateProfileString("Defaults", "Symbol", _symbolBox.Text, _iniFullPath);
            WritePrivateProfileString("Defaults", "Time_Series", SelectedTimeSeries().ToString(), _iniFullPath);
            WritePrivateProfileString("Defaults", "Directory", _selectedDirectory, _iniFullPath);
        }

        void LoadDataFromIni()
        {
            StringBuilder apiFromIni = new StringBuilder(255);
            GetPrivateProfileString("Defaults", "API_Key", "", apiFromIni, 255, _iniFullPath);
            _apiKeyBox.Text = apiFromIni.ToString();

            StringBuilder symbolFromIni = new StringBuilder(255);
            GetPrivateProfileString("Defaults", "Symbol", "", symbolFromIni, 255, _iniFullPath);
            _symbolBox.Text = symbolFromIni.ToString();

            StringBuilder timeSeriesFromIni = new StringBuilder();
            GetPrivateProfileString("Defaults", "Time_Series", "", timeSeriesFromIni, 255, _iniFullPath);
            _timeSeriesBox.SelectedItem = timeSeriesFromIni.ToString();

            StringBuilder directoryFromIni = new StringBuilder(255);
            GetPrivateProfileString("Defaults", "Directory", "", directoryFromIni, 255, _iniFullPath);
            _selectedDirectory = directoryFromIni.ToString();
            if(_selectedDirectory == "" )
                _selectedDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _directoryLabel.Text = _selectedDirectory;
        }

        void InitializeDefaultSettings()
        {
            _iniFullPath = Path.Combine(AppContext.BaseDirectory, "AVDataRetrieve.ini");
            LoadDataFromIni();
        }

        #endregion


        #region Data Handling

        internal async Task GetRawDataFromAlphaVantage()
        {
            AVDataSet.TimeSeries series = SelectedTimeSeries();
            string symbol = _symbolBox.Text;
            string apiKey = _apiKeyBox.Text;


            string url = $"https://www.alphavantage.co/query?function={ForURLString(series)}&symbol={symbol}&apikey={apiKey}&datatype=csv";

            using HttpClient client = new HttpClient();
            try
            {
                _rawData = await client.GetStringAsync(url);
                _retrieveRawDataSuccessful = true;
            }
            catch ( Exception exception )
            {
                MessageBox.Show(exception.Message, "Error");
            }
        }

        internal void WriteRawDataToFile(string completeFilePath)
        {
            Debug.Assert(_retrieveRawDataSuccessful, "Cannot print data. Retrieve was not successful (yet).");
            File.WriteAllText(completeFilePath, _rawData);
        }

        #endregion


        #region TimeSeries: enum --> string; and selection --> enum
        private static string ForURLString(AVDataSet.TimeSeries timeSeries)
        {
            switch ( timeSeries )
            {
                case AVDataSet.TimeSeries.None:
                    return "";
                case AVDataSet.TimeSeries.Intraday:
                    return "TIME_SERIES_INTRADAY";
                case AVDataSet.TimeSeries.Daily:
                    return "TIME_SERIES_DAILY";
                case AVDataSet.TimeSeries.DailyAdjusted:
                    return "TIME_SERIES_DAILY_ADJUSTED";
                case AVDataSet.TimeSeries.Weekly:
                    return "TIME_SERIES_WEEKLY";
                case AVDataSet.TimeSeries.WeeklyAdjusted:
                    return "TIME_SERIES_WEEKLY_ADJUSTED";
                case AVDataSet.TimeSeries.Monthly:
                    return "TIME_SERIES_MONTHLY";
                case AVDataSet.TimeSeries.MonthlyAdjusted:
                    return "TIME_SERIES_MONTHLY_ADJUSTED";
                default:
                    return "";

            }

        }
        AVDataSet.TimeSeries SelectedTimeSeries()
        {
            string? selection = _timeSeriesBox.SelectedItem?.ToString();
            if ( selection is null || selection == "" ) return AVDataSet.TimeSeries.None;
            if ( selection == "Intraday" ) return AVDataSet.TimeSeries.Intraday;
            if ( selection == "Daily" ) return AVDataSet.TimeSeries.Daily;
            if ( selection == "Daily Adjusted" ) return AVDataSet.TimeSeries.DailyAdjusted;
            if ( selection == "Weekly" ) return AVDataSet.TimeSeries.Weekly;
            if ( selection == "Weekly Adjusted" ) return AVDataSet.TimeSeries.WeeklyAdjusted;
            if ( selection == "Monthly" ) return AVDataSet.TimeSeries.Monthly;
            if ( selection == "Monthly Adjusted" ) return AVDataSet.TimeSeries.MonthlyAdjusted;
            return AVDataSet.TimeSeries.None;
        }

        #endregion


    }
}
