using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace FinDataAlphaVantage
{
    internal static class Program
    {
        [STAThread] // STA: single-threaded apartment, required for certain features
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public partial class MainForm : Form
    {
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

        string _selectedDirectory;
        string _iniFullPath;

        public MainForm()
        {
            InitializeComponent();
            InitializeUI();
            InitializeDefaultSettings();
        }

        async void GetDataButton_Click(object sender, EventArgs e)
        {
            if(!Path.Exists(_selectedDirectory))
            {
                MessageBox.Show("No valid directory provided.");
                return;
            }
            
            _getDataButtonLabel.Text = "Data request sent.";
            SaveDataToIni();
                                    
            FinDataSet finDataSet = new FinDataSet(_apiKeyBox.Text, _symbolBox.Text, TimeSeries.Daily);
            await finDataSet.GetRawDataFromAlphaVantage();

            string fileName = finDataSet.Symbol + "_" + finDataSet.Series + ".csv";
            finDataSet.WriteRawDataToFile(Path.Combine(_selectedDirectory, fileName));
            
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
            _iniFullPath = Path.Combine(AppContext.BaseDirectory, "FinDataAV.ini");
            LoadDataFromIni();
        }


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

        TimeSeries SelectedTimeSeries()
        {            
            string? selection = _timeSeriesBox.SelectedItem?.ToString();
            if ( selection is null || selection == "" ) return TimeSeries.None;
            if ( selection == "Intraday" )              return TimeSeries.Intraday;
            if ( selection == "Daily" )                 return TimeSeries.Daily;
            if ( selection == "Daily Adjusted" )        return TimeSeries.DailyAdjusted;
            if ( selection == "Weekly" )                return TimeSeries.Weekly;
            if ( selection == "Weekly Adjusted" )       return TimeSeries.WeeklyAdjusted;
            if ( selection == "Monthly" )               return TimeSeries.Monthly;
            if ( selection == "Monthly Adjusted" )      return TimeSeries.MonthlyAdjusted;
            return TimeSeries.None;
        }
    }
}
