using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using Infrastructure.Interfaces;
using Dto;
using Infrastructure.Commands;
using System.Diagnostics;
using Logging;
using System.Threading;
using System.Windows.Controls;
using DocumentFormat.OpenXml.Office2019.Drawing.Animation;

namespace Infrastructure.ViewModel
{
    public class MoscowReportVM : INotifyPropertyChanged, IViewModel
    {
        private DateTime _limitDate;
        private bool _onlyReportsA;
        private string? _moscowTablePath;
        private string? _piktsTablePath;
        private string? _reportsA_Path;
        private readonly string _filter;

        private RelayCommand? _moscowTablePathCommand;
        private RelayCommand? _piktsTablePathCommand;
        private RelayCommand? _annexesPathCommand;
        private RelayCommand? _runCommand;

        private readonly OpenFileDialog _fileDialog;
        private readonly OpenFolderDialog _folderDialog;

        private Progress<object[]> _progress;
        bool _infinite;
        int _statusValue;
        int _statusMax = 100;
        string _statusText = string.Empty;

        private ILauncher Launcher { get; set; }
        private MainFormDto MainForm { get; set; }
        public event PropertyChangedEventHandler? PropertyChanged;
        public MoscowReportVM(ILauncher launcher, IProgress<object[]> progress)
        {
            Launcher = launcher;
            MainForm = MainFormDto.GetInstance();
            _limitDate = DateTime.Now.AddMonths(-5);
            _progress = (Progress<object[]>)progress;
            _progress.ProgressChanged += ChangeProgressBar;
            _folderDialog = new OpenFolderDialog()
            {
                Title = "Укажите корневой каталог с приложениями А"
            };

            _filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            _fileDialog = new OpenFileDialog()
            {
                Filter = _filter,
            };
        }
        public DateTime LimitDate
        {
            get => _limitDate;
            set
            {
                if (_limitDate != value)
                {
                    _limitDate = value;
                    MainForm.LimitDate = value;
                    OnPropertyChanged("LimitDate");
                }
            }
        }
        public bool OnlyReportsA
        {
            get => _onlyReportsA;
            set
            {
                if (_onlyReportsA != value)
                {
                    _onlyReportsA = value;
                    MainForm.OnlyReportsA = value;
                    OnPropertyChanged("OnlyReportsA");
                }
            }
        }
        public string? MoscowTablePath
        {
            get => _moscowTablePath;
            set
            {
                if (_moscowTablePath != value)
                {
                    _moscowTablePath = value;
                    MainForm.MoscowTablePath = value!;
                    OnPropertyChanged("MoscowTablePath");
                }
            }
        }
        public string? PiktsTablePath
        {
            get => _piktsTablePath;
            set
            {
                if (_piktsTablePath != value)
                {
                    _piktsTablePath = value;
                    MainForm.PiktsTablePath = value!;
                    OnPropertyChanged("PiktsTablePath");
                }
            }
        }
        public string? ReportsA_Path
        {
            get => _reportsA_Path;
            set
            {
                if (_reportsA_Path != value)
                {
                    _reportsA_Path = value;
                    MainForm.ReportsA_Path = value!;
                    OnPropertyChanged("ReportsA_Path");
                }
            }
        }
        public RelayCommand MoscowTablePathCommand
        {
            get
            {
                return _moscowTablePathCommand ??= new RelayCommand((ob) =>
                {
                    MoscowTablePath = GetTablePath("Выберите файл с московской таблицей");
                });
            }
        }
        public RelayCommand PiktsTablePathCommand
        {
            get
            {
                return _piktsTablePathCommand ??= new RelayCommand((ob) =>
                {
                    PiktsTablePath = GetTablePath("Выберите файл с таблицей из ПИКТС");
                });
            }
        }
        public RelayCommand ReportsA_PathCommand
        {
            get
            {
                return _annexesPathCommand ??= new RelayCommand((ob) =>
                {
                    if (_folderDialog.ShowDialog() == false) return;

                    ReportsA_Path = _folderDialog.FolderName;
                });
            }
        }
        public RelayCommand RunCommand
        {
            get
            {
                return _runCommand ??= new RelayCommand((ob) =>
                {

                    Thread work = new Thread(() =>
                    {
                        try
                        {
                            Logger.ClearLogFile();
                            Launcher.Run();
                            MessageBox.Show("\tГОТОВО");
                            Logger.OpenLogFile();
                        }
                        catch (UnreachableException ex)
                        {
                            MessageBox.Show(ex.Message);
                            Logger.OpenLogFile();
                        }
                    });
                    work.IsBackground = true;
                    work.Start();
                }, (ob) => IsPathsValid());
            }
        }
        public bool Infinite
        {
            get => _infinite;
            set
            {
                if (_infinite != value)
                {
                    _infinite = value;
                    OnPropertyChanged("Infinite");
                }
            }
        }
        public int StatusValue
        {
            get => _statusValue;
            set
            {
                if (_statusValue != value)
                {
                    _statusValue = value;
                    OnPropertyChanged("StatusValue");
                }
            }
        }
        public int StatusMax
        {
            get => _statusMax;
            set
            {
                if (_statusMax != value)
                {
                    _statusMax = value;
                    OnPropertyChanged("StatusMax");
                }
            }
        }
        public string StatusText
        {
            get => _statusText;
            set
            {
                if (_statusText != value)
                {
                    _statusText = value;
                    OnPropertyChanged("StatusText");
                }
            }
        }
        private string GetTablePath(string title)
        {
            _fileDialog.Title = title;
            _fileDialog.ShowDialog();
            return _fileDialog.FileName;
        }
        private bool IsPathsValid()
        {
            if (OnlyReportsA)
            {
                return !string.IsNullOrEmpty(MoscowTablePath) &&
                 !string.IsNullOrEmpty(ReportsA_Path);
            }
            else
            {
                return MoscowTablePath != PiktsTablePath &&
                 !string.IsNullOrEmpty(MoscowTablePath) &&
                 !string.IsNullOrEmpty(PiktsTablePath) &&
                 !string.IsNullOrEmpty(ReportsA_Path);
            }
        }
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
        private void ChangeProgressBar(object sender, object[] args)
        {
            StatusMax = (int)args[0];
            StatusValue = (int)args[1];
            StatusText = (string)args[2];
            Infinite = (bool)args[3];
        }
    }
}
