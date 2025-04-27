using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Text;
using System.Windows;

namespace DllNetFrameworkAnalyzer
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
        }

        private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            dialog.Description = "フォルダを選択してください";
            if (dialog.ShowDialog())
            {
                _viewModel.FolderPath = dialog.SelectedPath;
            }
        }

        private void BrowseOutputFileButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                DefaultExt = "csv",
                Title = "CSV出力先を選択してください"
            };

            if (dialog.ShowDialog() == true)
            {
                _viewModel.OutputFilePath = dialog.FileName;
            }
        }

        private async void AnalyzeButton_Click(object sender, RoutedEventArgs e)
        {
            // 入力検証
            if (string.IsNullOrWhiteSpace(_viewModel.FolderPath))
            {
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(_viewModel.OutputFilePath))
            {
                MessageBox.Show("出力ファイルパスを入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 解析を開始
                _viewModel.IsAnalyzing = true;
                _viewModel.ProgressValue = 0;
                _viewModel.StatusMessage = "解析を開始しています...";

                await Task.Run(() => _viewModel.StartAnalysis());

                _viewModel.IsAnalyzing = false;
                _viewModel.StatusMessage = "解析が完了しました。";

                // CSVファイルを開くかどうか確認
                var result = MessageBox.Show(
                    $"解析が完了しました。CSVファイルを開きますか？\n\n処理完了: {_viewModel.ResultCount}ファイル\n成功: {_viewModel.SuccessCount}件\n失敗: {_viewModel.FailedCount}件",
                    "処理完了",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    OpenCsv();
                }
            }
            catch (Exception ex)
            {
                _viewModel.IsAnalyzing = false;
                _viewModel.StatusMessage = $"エラーが発生しました: {ex.Message}";
                MessageBox.Show($"解析中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void OpenCsvButton_Click(object sender, RoutedEventArgs e)
        {
            OpenCsv();
        }

        private void OpenCsv()
        {
            try
            {
                // ファイルが存在するか確認
                if (File.Exists(_viewModel.OutputFilePath))
                {
                    // ProcessStartInfoを使用してより詳細に制御
                    var psi = new ProcessStartInfo
                    {
                        FileName = _viewModel.OutputFilePath,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
                else
                {
                    // ログやエラーメッセージを表示
                    Console.WriteLine($"ファイルが存在しません: {_viewModel.OutputFilePath}");
                }
            }
            catch (Exception ex)
            {
                // 例外をログ
                Console.WriteLine($"ファイルを開く際にエラーが発生しました: {ex.Message}");
            }
        }
    }

    public class MainViewModel : INotifyPropertyChanged
    {
        private string _folderPath;
        private string _outputFilePath;
        private bool _includeSubfolders = true;
        private bool _isAnalyzing;
        private int _progressValue;
        private string _statusMessage;
        private string _progressDetail;
        private int _resultCount;
        private int _successCount;
        private int _failedCount;
        private ObservableCollection<AssemblyInfo> _assemblyInfos = new ObservableCollection<AssemblyInfo>();
        private readonly BackgroundWorker _worker = new BackgroundWorker();

        public event PropertyChangedEventHandler PropertyChanged;

        public MainViewModel()
        {
            _folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            _outputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DllAnalysisResult.csv");
            _statusMessage = "準備完了";

            _worker.WorkerReportsProgress = true;
            _worker.WorkerSupportsCancellation = true;
        }

        public string FolderPath
        {
            get => _folderPath;
            set
            {
                if (_folderPath != value)
                {
                    _folderPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string OutputFilePath
        {
            get => _outputFilePath;
            set
            {
                if (_outputFilePath != value)
                {
                    _outputFilePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IncludeSubfolders
        {
            get => _includeSubfolders;
            set
            {
                if (_includeSubfolders != value)
                {
                    _includeSubfolders = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsAnalyzing
        {
            get => _isAnalyzing;
            set
            {
                if (_isAnalyzing != value)
                {
                    _isAnalyzing = value;
                    OnPropertyChanged();
                }
            }
        }

        public int ProgressValue
        {
            get => _progressValue;
            set
            {
                if (_progressValue != value)
                {
                    _progressValue = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ProgressDetail
        {
            get => _progressDetail;
            set
            {
                if (_progressDetail != value)
                {
                    _progressDetail = value;
                    OnPropertyChanged();
                }
            }
        }

        public int ResultCount
        {
            get => _resultCount;
            set
            {
                if (_resultCount != value)
                {
                    _resultCount = value;
                    OnPropertyChanged();
                }
            }
        }

        public int SuccessCount
        {
            get => _successCount;
            set
            {
                if (_successCount != value)
                {
                    _successCount = value;
                    OnPropertyChanged();
                }
            }
        }

        public int FailedCount
        {
            get => _failedCount;
            set
            {
                if (_failedCount != value)
                {
                    _failedCount = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<AssemblyInfo> AssemblyInfos => _assemblyInfos;

        public void StartAnalysis()
        {
            // 準備
            Application.Current.Dispatcher.Invoke(() =>
            {
                _assemblyInfos.Clear();
            });

            ResultCount = 0;
            SuccessCount = 0;
            FailedCount = 0;

            // ファイル検索オプション
            var searchOption = IncludeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            // DLLファイルのリストを取得
            var dllFiles = Directory.GetFiles(FolderPath, "*.dll", searchOption);
            var exeFiles = Directory.GetFiles(FolderPath, "*.exe", searchOption);
            var allFiles = dllFiles.Concat(exeFiles).ToArray();

            var totalFiles = allFiles.Length;
            ResultCount = totalFiles;

            // ファイルが見つからない場合
            if (totalFiles == 0)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusMessage = "ファイルが見つかりませんでした。";
                });
                return;
            }

            var results = new List<AssemblyInfo>();
            int processed = 0;

            // 各ファイルを解析
            foreach (var file in allFiles)
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    var filename = fileInfo.Name;

                    // 進捗更新
                    processed++;
                    var progressPercent = (int)((double)processed / totalFiles * 100);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ProgressValue = progressPercent;
                        ProgressDetail = $"{processed}/{totalFiles} - {filename}";
                        StatusMessage = $"解析中... ({processed}/{totalFiles})";
                    });

                    // アセンブリ情報を取得
                    var assemblyInfo = GetAssemblyInfo(file);

                    // 結果を追加
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        results.Add(assemblyInfo);
                        _assemblyInfos.Add(assemblyInfo);

                        if (assemblyInfo.IsAnalyzed)
                            SuccessCount++;
                        else
                            FailedCount++;
                    });
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ファイル処理エラー: {ex.Message}");
                }
            }

            // CSV出力
            WriteResultsToCsv(results, OutputFilePath);
        }

        private AssemblyInfo GetAssemblyInfo(string filePath)
        {
            try
            {
                // アセンブリをロード
                var assembly = Assembly.LoadFile(filePath);

                // CLRバージョンを取得
                var runtimeVersion = assembly.ImageRuntimeVersion;

                // .NET Frameworkバージョンを判定
                var frameworkVersion = GetFrameworkVersion(assembly);

                // アセンブリ名を取得
                var assemblyName = assembly.GetName();

                return new AssemblyInfo
                {
                    FileName = Path.GetFileName(filePath),
                    FilePath = filePath,
                    RuntimeVersion = runtimeVersion,
                    FrameworkVersion = frameworkVersion,
                    FullName = assembly.FullName,
                    ProcessorArchitecture = GetProcessorArchitecture(assemblyName.ProcessorArchitecture),
                    IsAnalyzed = true,
                    ErrorMessage = string.Empty
                };
            }
            catch (Exception ex)
            {
                // エラーが発生した場合
                return new AssemblyInfo
                {
                    FileName = Path.GetFileName(filePath),
                    FilePath = filePath,
                    RuntimeVersion = string.Empty,
                    FrameworkVersion = string.Empty,
                    FullName = string.Empty,
                    ProcessorArchitecture = string.Empty,
                    IsAnalyzed = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private string GetFrameworkVersion(Assembly assembly)
        {
            try
            {
                var runtimeVersion = assembly.ImageRuntimeVersion;

                // TargetFrameworkAttributeを取得
                var targetFrameworkAttribute = assembly.GetCustomAttribute<TargetFrameworkAttribute>();

                // .NET Frameworkバージョンを判定
                if (runtimeVersion.StartsWith("v1.0."))
                    return ".NET Framework 1.0";
                else if (runtimeVersion.StartsWith("v1.1."))
                    return ".NET Framework 1.1";
                else if (runtimeVersion.StartsWith("v2.0."))
                    return ".NET Framework 2.0-3.5";
                else if (runtimeVersion.StartsWith("v4.0."))
                {
                    // .NET Framework 4.0以降
                    if (targetFrameworkAttribute != null)
                    {
                        string frameworkName = targetFrameworkAttribute.FrameworkName;

                        if (frameworkName.Contains(".NETCoreApp,"))
                            return ".NET Core / .NET 5+";
                        else if (frameworkName.Contains(".NETStandard,"))
                            return ".NET Standard";
                        else if (frameworkName.Contains(".NETFramework,"))
                        {
                            // バージョン番号を抽出
                            if (frameworkName.Contains("Version=v4.0"))
                                return ".NET Framework 4.0";
                            else if (frameworkName.Contains("Version=v4.5"))
                                return ".NET Framework 4.5";
                            else if (frameworkName.Contains("Version=v4.5.1"))
                                return ".NET Framework 4.5.1";
                            else if (frameworkName.Contains("Version=v4.5.2"))
                                return ".NET Framework 4.5.2";
                            else if (frameworkName.Contains("Version=v4.6"))
                                return ".NET Framework 4.6";
                            else if (frameworkName.Contains("Version=v4.6.1"))
                                return ".NET Framework 4.6.1";
                            else if (frameworkName.Contains("Version=v4.6.2"))
                                return ".NET Framework 4.6.2";
                            else if (frameworkName.Contains("Version=v4.7"))
                                return ".NET Framework 4.7";
                            else if (frameworkName.Contains("Version=v4.7.1"))
                                return ".NET Framework 4.7.1";
                            else if (frameworkName.Contains("Version=v4.7.2"))
                                return ".NET Framework 4.7.2";
                            else if (frameworkName.Contains("Version=v4.8"))
                                return ".NET Framework 4.8";
                            else if (frameworkName.Contains("Version=v4.8.1"))
                                return ".NET Framework 4.8.1";
                            else
                                return "不明 (.NET Framework 4.x)";
                        }
                        return "不明 (v4.0+)";
                    }
                    return ".NET Framework 4.0-4.8";
                }
                return "不明";
            }
            catch
            {
                return "取得できませんでした";
            }
        }

        private string GetProcessorArchitecture(ProcessorArchitecture architecture)
        {
            switch (architecture)
            {
                case ProcessorArchitecture.MSIL:
                    return "MSIL (Any CPU)";
                case ProcessorArchitecture.X86:
                    return "X86 (32-bit)";
                case ProcessorArchitecture.IA64:
                    return "IA64 (Itanium)";
                case ProcessorArchitecture.Amd64:
                    return "AMD64 (64-bit)";
                case ProcessorArchitecture.Arm:
                    return "ARM";
                default:
                    return "不明";
            }
        }

        private void WriteResultsToCsv(List<AssemblyInfo> results, string outputPath)
        {
            // CSV出力用のディレクトリがない場合は作成
            var directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // CSVにヘッダーと結果を書き込み
            using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // ヘッダー
                writer.WriteLine("ファイル名,ファイルパス,CLRバージョン,.NETバージョン,アセンブリ名,プロセッサアーキテクチャ,解析成功,エラーメッセージ");

                // データ行
                foreach (var info in results)
                {
                    writer.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\"",
                        EscapeCsvField(info.FileName),
                        EscapeCsvField(info.FilePath),
                        EscapeCsvField(info.RuntimeVersion),
                        EscapeCsvField(info.FrameworkVersion),
                        EscapeCsvField(info.FullName),
                        EscapeCsvField(info.ProcessorArchitecture),
                        info.IsAnalyzed,
                        EscapeCsvField(info.ErrorMessage)));
                }
            }
        }

        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            // ダブルクォートをエスケープ
            return field.Replace("\"", "\"\"");
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class AssemblyInfo
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string RuntimeVersion { get; set; }
        public string FrameworkVersion { get; set; }
        public string FullName { get; set; }
        public string ProcessorArchitecture { get; set; }
        public bool IsAnalyzed { get; set; }
        public string ErrorMessage { get; set; }
    }
}