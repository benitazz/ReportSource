#region

using System.Collections.ObjectModel;
using System.Windows;

using ExcelReportsGenerator.Models;
using ExcelReportsGenerator.ViewModels;
using ExcelReportsGenerator.Views;

#endregion

namespace ExcelReportsGenerator
{
  /// <summary>
  ///   Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    #region Constructors and Destructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="MainWindow" /> class.
    /// </summary>
    public MainWindow()
    {
      this.InitializeComponent();

      this.Loaded += this.MainWindowLoaded;
    }

    #endregion

    #region Methods

    /// <summary>
    /// The main window loaded.
    /// </summary>
    /// <param name="sender">
    /// The sender.
    /// </param>
    /// <param name="e">
    /// The e.
    /// </param>
    private void MainWindowLoaded(object sender, RoutedEventArgs e)
    {
      var shellContent = new TabControlModel
                           {
                             TabContent = new ExcelReportViewModel(),
                             HeaderName = "Excel Reports",
                             IsActive = true,
                             ImageSource = "/ExcelReportsGenerator;component/Resources/Images/ExelIcon.jpg",
                             DllKey = "ExcelReportsGenerator.ViewModels.ExcelReportViewModel"
                           };

      var hostViewModel = new ShellViewModel
                            {
                              TabControlsObservableCollection =
                                new ObservableCollection<TabControlModel>
                                  {
                                    shellContent,
                                  },
                                SelectedTabControl = shellContent
                            };
      
     this.MainViewContent.Content = new ShellView { DataContext = hostViewModel };
    }

    #endregion
  }
}